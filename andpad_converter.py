"""ANDPAD工務予算振り分けモジュール - 積算原価から材工分離を行う

『積算原価⇒工務予算振り分け方法 (1).xlsx』（2026年7月版）に準拠。
- 1明細を複数の「配分」（発注先 × 材/工 × 金額）に分解する。
  材工分離した施工費は、大工手間・造作建具（フジタ）・左官・電気工事等の
  「加算されるべき発注先」の集計に加算される。
- 発注先（細目工種）の名称・並び順は予実管理表（別表）に合わせる。
"""
from dataclasses import dataclass, field
from estimate_calculator import Estimate, EstimateItem, WorkCategory


# 発注先（細目工種）の並び順 - 予実管理表（別表）の掲載順
VENDOR_ORDER = [
    # 仮設工事
    "駐車場代", "基礎廻り養生シート", "仮設工事", "整地費用",
    "仮設足場", "産廃", "運搬費",
    # 基礎工事
    "基礎工事", "仮設砕石",
    # 木工事
    "プレカット", "金物", "置き家具",
    "建材（土橋）", "大工手間",
    "建材（特殊床材・框）", "建材（階段）北恵", "FB手摺（北恵）", "FB手摺",
    "階段手摺金具（カワジュン）",
    "インテリア建具", "洋風造作材",
    "シンホリ（ダイライト）", "床養生", "補修工事", "MIRAIE", "レッカー",
    "建材（標準床材）", "建材（特殊床材）北恵", "ジツダヤ",
    # 断熱工事
    "基礎断熱", "断熱材・気密シート", "気密手間",
    # 屋根工事
    "瓦屋根", "鋼板屋根",
    # 板金・樋工事
    "板金工事",
    # 防水工事
    "シンホリ",
    # 外部建具工事
    "サッシ",
    # 内部建具工事
    "可動棚（フィットラック）", "造作建具（フジタ）",
    "インテリア建具（LIXIL）", "HIKARI", "神谷コーポレーション",
    "インテリア建具（永大）",
    # 金属工事
    "アルミ笠木", "バルコニー手摺", "アルミ格子", "太田商事",
    # 外装工事
    "サイディング", "外壁板金",
    # 左官工事
    "CTバリヤ", "左官",
    # タイル工事
    "タイル建材", "タイル工事",
    # デッキ工事
    "デッキ工事",
    # 塗装工事
    "内部塗装", "外部塗装",
    # 内装工事
    "畳", "クロス・CF",
    # 住宅設備機器工事
    "キッチン", "タカギ水栓", "カップボード", "システムバス",
    "トイレ", "トイレ手摺", "造作洗面", "アクセサリー",
    "洗面化粧台（LIXIL）", "洗濯パン", "草川工業",
    "ガス給湯器", "エコキュート", "コスティム",
    # 給排水設備工事
    "内部給排水",
    # 電気設備工事
    "仮設電気", "仮設電気料金（中電）", "電気工事",
    "火災警報器", "インターホン", "HEMS",
    # 換気システム工事
    "パナソニックリビング", "24H換気施工", "屋外フード", "スリーブ工事",
    # 空調設備工事
    "エアコン工事",
    # 太陽光発電システム工事
    "太陽光",
    # 造作家具工事
    "造作家具",
    # 外構工事
    "外構工事",
    # 雑工事
    "クリーニング", "ホスクリーン", "予備費",
    # 屋外付帯工事
    "地盤改良工事", "外部給排水", "配管パック", "ガス費用", "ガス工事",
    "浄化槽", "解体工事",
    # 内部付帯工事
    "照明工事", "カーテン工事",
    # 諸費用
    "コーディネート費用", "設計管理費",
    "瑕疵担保責任保険", "二次防水", "地盤調査", "現場調査費",
    "設備保証", "地震保証",
    "各種キャンペーン", "管理費",
]
_VENDOR_ORDER_INDEX = {name: i for i, name in enumerate(VENDOR_ORDER)}


def vendor_sort_key(vendor: str):
    """発注先を予実管理表の並び順でソートするためのキー"""
    idx = _VENDOR_ORDER_INDEX.get(vendor)
    if idx is not None:
        return (0, idx, vendor)
    return (1, 0, vendor)


# 材工分離の按分先（三分割時等の共通先）
_DAIKU = "大工手間"
_CLOTH = "クロス・CF"

# 『クロス』『大工手間』『建材』で3等分（ニッチ・垂れ壁等）
def _three_way_parts():
    return [
        {"vendor": _CLOTH, "kind": "工", "ratio": 1 / 3, "note": "クロス分"},
        {"vendor": _DAIKU, "kind": "工", "ratio": 1 / 3, "note": "大工手間分"},
    ]


# 細目工種（ANDPAD発注先）マッピング
#   vendor / kind : 残額（本体）の発注先と区分（材 or 工）
#   parts         : 先取りで他の発注先へ振り分ける配分
#                   ppu=数量あたり定額, ratio=金額比率
VENDOR_MAPPING = {
    # ===== 仮設工事 =====
    10001: {"vendor": "仮設足場", "kind": "工"},
    10002: {"vendor": "仮設工事", "kind": "工"},
    10003: {"vendor": "産廃", "kind": "工"},
    10004: {"vendor": "運搬費", "kind": "工"},
    10005: {"vendor": "駐車場代", "kind": "工"},
    10006: {"vendor": "基礎廻り養生シート", "kind": "工"},
    10007: {"vendor": "仮設工事", "kind": "工"},
    10008: {"vendor": "仮設工事", "kind": "工"},
    10009: {"vendor": "仮設工事", "kind": "工"},
    10010: {"vendor": "整地費用", "kind": "工"},

    # ===== 基礎工事 =====
    20001: {"vendor": "基礎工事", "kind": "工"},
    20002: {"vendor": "基礎工事", "kind": "工"},
    20003: {"vendor": "基礎工事", "kind": "工"},
    20004: {"vendor": "基礎工事", "kind": "工"},
    20005: {"vendor": "基礎工事", "kind": "工"},
    20006: {"vendor": "基礎工事", "kind": "工"},
    20007: {"vendor": "基礎工事", "kind": "工"},
    20008: {"vendor": "基礎工事", "kind": "工"},

    # ===== 木工事 - 床材（施工費は大工手間へ加算） =====
    # 30001は北恵の無垢材（※施工費込）のため建材（特殊床材）北恵
    30001: {"vendor": "建材（特殊床材）北恵", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ratio": 0.3, "note": "施工費分"}]},
    30002: {"vendor": "建材（標準床材）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ratio": 0.3, "note": "施工費分"}]},
    30003: {"vendor": "建材（標準床材）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ratio": 0.3, "note": "施工費分"}]},
    # 水廻り床: 『ジツダヤ：10,000円』『大工手間：4,000円』（坪あたり）
    30004: {"vendor": "ジツダヤ", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ppu": 4000, "note": "施工費分"}]},

    # ===== 木工事 - プレカット =====
    30005: {"vendor": "プレカット", "kind": "材"},
    30006: {"vendor": "プレカット", "kind": "材"},
    30007: {"vendor": "プレカット", "kind": "材"},
    30008: {"vendor": "プレカット", "kind": "材"},
    30009: {"vendor": "プレカット", "kind": "材"},
    30010: {"vendor": "金物", "kind": "材"},
    30011: {"vendor": "プレカット", "kind": "材"},
    30012: {"vendor": "プレカット", "kind": "材"},
    30013: {"vendor": "プレカット", "kind": "材"},

    # ===== 木工事 - 建材 =====
    30014: {"vendor": "建材（土橋）", "kind": "材"},
    30015: {"vendor": "建材（土橋）", "kind": "材"},
    30016: {"vendor": "建材（土橋）", "kind": "材"},
    30017: {"vendor": "建材（土橋）", "kind": "材"},
    30018: {"vendor": "建材（特殊床材・框）", "kind": "材"},
    30019: {"vendor": "建材（土橋）", "kind": "材"},   # 玄関框(挽板･無垢床用)は建材（土橋）
    30020: {"vendor": "建材（土橋）", "kind": "材"},
    30021: {"vendor": "建材（土橋）", "kind": "材"},
    30022: {"vendor": "階段手摺金具（カワジュン）", "kind": "材"},
    30023: {"vendor": "建材（土橋）", "kind": "材"},
    30024: {"vendor": "建材（土橋）", "kind": "材"},
    30025: {"vendor": "建材（土橋）", "kind": "材"},
    30026: {"vendor": "建材（土橋）", "kind": "材"},
    30027: {"vendor": "建材（土橋）", "kind": "材"},
    30028: {"vendor": "建材（土橋）", "kind": "材"},
    # 下地補強 小物: 『建材』と『大工手間』で半々
    30029: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ratio": 0.5, "note": "施工費分"}]},
    30030: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    30031: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    30032: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    30033: {"vendor": "MIRAIE", "kind": "材"},
    30034: {"vendor": "インテリア建具", "kind": "材"},
    30035: {"vendor": "インテリア建具", "kind": "材"},
    # ニッチ: 『クロス』『大工手間』『建材』で3等分
    30036: {"vendor": "建材（土橋）", "kind": "材", "parts": _three_way_parts(),
            "remainder_note": "建材分"},
    30037: {"vendor": "金物", "kind": "材"},
    30038: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    30039: {"vendor": "建材（土橋）", "kind": "材"},
    30040: {"vendor": "内部塗装", "kind": "工"},
    # アーチ垂れ壁: 3等分
    30041: {"vendor": "建材（土橋）", "kind": "材", "parts": _three_way_parts(),
            "remainder_note": "建材分"},
    30042: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    30043: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    30044: {"vendor": "床養生", "kind": "工"},
    30045: {"vendor": "補修工事", "kind": "工"},
    30046: {"vendor": "建材（土橋）", "kind": "材"},
    # 垂れ壁 水平: 3等分
    30047: {"vendor": "建材（土橋）", "kind": "材", "parts": _three_way_parts(),
            "remainder_note": "建材分"},
    30048: {"vendor": "建材（土橋）", "kind": "材"},
    30049: {"vendor": "建材（土橋）", "kind": "材"},
    # 縦格子: 『大工手間１箇所：15,000円』『建材：残り』
    30050: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ppu": 15000, "note": "施工費分"}]},
    # ししまるくんヌックスペース造作追加: 施工費は大工手間へ、材料費は建材（土橋）へ
    30051: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ratio": 0.5, "note": "施工費分"}]},
    30052: {"vendor": "建材（土橋）", "kind": "材"},
    30053: {"vendor": _DAIKU, "kind": "工"},
    # 天井埋込カーテンBOX造作: 3等分
    30054: {"vendor": "建材（土橋）", "kind": "材", "parts": _three_way_parts(),
            "remainder_note": "建材分"},
    # トイレ壁ニッチ: 3等分
    30055: {"vendor": "建材（土橋）", "kind": "材", "parts": _three_way_parts(),
            "remainder_note": "建材分"},
    30056: {"vendor": "建材（土橋）", "kind": "材"},
    # 電気設備用点検口: 『大工手間：（１箇所）6,000円』『建材：残り』
    30057: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ppu": 6000, "note": "取付手間分"}]},
    30058: {"vendor": "建材（土橋）", "kind": "材"},
    30059: {"vendor": "エアコン工事", "kind": "工"},
    30060: {"vendor": "エアコン工事", "kind": "工"},

    # ===== 木工事 - 大工手間 =====
    30101: {"vendor": _DAIKU, "kind": "工"},
    30102: {"vendor": _DAIKU, "kind": "工"},
    30103: {"vendor": _DAIKU, "kind": "工"},
    30104: {"vendor": _DAIKU, "kind": "工"},
    30105: {"vendor": _DAIKU, "kind": "工"},
    30106: {"vendor": _DAIKU, "kind": "工"},
    30107: {"vendor": _DAIKU, "kind": "工"},
    30108: {"vendor": _DAIKU, "kind": "工"},
    30109: {"vendor": _DAIKU, "kind": "工"},
    30110: {"vendor": _DAIKU, "kind": "工"},
    30111: {"vendor": _DAIKU, "kind": "工"},
    30112: {"vendor": _DAIKU, "kind": "工"},
    30113: {"vendor": _DAIKU, "kind": "工"},
    30114: {"vendor": _DAIKU, "kind": "工"},
    30115: {"vendor": _DAIKU, "kind": "工"},
    30116: {"vendor": _DAIKU, "kind": "工"},
    30117: {"vendor": _DAIKU, "kind": "工"},
    30118: {"vendor": _DAIKU, "kind": "工"},
    # 壁ふかし工事: 『大工手間：1/2』『建材：1/2』
    30119: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": _DAIKU, "kind": "工", "ratio": 0.5, "note": "施工費分"}]},
    30120: {"vendor": "レッカー", "kind": "工"},
    30128: {"vendor": _DAIKU, "kind": "工"},
    30129: {"vendor": _DAIKU, "kind": "工"},

    # ===== 含み予算（諸費用扱い・発注先は個別名称） =====
    30121: {"vendor": "瑕疵担保責任保険", "kind": "材"},
    30122: {"vendor": "瑕疵担保責任保険", "kind": "材"},
    30123: {"vendor": "二次防水", "kind": "材"},
    30124: {"vendor": "地盤調査", "kind": "材"},
    30125: {"vendor": "設備保証", "kind": "材"},
    30126: {"vendor": "地震保証", "kind": "材"},
    30127: {"vendor": "現場調査費", "kind": "材"},

    # ===== 断熱工事 =====
    40001: {"vendor": "断熱材・気密シート", "kind": "材"},
    40002: {"vendor": "断熱材・気密シート", "kind": "材"},
    40003: {"vendor": "断熱材・気密シート", "kind": "材"},
    40004: {"vendor": "基礎断熱", "kind": "材"},
    40005: {"vendor": "基礎断熱", "kind": "材"},
    40006: {"vendor": "基礎断熱", "kind": "材"},
    40007: {"vendor": "基礎断熱", "kind": "材"},
    40008: {"vendor": "断熱材・気密シート", "kind": "材"},
    40009: {"vendor": "シンホリ（ダイライト）", "kind": "材"},
    40010: {"vendor": "気密手間", "kind": "工"},
    40011: {"vendor": "気密手間", "kind": "工"},
    40012: {"vendor": "気密手間", "kind": "工"},

    # ===== 屋根工事 =====
    50001: {"vendor": "鋼板屋根", "kind": "工"},
    50002: {"vendor": "瓦屋根", "kind": "工"},
    50003: {"vendor": "鋼板屋根", "kind": "工"},
    50004: {"vendor": "鋼板屋根", "kind": "工"},
    50005: {"vendor": "瓦屋根", "kind": "工"},

    # ===== 板金・樋工事 =====
    60001: {"vendor": "板金工事", "kind": "工"},
    60002: {"vendor": "板金工事", "kind": "工"},
    60003: {"vendor": "板金工事", "kind": "工"},

    # ===== 外部建具工事 =====
    70001: {"vendor": "サッシ", "kind": "材"},
    70002: {"vendor": "サッシ", "kind": "材"},
    70003: {"vendor": "サッシ", "kind": "材"},
    70004: {"vendor": "サッシ", "kind": "材"},
    70005: {"vendor": "サッシ", "kind": "材"},
    70006: {"vendor": "サッシ", "kind": "材"},
    70007: {"vendor": "サッシ", "kind": "材"},
    70008: {"vendor": "サッシ", "kind": "材"},
    70009: {"vendor": "サッシ", "kind": "材"},
    70010: {"vendor": "サッシ", "kind": "材"},
    70011: {"vendor": "サッシ", "kind": "材"},
    70012: {"vendor": "サッシ", "kind": "材"},
    70013: {"vendor": "サッシ", "kind": "材"},
    70014: {"vendor": "サッシ", "kind": "材"},
    70015: {"vendor": "サッシ", "kind": "材"},
    70016: {"vendor": "サッシ", "kind": "工"},
    70017: {"vendor": "サッシ", "kind": "材"},

    # ===== 内部建具工事 =====
    # 可動棚（D450汎用品）: 材＝建材（土橋）残り、
    # 工＝造作建具（フジタ）に手間8,000円＋金具12,000円＝20,000円/ヶ所
    80001: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": "造作建具（フジタ）", "kind": "工", "ppu": 20000,
                       "note": "取付手間8,000＋金具12,000/ヶ所"}]},
    80002: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": "造作建具（フジタ）", "kind": "工", "ppu": 20000,
                       "note": "取付手間8,000＋金具12,000/ヶ所"}]},
    80003: {"vendor": "建材（土橋）", "kind": "材",
            "parts": [{"vendor": "造作建具（フジタ）", "kind": "工", "ppu": 20000,
                       "note": "取付手間8,000＋金具12,000/ヶ所"}]},
    # 可動棚（フィットラック品）: 材＝可動棚（フィットラック）残り、
    # 工＝造作建具（フジタ）8,000円/ヶ所
    80004: {"vendor": "可動棚（フィットラック）", "kind": "材",
            "parts": [{"vendor": "造作建具（フジタ）", "kind": "工", "ppu": 8000, "note": "取付手間分"}]},
    80005: {"vendor": "可動棚（フィットラック）", "kind": "材",
            "parts": [{"vendor": "造作建具（フジタ）", "kind": "工", "ppu": 8000, "note": "取付手間分"}]},
    80006: {"vendor": "可動棚（フィットラック）", "kind": "材",
            "parts": [{"vendor": "造作建具（フジタ）", "kind": "工", "ppu": 8000, "note": "取付手間分"}]},
    80007: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80008: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80009: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80010: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80011: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80012: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80013: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80014: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80015: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80016: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80017: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80018: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80019: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80020: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80021: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    80022: {"vendor": "インテリア建具（LIXIL）", "kind": "材"},
    # ししまるくんヌックスペース収納棚: 造作家具のまま（材工分けない）
    80023: {"vendor": "造作家具", "kind": "材"},
    80024: {"vendor": "造作家具", "kind": "材"},

    # ===== 外装工事 =====
    90001: {"vendor": "サイディング", "kind": "工"},
    90002: {"vendor": "板金工事", "kind": "工"},   # 破風・鼻隠し SGL鋼板
    90003: {"vendor": "サイディング", "kind": "工"},
    90004: {"vendor": "サイディング", "kind": "工"},
    90005: {"vendor": "サイディング", "kind": "工"},
    90006: {"vendor": "サイディング", "kind": "工"},

    # ===== 左官工事 =====
    100001: {"vendor": "左官", "kind": "工"},
    100002: {"vendor": "左官", "kind": "工"},
    100003: {"vendor": "CTバリヤ", "kind": "材"},
    100004: {"vendor": "CTバリヤ", "kind": "材"},
    100005: {"vendor": "CTバリヤ", "kind": "材"},
    100006: {"vendor": "左官", "kind": "工"},

    # ===== タイル工事 =====
    # 玄関・ポーチ タイル貼り: 『左官：㎡＝10,000円』『タイル建材：残り』
    110001: {"vendor": "タイル建材", "kind": "材",
             "parts": [{"vendor": "左官", "kind": "工", "ppu": 10000, "note": "施工費分"}]},

    # ===== 塗装工事 =====
    120001: {"vendor": "内部塗装", "kind": "工"},

    # ===== 内装工事 =====
    130001: {"vendor": "畳", "kind": "材"},
    130002: {"vendor": "クロス・CF", "kind": "工"},
    130003: {"vendor": "クロス・CF", "kind": "工"},
    130004: {"vendor": "クロス・CF", "kind": "工"},
    130005: {"vendor": "クロス・CF", "kind": "工"},
    130006: {"vendor": "クロス・CF", "kind": "工"},

    # ===== 住宅設備機器工事（発注先＝設備名） =====
    140001: {"vendor": "キッチン", "kind": "材"},
    140002: {"vendor": "キッチン", "kind": "材"},
    140003: {"vendor": "キッチン", "kind": "材"},
    140004: {"vendor": "カップボード", "kind": "材"},
    140005: {"vendor": "タカギ水栓", "kind": "材"},
    140006: {"vendor": "システムバス", "kind": "材"},
    140007: {"vendor": "システムバス", "kind": "材"},
    140008: {"vendor": "トイレ", "kind": "材"},
    140009: {"vendor": "トイレ", "kind": "材"},
    140010: {"vendor": "トイレ", "kind": "材"},
    140011: {"vendor": "造作洗面", "kind": "材"},
    140012: {"vendor": "造作洗面", "kind": "材"},
    140013: {"vendor": "アクセサリー", "kind": "材"},
    140014: {"vendor": "アクセサリー", "kind": "材"},
    140015: {"vendor": "洗濯パン", "kind": "材"},
    # エコキュート: ①架台（左官）35,000円 ②電気工事30,000円 ③エコキュート:残り
    140016: {"vendor": "エコキュート", "kind": "材", "remainder_note": "本体",
             "parts": [
                 {"vendor": "左官", "kind": "工", "ppu": 35000, "note": "架台分"},
                 {"vendor": "電気工事", "kind": "工", "ppu": 30000, "note": "電源工事分"},
             ]},
    140017: {"vendor": "エコキュート", "kind": "材", "remainder_note": "本体",
             "parts": [
                 {"vendor": "左官", "kind": "工", "ppu": 35000, "note": "架台分"},
                 {"vendor": "電気工事", "kind": "工", "ppu": 30000, "note": "電源工事分"},
             ]},
    140018: {"vendor": "トイレ", "kind": "材"},

    # ===== 給排水設備工事 =====
    150001: {"vendor": "内部給排水", "kind": "工"},
    150002: {"vendor": "内部給排水", "kind": "工"},
    150003: {"vendor": "内部給排水", "kind": "工"},

    # ===== 電気設備工事 =====
    # 仮設電気: ①仮設電気（電柱設置）20,000円 ②仮設電気料金（中電）25,000円
    160001: {"vendor": "仮設電気料金（中電）", "kind": "工", "remainder_note": "電気料金分",
             "parts": [{"vendor": "仮設電気", "kind": "工", "ppu": 20000, "note": "電柱設置分"}]},
    160002: {"vendor": "電気工事", "kind": "工"},
    160003: {"vendor": "電気工事", "kind": "工"},
    160004: {"vendor": "電気工事", "kind": "工"},
    160005: {"vendor": "電気工事", "kind": "工"},
    160006: {"vendor": "電気工事", "kind": "工"},
    160007: {"vendor": "電気工事", "kind": "工"},
    160008: {"vendor": "電気工事", "kind": "工"},
    160009: {"vendor": "電気工事", "kind": "工"},
    160010: {"vendor": "電気工事", "kind": "工"},
    160011: {"vendor": "電気工事", "kind": "工"},
    160012: {"vendor": "電気工事", "kind": "工"},
    160013: {"vendor": "電気工事", "kind": "工"},
    160014: {"vendor": "電気工事", "kind": "工"},
    160015: {"vendor": "電気工事", "kind": "工"},
    160016: {"vendor": "火災警報器", "kind": "材"},   # 本体（パナソニックリビング）
    160017: {"vendor": "電気工事", "kind": "工"},      # 取付費
    160018: {"vendor": "電気工事", "kind": "材"},
    160019: {"vendor": "スリーブ工事", "kind": "工"},  # コスティム
    160020: {"vendor": "電気工事", "kind": "材"},
    160021: {"vendor": "インターホン", "kind": "材"},
    160022: {"vendor": "電気工事", "kind": "工"},
    160023: {"vendor": "電気工事", "kind": "工"},
    160024: {"vendor": "電気工事", "kind": "工"},
    160025: {"vendor": "電気工事", "kind": "工"},
    160026: {"vendor": "HEMS", "kind": "材"},
    160027: {"vendor": "HEMS", "kind": "工"},
    160028: {"vendor": "電気工事", "kind": "工"},

    # ===== 換気システム工事 =====
    170001: {"vendor": "パナソニックリビング", "kind": "材"},
    170002: {"vendor": "パナソニックリビング", "kind": "材"},
    170003: {"vendor": "パナソニックリビング", "kind": "材"},
    170004: {"vendor": "パナソニックリビング", "kind": "材"},
    170005: {"vendor": "パナソニックリビング", "kind": "材"},
    170006: {"vendor": "パナソニックリビング", "kind": "材"},
    170007: {"vendor": "24H換気施工", "kind": "工"},
    170008: {"vendor": "電気工事", "kind": "工"},
    170009: {"vendor": "パナソニックリビング", "kind": "材"},
    170010: {"vendor": "電気工事", "kind": "工"},
    170011: {"vendor": "電気工事", "kind": "工"},
    170012: {"vendor": "電気工事", "kind": "工"},
    170013: {"vendor": "屋外フード", "kind": "材"},   # 更科製作所
    170014: {"vendor": "屋外フード", "kind": "材"},
    170015: {"vendor": "屋外フード", "kind": "材"},
    170016: {"vendor": "屋外フード", "kind": "材"},

    # ===== 空調設備工事（材工振り分けしない・発注先＝エアコン工事(ゼネラル)） =====
    180001: {"vendor": "エアコン工事", "kind": "工"},
    180002: {"vendor": "エアコン工事", "kind": "工"},
    180003: {"vendor": "エアコン工事", "kind": "工"},
    180004: {"vendor": "エアコン工事", "kind": "工"},
    180005: {"vendor": "エアコン工事", "kind": "材"},
    180006: {"vendor": "エアコン工事", "kind": "材"},
    180007: {"vendor": "エアコン工事", "kind": "工"},
    180008: {"vendor": "エアコン工事", "kind": "工"},
    180009: {"vendor": "エアコン工事", "kind": "工"},

    # ===== 太陽光発電システム工事（材工振り分けしない） =====
    190001: {"vendor": "太陽光", "kind": "材"},

    # ===== 造作家具工事（仕様変更があるため材工分けない） =====
    200001: {"vendor": "造作家具", "kind": "材"},
    200002: {"vendor": "造作家具", "kind": "材"},
    200003: {"vendor": "造作家具", "kind": "材"},
    200004: {"vendor": "造作家具", "kind": "材"},

    # ===== 雑工事 =====
    210001: {"vendor": "クリーニング", "kind": "工"},
    # 点検口: 『大工手間：（１箇所）6,000円』『建材：残り』
    210002: {"vendor": "建材（土橋）", "kind": "材",
             "parts": [{"vendor": _DAIKU, "kind": "工", "ppu": 6000, "note": "取付手間分"}]},
    210003: {"vendor": "建材（土橋）", "kind": "材",
             "parts": [{"vendor": _DAIKU, "kind": "工", "ppu": 6000, "note": "取付手間分"}]},
    # 室内物干し金物 KACU: シンホリへ（材工分けない）
    210004: {"vendor": "シンホリ", "kind": "材"},
    # ホスクリーン: 『大工手間：（１箇所）3,000円』『ホスクリーン：残り』
    210005: {"vendor": "ホスクリーン", "kind": "材",
             "parts": [{"vendor": _DAIKU, "kind": "工", "ppu": 3000, "note": "取付手間分"}]},

    # ===== 管理費 =====
    1: {"vendor": "管理費", "kind": "材"},

    # ===== 屋外付帯工事 =====
    220001: {"vendor": "外部給排水", "kind": "工"},
    220002: {"vendor": "仮設工事", "kind": "工"},
    220003: {"vendor": "ガス工事", "kind": "工"},

    # ===== 諸費用 =====
    230001: {"vendor": "コーディネート費用", "kind": "工"},
    230002: {"vendor": "設計管理費", "kind": "工"},
}


@dataclass
class ANDPADAllocation:
    """発注先への配分（材工分離後の1配分）"""
    vendor: str    # 発注先（細目工種）
    kind: str      # '材' or '工'
    amount: int    # 金額
    note: str = ""  # 補足（例: 施工費分・架台分）


@dataclass
class ANDPADItem:
    """ANDPAD発注用項目（1積算明細 = 1項目、内部に複数配分を持つ）"""
    work_category: str      # 工事場所（大項目）
    item_name: str          # 名称
    summary: str            # 摘要
    quantity: float = 0.0
    unit: str = ""
    allocations: list = field(default_factory=list)
    # 修正保存用の安定キー（工事番号_明細コード_同一コード内の出現順）。
    # 明細の増減で通し番号がずれても、修正が別の明細に誤適用されないようにする。
    stable_key: str = ""

    @property
    def vendor(self) -> str:
        """主発注先（材の配分を優先）"""
        for a in self.allocations:
            if a.kind == "材":
                return a.vendor
        return self.allocations[0].vendor if self.allocations else ""

    @property
    def material_cost(self) -> int:
        return sum(a.amount for a in self.allocations if a.kind == "材")

    @property
    def labor_cost(self) -> int:
        return sum(a.amount for a in self.allocations if a.kind == "工")

    @property
    def total_cost(self) -> int:
        return sum(a.amount for a in self.allocations)


@dataclass
class ANDPADBudget:
    """ANDPAD工務予算"""
    items: list = field(default_factory=list)

    def by_vendor(self) -> dict:
        """発注先別に集計（材工分離後の配分ベース・予実管理表の並び順）"""
        result = {}
        for item in self.items:
            for alloc in item.allocations:
                if alloc.vendor not in result:
                    result[alloc.vendor] = {
                        'material': 0,
                        'labor': 0,
                        'total': 0,
                        'details': [],  # (item, alloc) のリスト
                    }
                d = result[alloc.vendor]
                if alloc.kind == "材":
                    d['material'] += alloc.amount
                else:
                    d['labor'] += alloc.amount
                d['total'] += alloc.amount
                d['details'].append((item, alloc))
        # 予実管理表の並び順でソート
        return dict(sorted(result.items(), key=lambda x: vendor_sort_key(x[0])))

    def by_work_category(self) -> dict:
        """工事種別に集計"""
        result = {}
        for item in self.items:
            cat = item.work_category
            if cat not in result:
                result[cat] = {
                    'material': 0,
                    'labor': 0,
                    'total': 0,
                    'items': [],
                }
            result[cat]['material'] += item.material_cost
            result[cat]['labor'] += item.labor_cost
            result[cat]['total'] += item.total_cost
            result[cat]['items'].append(item)
        return result

    @property
    def total_material(self) -> int:
        return sum(i.material_cost for i in self.items)

    @property
    def total_labor(self) -> int:
        return sum(i.labor_cost for i in self.items)

    @property
    def grand_total(self) -> int:
        return sum(i.total_cost for i in self.items)


def _build_allocations(mapping: dict, item: EstimateItem) -> list:
    """マッピング定義から配分リストを構築する。

    parts（先取り配分）を計算し、残額を主発注先（vendor/kind）に割り当てる。
    配分合計は必ず発注金額と一致する（残額方式）。
    """
    total = item.order_amount
    allocs = []
    remaining = total

    for part in mapping.get('parts', []):
        if 'ppu' in part:
            amt = int(round(part['ppu'] * item.quantity))
        elif 'ratio' in part:
            amt = int(round(total * part['ratio']))
        else:
            amt = int(part.get('fixed', 0))
        # マイナス明細や残額超過では先取りしない（安全側）
        if total > 0:
            amt = max(0, min(amt, remaining))
        else:
            amt = 0
        if amt != 0:
            allocs.append(ANDPADAllocation(
                vendor=part['vendor'], kind=part['kind'],
                amount=amt, note=part.get('note', ''),
            ))
            remaining -= amt

    if remaining != 0 or not allocs:
        note = mapping.get('remainder_note', '') if allocs else ''
        allocs.insert(0, ANDPADAllocation(
            vendor=mapping['vendor'], kind=mapping['kind'],
            amount=remaining, note=note,
        ))
    return allocs


def convert_to_andpad(estimate: Estimate) -> ANDPADBudget:
    """見積データからANDPAD工務予算データに変換"""
    budget = ANDPADBudget()
    occurrence = {}  # (工事番号, 明細コード) -> 出現回数

    for category in estimate.categories:
        for item in category.items:
            mapping = VENDOR_MAPPING.get(item.child_no)

            if mapping is None:
                # マッピングがない場合は工事種別名を発注先とする
                mapping = {"vendor": category.name, "kind": "材"}

            occ_key = (category.no, item.child_no)
            occ = occurrence.get(occ_key, 0)
            occurrence[occ_key] = occ + 1

            andpad_item = ANDPADItem(
                work_category=category.name,
                item_name=item.name,
                summary=item.summary,
                quantity=item.quantity,
                unit=item.unit,
                allocations=_build_allocations(mapping, item),
                stable_key=f"{category.no}_{item.child_no}_{occ}",
            )
            budget.items.append(andpad_item)

    return budget
