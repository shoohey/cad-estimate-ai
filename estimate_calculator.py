"""見積計算エンジン - CADデータから概算見積を算出

計算式は keisanshiki-master_updated.md（2026年2月版・古居様邸実績反映）に準拠。
CAD数量データに存在しない数量（外壁面積・屋根面積・軒長さ等）は、
物件条件の建物間口・奥行（R000081/R000082）から幾何学的に概算する。
"""
import csv
import math
import os
from dataclasses import dataclass, field
from cad_parser import (CADData, FittingData, LIVING_ROOMS, NON_LIVING_ROOMS,
                         WATER_ROOMS, ENTRANCE_ROOMS, BALCONY_ROOMS,
                         BATH_ROOMS, DOMA_ROOMS, NOOK_ROOMS, STORAGE_ROOMS,
                         MIZUMAWARI_FLOOR_ROOMS)

TSUBO = 3.3057  # ㎡→坪換算


@dataclass
class EstimateItem:
    """見積明細項目"""
    parent_no: int        # 工事種別番号
    child_no: int         # 明細番号
    seq_no: int = 0       # 通番
    name: str = ""        # 名称
    summary: str = ""     # 摘要
    quantity: float = 0.0
    unit: str = ""
    estimate_price: int = 0    # 見積単価
    order_price: int = 0       # 発注単価
    estimate_amount: int = 0   # 見積金額
    order_amount: int = 0      # 発注金額
    formula: str = ""          # 計算式（参考）
    category: str = ""         # 中間階層カテゴリ
    # ANDPAD用材工分離
    andpad_vendor: str = ""    # 発注先（細目工種）
    andpad_material: int = 0   # 材料費
    andpad_labor: int = 0      # 施工費


@dataclass
class WorkCategory:
    """工事種別"""
    no: int
    name: str
    items: list = field(default_factory=list)

    @property
    def estimate_total(self) -> int:
        return sum(item.estimate_amount for item in self.items)

    @property
    def order_total(self) -> int:
        return sum(item.order_amount for item in self.items)

    @property
    def profit(self) -> int:
        return self.estimate_total - self.order_total

    @property
    def profit_rate(self) -> float:
        if self.estimate_total == 0:
            return 0.0
        return self.profit / self.estimate_total


@dataclass
class Estimate:
    """見積書"""
    property_name: str = ""
    owner_name: str = ""
    location: str = ""
    estimate_no: str = ""
    total_floor_area_m2: float = 0.0
    total_floor_area_tsubo: float = 0.0
    categories: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    @property
    def subtotal_estimate(self) -> int:
        """小計（01〜21工事）"""
        return sum(c.estimate_total for c in self.categories if c.no <= 21)

    @property
    def subtotal_order(self) -> int:
        """小計（01〜21工事）"""
        return sum(c.order_total for c in self.categories if c.no <= 21)

    @property
    def management_fee_estimate(self) -> int:
        """管理費（見積3%）"""
        return int(self.subtotal_estimate * 0.03)

    @property
    def management_fee_order(self) -> int:
        """管理費（発注2%）"""
        return int(self.subtotal_order * 0.02)

    @property
    def total_estimate_excl_tax(self) -> int:
        """税抜合計"""
        total = self.subtotal_estimate
        for c in self.categories:
            if c.no >= 22:
                total += c.estimate_total
        return total

    @property
    def total_order_excl_tax(self) -> int:
        total = self.subtotal_order
        for c in self.categories:
            if c.no >= 22:
                total += c.order_total
        return total

    @property
    def tax(self) -> int:
        return int(self.total_estimate_excl_tax * 0.10)

    @property
    def total_estimate_incl_tax(self) -> int:
        raw = self.total_estimate_excl_tax + self.tax
        # 1000円単位に切り捨て
        return (raw // 1000) * 1000

    @property
    def total_profit(self) -> int:
        return self.total_estimate_excl_tax - self.total_order_excl_tax

    @property
    def total_profit_rate(self) -> float:
        if self.total_estimate_excl_tax == 0:
            return 0.0
        return self.total_profit / self.total_estimate_excl_tax


def load_unit_prices(csv_path: str) -> dict:
    """単価マスタCSVを読み込む"""
    prices = {}
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                code = int(row['工事コード'])
                prices[code] = {
                    'name': row['名称'],
                    'summary': row.get('摘要', ''),
                    'unit': row.get('単位', ''),
                    'order_price': int(row.get('発注単価', 0) or 0),
                    'estimate_price': int(row.get('見積単価', 0) or 0),
                    'note': row.get('数量根拠', ''),
                }
            except (ValueError, KeyError):
                continue
    return prices


def calculate_estimate(cad_data: CADData, prices: dict,
                       owner_name: str = "", location: str = "",
                       estimate_no: str = "",
                       include_adjustment: bool = True) -> Estimate:
    """CADデータから概算見積を算出

    Args:
        include_adjustment: ●33調整分（見積のみ計上・発注0円の調整項目）を
            含めるかどうか。実績見積では標準的に計上されるためデフォルトTrue。
    """
    estimate = Estimate(
        property_name=cad_data.property_name,
        owner_name=owner_name,
        location=location,
        estimate_no=estimate_no,
        total_floor_area_m2=cad_data.total_floor_area(),
        total_floor_area_tsubo=cad_data.total_floor_area_tsubo(),
    )

    # ===== 基本面積 =====
    total_area = cad_data.total_floor_area()          # 延床（部屋面積計）
    total_tsubo = total_area / TSUBO
    floor1_area = cad_data.floor_area("1階")
    floor2_area = cad_data.floor_area("2階")
    r_floor_area = cad_data.floor_area("R階")
    is_two_story = floor2_area > 0

    # ===== 特長条件 =====
    cond01 = cad_data.get_condition("R000051")  # 鋼板屋根
    cond02 = cad_data.get_condition("R000052")  # 準防火地域
    cond03 = cad_data.get_condition("R000053")  # IH採用
    cond04 = cad_data.get_condition("R000054")  # 太陽光発電
    is_gable = (cad_data.get_condition("R000031")
                + cad_data.get_condition("R000036")) >= 1  # 切妻屋根

    # ===== 建物外形（間口・奥行から概算） =====
    maguchi = cad_data.get_condition("R000081")   # 建物間口
    okuyuki = cad_data.get_condition("R000082")   # 建物奥行
    if maguchi > 0 and okuyuki > 0:
        perimeter = 2 * (maguchi + okuyuki)       # 建物外周長
        footprint = maguchi * okuyuki             # 建築面積（矩形近似）
    else:
        perimeter = math.sqrt(max(floor1_area, 1.0)) * 4
        footprint = floor1_area
        estimate.warnings.append(
            "建物間口・奥行データが無いため、外周長・屋根面積を1階床面積から概算しています。精度が低下する可能性があります。")

    # ===== 部屋タイプ別面積 =====
    genkan_area = cad_data.room_area_by_type("全階", ["玄関"])
    porch_area = cad_data.room_area_by_type("全階", ["ポーチ"])
    if porch_area == 0 and genkan_area > 0:
        porch_area = round(genkan_area * 1.25, 2)
        estimate.warnings.append(
            f"ポーチがCADデータに無いため、玄関面積×1.25＝{porch_area}㎡で概算計上しています。実際のポーチ面積と異なる場合は「見積修正」タブで調整してください。")
    balcony_area = cad_data.room_area_by_type("全階", BALCONY_ROOMS)
    has_balcony = balcony_area > 0
    doma_area = cad_data.room_area_by_type("全階", DOMA_ROOMS)   # 玄関・SIC等の土間床
    nook_area = cad_data.room_area_by_type("全階", NOOK_ROOMS)   # ヌック・畳コーナー
    bath_area = cad_data.room_area_by_type("全階", BATH_ROOMS)   # 浴室(UB)
    washitsu_area = cad_data.room_area_by_type("全階", ["和室"])

    # 水廻り床（洗面・脱衣・便所）
    mizu1_area = cad_data.room_area_by_type("1階", MIZUMAWARI_FLOOR_ROOMS)
    mizu2_area = cad_data.room_area_by_type("2階", MIZUMAWARI_FLOOR_ROOMS)

    # フローリング面積 = 床面積 −（水廻り床・浴室・土間・ヌック）
    excl1 = (mizu1_area + cad_data.room_area_by_type("1階", BATH_ROOMS)
             + cad_data.room_area_by_type("1階", DOMA_ROOMS)
             + cad_data.room_area_by_type("1階", NOOK_ROOMS))
    floor1_flooring = max(0.0, floor1_area - excl1)
    excl2 = (mizu2_area + cad_data.room_area_by_type("2階", BATH_ROOMS)
             + cad_data.room_area_by_type("2階", DOMA_ROOMS))
    floor2_flooring = max(0.0, floor2_area - excl2)

    # 施工床面積（部屋＋バルコニー＋ポーチ）
    sekou_area = total_area + balcony_area + porch_area
    sekou_tsubo = sekou_area / TSUBO

    # ===== 部屋数カウント =====
    toilet_count = cad_data.room_count("全階", ["トイレ", "便所"])
    senmen_count = cad_data.room_count("全階", ["洗面"])
    datsui_count = cad_data.room_count("全階", ["脱衣"])
    washroom_count = senmen_count + datsui_count
    bath_count = cad_data.room_count("全階", BATH_ROOMS)
    storage_count = cad_data.room_count("全階", STORAGE_ROOMS)
    living_room_count = cad_data.room_count("全階", LIVING_ROOMS)
    ldk_count = cad_data.room_count("全階", ["LDK", "居間", "リビング"])
    ldk_area = cad_data.room_area_by_type("全階", ["LDK", "居間", "リビング"])
    has_kitchen = cad_data.room_count("全階", ["LDK", "台所", "キッチン", "DK"]) > 0
    hall_count = cad_data.room_count("全階", ["ホール", "廊下"])
    genkan_count = cad_data.room_count("全階", ["玄関"])
    kaidan_count = max(1, cad_data.room_count("全階", ["階段"]))
    fukinuke_count = cad_data.room_count("全階", ["吹抜", "吹き抜け"])
    attic_count = cad_data.room_count("全階", ["小屋裏", "ロフト"])
    wic_count = cad_data.room_count("全階", ["WIC", "クローゼット"])
    nando_count = cad_data.room_count("全階", ["納戸", "押入"])
    garage_count = cad_data.room_count("全階", ["車庫", "ガレージ"])
    room_total_count = len(cad_data.rooms)
    has_tatami = washitsu_area > 0 or nook_area > 0
    has_attic = attic_count > 0
    attic_tsubo = cad_data.room_area_by_type("全階", ["小屋裏", "ロフト"]) / TSUBO

    # LDKは「居間」「台所」の両方として扱う（3路スイッチ・E付コンセント等）
    ima_count = max(cad_data.room_count("全階", ["居間", "リビング"]), ldk_count)
    daidokoro_count = max(cad_data.room_count("全階", ["台所", "キッチン"]),
                          ldk_count)

    # ===== 外皮の概算 =====
    # 外壁面積（開口含む）: 各階外周 × 階高2.9m。2階外周は床面積比の平方根で縮小
    if is_two_story and floor1_area > 0:
        upper_perimeter = min(perimeter,
                              perimeter * math.sqrt(floor2_area / floor1_area))
        ext_wall_area = (perimeter + upper_perimeter) * 2.9
    else:
        ext_wall_area = perimeter * 3.0
    # 妻壁面積（切妻の場合、勾配4寸想定）
    gable_area = (maguchi ** 2) * 0.2 if (is_gable and maguchi > 0) else 0.0

    # 屋根面積 = 建築面積 × 1.3（軒の出・勾配伸び分）
    roof_area = footprint * 1.3

    # 軒先・ケラバ長さ（屋根伏図データが無いため外周から概算）
    eaves_length = perimeter * 0.86                # 軒先（軒樋）長さ
    keraba_length = perimeter * (0.75 if is_gable else 0.15)
    roof_edge_length = eaves_length + keraba_length  # 軒＋ケラバ
    soffit_area = roof_edge_length * 0.56          # 軒天面積

    # ===== 入隅・巾木 =====
    dezumi_sum = cad_data.sum_quantity("N000081")  # 部屋出隅の合計
    if dezumi_sum > 0:
        ext_corner_count = max(2, round(dezumi_sum * 0.67))
    else:
        ext_corner_count = max(2, int(cad_data.sum_quantity("N000080") * 0.05))

    baseboard_length = cad_data.sum_quantity("N000130")
    if baseboard_length > 0:
        # 開口分20%を控除し、定尺2.7mで割って切り上げ
        baseboard_count = math.ceil(baseboard_length * 0.8 / 2.7)
    else:
        baseboard_count = math.ceil(total_tsubo * 1.2)

    # ===== 建具 =====
    ext_fittings = cad_data.get_external_fittings()
    int_fittings = cad_data.get_internal_fittings()
    ext_fitting_count = cad_data.get_fitting_count("金属")
    # 片引き戸の本数（造作手間の算定用）
    sliding_door_count = sum(
        int(f.quantities.get("T100001", 1) or 1)
        for f in int_fittings if "片引" in f.type_name)

    # ===== 条件警告 =====
    if cond01 == 1:
        estimate.warnings.append(
            "鋼板屋根（条件01=1）で計算しています。実際が瓦屋根の場合は「見積修正」タブで屋根工事を修正してください。※CAD設定と実仕様が異なる事例があります（図面確認必須）。")
    else:
        estimate.warnings.append(
            "瓦屋根（条件01=0）で計算しています。実際が鋼板屋根の場合は「見積修正」タブで屋根工事を修正してください。")
    if cond04 == 1:
        estimate.warnings.append(
            "太陽光発電システム工事を参考単価（長州産業5.09kw相当）で計上しています。パネル枚数・kW数により変動するため、確定見積で必ず調整してください。")
    else:
        estimate.warnings.append(
            "条件04=0のため太陽光発電システム工事・太陽光対応分電盤・HEMSは未計上です。採用する場合は「見積修正」タブで追加してください。")
    if cond03 == 1:
        estimate.warnings.append(
            "IH採用（条件03=1）のためガス工事は未計上です。ガス併用の場合は「見積修正」タブの23.屋外付帯工事から追加してください。")
    else:
        estimate.warnings.append(
            "ガス工事を参考単価で計上しています（条件03=0）。敷地条件により変動します。")
    estimate.warnings.append(
        "深基礎・布基礎・下がり天井・ニッチ・造作家具等の個別造作はCADデータから判定できないため未計上です。該当がある場合は「見積修正」タブで追加してください。")
    estimate.warnings.append(
        "屋外給排水・共通仮設・設計申請費は参考単価です。敷地条件・申請内容により変動します。")

    # ===== 01 仮設工事 =====
    cat01 = WorkCategory(no=1, name="仮設工事")

    # 外部足場（見付面積＝外壁＋妻壁 ×1.15）
    scaffold_area = (ext_wall_area + gable_area) * 1.15
    add_item(cat01, 10001, prices, scaffold_area)

    add_item(cat01, 10002, prices, 1)            # 仮設トイレ
    add_item(cat01, 10003, prices, sekou_area)   # 廃材処分費（施工床面積）
    add_item(cat01, 10004, prices, sekou_area)   # 運搬費（施工床面積）
    add_item(cat01, 10005, prices, 1)            # 駐車場代
    add_item(cat01, 10006, prices, 1)            # 基礎廻り養生シート

    # 安全対策費（施工床面積帯別）
    if sekou_area <= 120:
        add_item(cat01, 10007, prices, 1)
    elif sekou_area <= 180:
        add_item(cat01, 10008, prices, 1)
    else:
        add_item(cat01, 10009, prices, 1)

    add_item(cat01, 10010, prices, 1)            # 整地費用

    estimate.categories.append(cat01)

    # ===== 02 基礎工事 =====
    cat02 = WorkCategory(no=2, name="基礎工事")

    add_item(cat02, 20001, prices, 1)                          # 水盛り遣方
    foundation_area = floor1_area + porch_area
    add_item(cat02, 20002, prices, foundation_area)            # ベタ基礎（1階+ポーチ）

    # 人通口補強（基礎面積60㎡超過分は㎡追加）
    add_item(cat02, 20003, prices, 1)
    if foundation_area > 60:
        add_item(cat02, 20004, prices, foundation_area - 60)

    # 下地土間コンクリート（土間タイプの部屋＋ポーチ）
    add_item(cat02, 20005, prices, doma_area + porch_area)

    add_item(cat02, 20006, prices, 2)                          # ポンプ車（定数2台）

    estimate.categories.append(cat02)

    # ===== 03 木工事 =====
    cat03 = WorkCategory(no=3, name="木工事")

    # --- 直接明細（床材） ---
    add_item(cat03, 30001, prices, floor1_flooring / TSUBO, category="床工事")
    if floor2_flooring > 0:
        add_item(cat03, 30002, prices, floor2_flooring / TSUBO, category="床工事")
    if r_floor_area > 0:
        add_item(cat03, 30003, prices, r_floor_area / TSUBO, category="床工事")
    mizu_total = mizu1_area + mizu2_area
    if mizu_total > 0:
        add_item(cat03, 30004, prices, mizu_total / TSUBO, category="床工事")

    # --- 軸組・羽柄材・PC ---
    add_item(cat03, 30005, prices, sekou_tsubo, category="軸組・羽柄材・PC")
    add_item(cat03, 30006, prices, 1, category="軸組・羽柄材・PC")   # 金物費捻出(マイナス)
    add_item(cat03, 30007, prices, ext_corner_count, category="軸組・羽柄材・PC")  # 割増A
    # 割増B：2階比（平屋は1階床面積の坪数）
    if is_two_story:
        diff_tsubo = max(0.0, (floor1_area - floor2_area)) / TSUBO
    else:
        diff_tsubo = floor1_area / TSUBO
    add_item(cat03, 30008, prices, diff_tsubo, category="軸組・羽柄材・PC")
    add_item(cat03, 30010, prices, 1, category="軸組・羽柄材・PC")   # 金物費
    add_item(cat03, 30012, prices, 1, category="軸組・羽柄材・PC")   # 構造追加費

    # --- 建材・下地材 ---
    add_item(cat03, 30014, prices, sekou_tsubo, category="建材・下地材")  # 下地材基本
    add_item(cat03, 30015, prices, ext_corner_count, category="建材・下地材")
    add_item(cat03, 30016, prices, 1, category="建材・下地材")   # 軒天無し(マイナス・破風レス)
    add_item(cat03, 30017, prices, 1, category="建材・下地材")   # 養生費・補修費捻出
    add_item(cat03, 30018, prices, 1, category="建材・下地材")   # 玄関框
    add_item(cat03, 30019, prices, 1, category="建材・下地材")   # 上框（無垢床用）
    if doma_area > 0:
        add_item(cat03, 30020, prices, doma_area, category="建材・下地材")  # 床組無し(マイナス)
    if is_two_story:
        add_item(cat03, 30021, prices, 1, category="建材・下地材")  # 階段
        add_item(cat03, 30022, prices, 1, category="建材・下地材")  # 階段手摺
    add_item(cat03, 30023, prices, max(4, storage_count), category="建材・下地材")  # 枕棚+パイプ
    add_item(cat03, 30024, prices, toilet_count + 1, category="建材・下地材")  # 下地補強
    if nook_area > 0:
        add_item(cat03, 30026, prices, nook_area, category="建材・下地材")  # 大壁和室造作材
        add_item(cat03, 30027, prices, nook_area, category="建材・下地材")  # 床上げ用材料
        add_item(cat03, 30051, prices, 1, category="建材・下地材")  # ヌックスペース造作追加
    if nando_count > 0:
        add_item(cat03, 30028, prices, nando_count, category="建材・下地材")  # 中段・枕棚セット
    if cond01 == 1:
        add_item(cat03, 30030, prices, roof_area, category="建材・下地材")  # 鋼板屋根遮音下地
    add_item(cat03, 30031, prices, ext_wall_area, category="建材・下地材")  # 透湿防水シート
    panel_count = math.ceil((ext_wall_area + gable_area) / 2.72)
    add_item(cat03, 30032, prices, panel_count, category="建材・下地材")  # 耐力面材
    add_item(cat03, 30033, prices, 4, category="建材・下地材")   # 制震ダンパー
    add_item(cat03, 30034, prices, ext_fitting_count, category="建材・下地材")  # 窓枠
    add_item(cat03, 30035, prices, baseboard_count, category="建材・下地材")   # 化粧巾木
    add_item(cat03, 30038, prices, roof_edge_length, category="建材・下地材")  # 換気部材
    add_item(cat03, 30043, prices, 1, category="建材・下地材")   # 小屋裏耐力面材(3寸超6寸未満)
    add_item(cat03, 30044, prices, 1, category="建材・下地材")   # 床養生費
    add_item(cat03, 30045, prices, 1, category="建材・下地材")   # 補修費
    if has_kitchen:
        # 床フロアタイル下地合板（キッチン・パントリー部）
        tile_floor_area = ldk_area * 0.3 if ldk_area > 0 else 10.0
        add_item(cat03, 30048, prices, math.ceil(tile_floor_area / 1.65),
                 category="建材・下地材")
        add_item(cat03, 30049, prices, 1, category="建材・下地材")  # 床見切
    add_item(cat03, 30057, prices, 1, category="建材・下地材")   # 電気設備用点検口

    # --- 大工手間 ---
    add_item(cat03, 30101, prices, sekou_tsubo, category="大工手間")  # 基本手間
    add_item(cat03, 30102, prices, ext_corner_count, category="大工手間")  # 外周入隅
    if has_balcony or porch_area > 0:
        add_item(cat03, 30103, prices, balcony_area + porch_area, category="大工手間")
    add_item(cat03, 30104, prices, 1, category="大工手間")       # 新テンプレート用調整費
    if sliding_door_count > 0:
        add_item(cat03, 30105, prices, max(1, round(sliding_door_count * 0.5)),
                 category="大工手間")                             # 片引き戸造作手間
    if is_two_story and floor1_area > floor2_area:
        add_item(cat03, 30106, prices, floor1_area - floor2_area, category="大工手間")  # 下屋
    add_item(cat03, 30107, prices, total_area, category="大工手間")  # 外部耐力面材貼り
    add_item(cat03, 30108, prices, max(0, soffit_area - 5), category="大工手間")  # 軒天下地
    add_item(cat03, 30109, prices, max(0, soffit_area - 5), category="大工手間")  # 軒天張り手間
    add_item(cat03, 30110, prices, storage_count, category="大工手間")  # 中段・枕棚取付
    extra_storage = max(0, storage_count - 4)
    if extra_storage > 0:
        add_item(cat03, 30111, prices, extra_storage, category="大工手間")  # 収納造作
    if is_two_story:
        add_item(cat03, 30113, prices, 1, category="大工手間")   # 階段（非フルプレカット）
        add_item(cat03, 30114, prices, 1, category="大工手間")   # 階段下収納・トイレ造作
    add_item(cat03, 30115, prices, living_room_count, category="大工手間")  # 床ガラリ
    add_item(cat03, 30116, prices, total_area, category="大工手間")  # 壁・天井断熱材手間
    add_item(cat03, 30117, prices, total_area, category="大工手間")  # 省令準耐火手間
    if nook_area > 0:
        add_item(cat03, 30118, prices, nook_area, category="大工手間")  # 床上げ加工
    add_item(cat03, 30120, prices, 1, category="大工手間")       # 建て方レッカー費

    # --- 含み予算 ---
    add_item(cat03, 30121, prices, 1, category="【含み予算】")
    if total_area >= 100:
        add_item(cat03, 30122, prices, total_area - 100, category="【含み予算】")
    add_item(cat03, 30123, prices, 1, category="【含み予算】")
    add_item(cat03, 30124, prices, 1, category="【含み予算】")
    add_item(cat03, 30125, prices, 1, category="【含み予算】")
    add_item(cat03, 30126, prices, 1, category="【含み予算】")
    add_item(cat03, 30127, prices, 1, category="【含み予算】")

    # --- 見積調整（33調整分：見積のみ計上・発注0） ---
    if include_adjustment:
        add_item(cat03, 30056, prices, 1, category="【見積調整】")
        add_item(cat03, 30129, prices, 1, category="【見積調整】")
        estimate.warnings.append(
            "●33調整分（建材・下地材/大工手間 各475,000円・発注0円）を計上しています。除外する場合はサイドバーのオプションを外してください。")

    estimate.categories.append(cat03)

    # ===== 04 断熱工事 =====
    cat04 = WorkCategory(no=4, name="断熱工事")

    add_item(cat04, 40001, prices, ext_wall_area * 1.05)   # 壁断熱材（外気に接する面）
    add_item(cat04, 40002, prices, footprint)              # 天井断熱材（水平投影）
    if is_two_story:
        add_item(cat04, 40003, prices, floor2_area)        # 2階直下天井吸音材
    # 基礎断熱（外周長ベース）
    add_item(cat04, 40004, prices, math.ceil(perimeter * 0.43 / 1.62))  # 外周立上り
    add_item(cat04, 40005, prices, math.ceil(perimeter * 0.91 / 1.62))  # 外周土間
    if doma_area > 0:
        add_item(cat04, 40006, prices, math.ceil(doma_area / 1.62))     # 内部土間床
    add_item(cat04, 40007, prices, math.ceil((doma_area + porch_area) / 1.62))  # ポーチ部土間
    # 気密工事
    add_item(cat04, 40008, prices, 2)                                   # 可変透湿気密シート
    add_item(cat04, 40009, prices, math.ceil(total_area / 50))          # 気密テープ
    add_item(cat04, 40010, prices, footprint)              # 気密シート貼り手間（天井面）
    add_item(cat04, 40011, prices, footprint)              # 気流止め
    add_item(cat04, 40012, prices, total_tsubo)            # 気密手間（大工）

    estimate.categories.append(cat04)

    # ===== 05 屋根工事 =====
    cat05 = WorkCategory(no=5, name="屋根工事")

    if cond01 == 1:  # 鋼板屋根
        add_item(cat05, 50001, prices, roof_area)   # 小屋裏換気(鋼板用)
        add_item(cat05, 50003, prices, 1)           # 荷揚げ費
        add_item(cat05, 50004, prices, roof_area)   # SGL鋼板
    else:  # 瓦屋根
        add_item(cat05, 50002, prices, roof_area)   # 小屋裏換気(瓦用)
        add_item(cat05, 50005, prices, roof_area)   # 平板瓦葺

    estimate.categories.append(cat05)

    # ===== 06 板金・樋工事 =====
    cat06 = WorkCategory(no=6, name="板金・樋工事")

    add_item(cat06, 60001, prices, eaves_length)               # 軒樋
    catcher_count = math.ceil(roof_area / 45) + 2              # 集水器（下屋・バルコニー分含む）
    add_item(cat06, 60002, prices, catcher_count)
    downspout_length = catcher_count * (6 if is_two_story else 3.5)
    add_item(cat06, 60003, prices, downspout_length)           # 竪樋

    estimate.categories.append(cat06)

    # ===== 07 外部建具工事 =====
    cat07 = WorkCategory(no=7, name="外部建具工事")

    for fitting in ext_fittings:
        count = int(fitting.quantities.get("T100001", 1) or 1)
        price_codes = match_external_fitting_price(fitting, prices)
        for pc in price_codes:
            item = create_item_from_fitting(pc, prices, count, fitting,
                                            parent_no=7)
            if item:
                cat07.items.append(item)

    # 建具がマッチしなかった場合のフォールバック
    if not cat07.items:
        add_item(cat07, 70002, prices, 1)   # 玄関ドア
        ext_count = max(1, ext_fitting_count - 1)
        add_item(cat07, 70010, prices, ext_count)  # 標準サッシ（概算）

    estimate.warnings.append(
        "外部建具はYKKap標準（樹脂サッシAPW430・玄関ドア ヴェナートD30＋電気錠）で計上しています。内外観ブラック色等の追加費用（63,900円）は含みません。")

    estimate.categories.append(cat07)

    # ===== 08 内部建具工事 =====
    cat08 = WorkCategory(no=8, name="内部建具工事")

    for fitting in int_fittings:
        count = int(fitting.quantities.get("T100001", 1) or 1)
        price_code = match_internal_fitting_price(fitting, prices)
        if price_code:
            item = create_item_from_fitting(price_code, prices, count, fitting,
                                            parent_no=8)
            if item:
                cat08.items.append(item)

    # フォールバック
    if not cat08.items:
        door_count = max(5, cad_data.get_fitting_count("木製"))
        add_item(cat08, 80007, prices, door_count)

    # 可動棚
    if storage_count > 0:
        add_item(cat08, 80001, prices, min(storage_count, 4))

    # ヌックスペースの棚板（標準構成：上部2枚＋固定棚板2枚）
    if nook_area > 0:
        add_item(cat08, 80023, prices, 2)
        add_item(cat08, 80024, prices, 2)

    estimate.warnings.append(
        "内部建具は標準グレード（LIXIL ラシッサS）で計上しています。ラフィス等の上位グレード採用時は差額が発生します（1本あたり+3〜7万円程度）。")

    estimate.categories.append(cat08)

    # ===== 09 外装工事 =====
    cat09 = WorkCategory(no=9, name="外装工事")

    add_item(cat09, 90004, prices, ext_wall_area)   # サイディング
    # 出隅コーナー役物（出隅数×建物高さ）
    corner_length = (ext_corner_count + 4) * (5.8 if is_two_story else 3.0)
    add_item(cat09, 90001, prices, corner_length)
    add_item(cat09, 90003, prices, ext_wall_area)    # 残材処理
    add_item(cat09, 90002, prices, roof_edge_length)  # 破風・鼻隠し
    add_item(cat09, 90005, prices, soffit_area)      # 軒天

    estimate.categories.append(cat09)

    # ===== 10 左官工事 =====
    cat10 = WorkCategory(no=10, name="左官工事")

    # 玄関・ポーチ タイル下地（17.5㎡以下は一式）
    if genkan_area + porch_area <= 17.5:
        add_item(cat10, 100002, prices, 1)
    else:
        add_item(cat10, 100002, prices, genkan_area + porch_area)
    # 基礎巾木CTバリヤ（基礎外周×立上り0.43m）
    base_moru_area = perimeter * 0.43
    add_item(cat10, 100003, prices, math.ceil(base_moru_area / 14))  # 下地調整材
    add_item(cat10, 100004, prices, math.ceil(base_moru_area / 8))   # 本塗材
    add_item(cat10, 100005, prices, math.ceil(base_moru_area / 20))  # トップコート
    add_item(cat10, 100006, prices, base_moru_area)                  # モルタル刷毛引き

    estimate.categories.append(cat10)

    # ===== 11 タイル工事 =====
    cat11 = WorkCategory(no=11, name="タイル工事")
    # ポーチ面積＋ポーチ側面＋玄関床＋玄関巾木面積
    genkan_habaki = cad_data.sum_quantity("N000122", room_types=["玄関"])
    tile_area = porch_area + porch_area * 0.3 + genkan_area + genkan_habaki
    add_item(cat11, 110001, prices, tile_area)
    estimate.warnings.append(
        "玄関・ポーチのタイル面積は概算です（ポーチ側面・框立上りを含む実面積はCADから取得できません）。実績では概算の1.5〜2倍になる事例があります。")
    estimate.categories.append(cat11)

    # ===== 12 塗装工事 =====
    cat12 = WorkCategory(no=12, name="塗装工事")
    add_item(cat12, 120001, prices, 1)   # 室内木部塗装（框・造作材等の標準塗装）
    estimate.categories.append(cat12)

    # ===== 13 内装工事 =====
    cat13 = WorkCategory(no=13, name="内装工事")

    add_item(cat13, 130002, prices, total_tsubo)      # クロス工事（一般部）
    if has_attic:
        add_item(cat13, 130003, prices, attic_tsubo)  # 小屋裏クロス
    extra_closet = max(0, storage_count - 6)
    if extra_closet > 0:
        add_item(cat13, 130004, prices, extra_closet)  # 収納追加クロス
    if has_tatami:
        tatami_area = washitsu_area + nook_area
        tatami_count = max(1, round(tatami_area / 0.81))  # 半帖換算
        add_item(cat13, 130001, prices, tatami_count)
    if has_kitchen:
        add_item(cat13, 130006, prices, 1)            # キッチン・パントリー フロアタイル

    estimate.categories.append(cat13)

    # ===== 14 住宅設備機器工事 =====
    cat14 = WorkCategory(no=14, name="住宅設備機器工事")

    add_item(cat14, 140001, prices, 1)   # キッチン
    add_item(cat14, 140004, prices, 1)   # カップボード（予算組）
    add_item(cat14, 140005, prices, 1)   # キッチン標準水栓（タッチレス）

    # UB（部屋面積でサイズ選定：1616は約3.3㎡、1620は約4.1㎡）
    if bath_area >= 3.7:
        add_item(cat14, 140006, prices, 1)  # UB 1620
    else:
        add_item(cat14, 140007, prices, 1)  # UB 1616

    # トイレ
    if toilet_count >= 1:
        add_item(cat14, 140008, prices, 1)  # タンクレストイレ
    if toilet_count >= 2:
        add_item(cat14, 140009, prices, toilet_count - 1)  # 2台目以降

    add_item(cat14, 140011, prices, 1)                  # 造作洗面
    add_item(cat14, 140013, prices, toilet_count + 1)   # タオル掛け（トイレ＋洗面）
    add_item(cat14, 140014, prices, toilet_count)       # ペーパーホルダー
    add_item(cat14, 140015, prices, 1)                  # 洗濯パン

    # エコキュート
    if total_area >= 120:
        add_item(cat14, 140016, prices, 1)  # 460L
    else:
        add_item(cat14, 140017, prices, 1)  # 370L

    estimate.categories.append(cat14)

    # ===== 15 給排水設備工事 =====
    cat15 = WorkCategory(no=15, name="給排水設備工事")

    add_item(cat15, 150001, prices, 1)  # 基本（給排水6ヶ所以下・給湯3ヶ所以下）
    extra_plumbing = max(0, (toilet_count + washroom_count + bath_count + 1) - 6)
    if extra_plumbing > 0:
        add_item(cat15, 150002, prices, extra_plumbing)

    estimate.categories.append(cat15)

    # ===== 16 電気設備工事 =====
    cat16 = WorkCategory(no=16, name="電気設備工事")

    add_item(cat16, 160001, prices, 1)  # 仮設電気

    # 片切スイッチ：居室＋玄関＋便所＋浴室＋脱衣＋洗面＋吹抜＋小屋裏＋4
    switch_count = (living_room_count + genkan_count + toilet_count + bath_count
                    + datsui_count + senmen_count + fukinuke_count
                    + attic_count + 4)
    add_item(cat16, 160002, prices, switch_count)

    # 3路4路スイッチ：(ホール廊下＋居間＋台所＋階段室＋車庫)×2
    three_way_count = (hall_count + ima_count + daidokoro_count
                       + kaidan_count + garage_count) * 2
    add_item(cat16, 160003, prices, three_way_count)

    # 電灯配線：全部屋＋居間＋ホール廊下＋玄関＋12
    light_count = room_total_count + ima_count + hall_count + genkan_count + 12
    add_item(cat16, 160004, prices, light_count)

    # 一般コンセント
    outlet_count = (living_room_count * 2 + hall_count * 2 + ima_count * 2
                    + wic_count + datsui_count + senmen_count
                    + kaidan_count + attic_count)
    add_item(cat16, 160005, prices, outlet_count)

    # E付標準コンセント：キッチン系×2＋便所＋浴室＋脱衣/洗面
    e_outlet_count = ((2 if (ima_count or daidokoro_count) else 0)
                      + toilet_count + (1 if bath_count else 0)
                      + (1 if (datsui_count or senmen_count) else 0))
    add_item(cat16, 160006, prices, e_outlet_count)

    if has_kitchen:
        add_item(cat16, 160007, prices, 2)   # E付専用（レンジ・食洗機）
    if cond03 == 1 and has_kitchen:
        add_item(cat16, 160008, prices, 1)   # IH用
    add_item(cat16, 160009, prices, 1)       # 自動水栓用（タッチレス水栓標準）

    add_item(cat16, 160010, prices, living_room_count)       # エアコン用コンセント
    add_item(cat16, 160011, prices, living_room_count + 2)   # TV配線
    add_item(cat16, 160012, prices, 1)                       # AV配管（壁掛TV用）
    add_item(cat16, 160013, prices, 2)                       # 防水コンセント
    add_item(cat16, 160014, prices, 1)                       # TEL・INT引込配管
    add_item(cat16, 160015, prices, 1)                       # LAN配線

    # 火災警報器：居室＋階段室＋1
    alarm_count = living_room_count + kaidan_count + 1
    add_item(cat16, 160016, prices, alarm_count)
    add_item(cat16, 160017, prices, alarm_count)

    add_item(cat16, 160018, prices, 2)   # ホーム保安灯
    add_item(cat16, 160019, prices, 1)   # エアコンスリーブ・ダクト周り
    add_item(cat16, 160020, prices, 1)   # 省令準耐火対応プレート
    add_item(cat16, 160021, prices, 1)   # テレビドアホン
    add_item(cat16, 160022, prices, 1)   # インターホン取付費
    add_item(cat16, 160023, prices, 1)   # 幹線引込・分電盤

    # 太陽光対応（条件04）
    if cond04 == 1:
        add_item(cat16, 160024, prices, 1)  # 太陽光対応分電盤変更
        add_item(cat16, 160025, prices, 1)  # スマートコスモ分電盤変更
        add_item(cat16, 160026, prices, 1)  # HEMS本体
        add_item(cat16, 160027, prices, 1)  # HEMS設定

    add_item(cat16, 160028, prices, 1)   # 電力会社申請

    estimate.categories.append(cat16)

    # ===== 17 換気システム工事 =====
    cat17 = WorkCategory(no=17, name="換気システム工事")

    # 換気対象面積 = 延床＋吹抜−浴室(UB)
    vent_area = total_area - bath_area
    if vent_area <= 63:
        vent_code, vent_units = 170001, 1
    elif vent_area <= 103:
        vent_code, vent_units = 170002, 1
    elif vent_area <= 123:
        vent_code, vent_units = 170003, 2   # 小×2台
    elif vent_area <= 162:
        vent_code, vent_units = 170004, 2   # 小1大1
    else:
        vent_code, vent_units = 170005, 2   # 大×2台
    add_item(cat17, vent_code, prices, 1)

    add_item(cat17, 170006, prices, total_tsubo)   # ダクト材
    add_item(cat17, 170007, prices, 1)             # 取付施工費
    add_item(cat17, 170008, prices, vent_units)    # 電源工事（本体台数）

    # 屋外フード（φ100：便所＋浴室＋2、φ150：給排気用2）
    hood100_count = toilet_count + bath_count + 2
    if cond02 == 1:  # 準防火
        add_item(cat17, 170014, prices, hood100_count)
        add_item(cat17, 170016, prices, 2)
    else:
        add_item(cat17, 170013, prices, hood100_count)
        add_item(cat17, 170015, prices, 2)

    estimate.warnings.append(
        "第一種熱交換換気を前提に局所換気扇は未計上です。壁付換気扇が必要な場合は「見積修正」タブで追加してください（1台あたり約25,000円）。")

    estimate.categories.append(cat17)

    # ===== 18 エアコン工事 =====
    cat18 = WorkCategory(no=18, name="エアコン工事")

    ac_count = 0
    # LDK用（面積でサイズ選定）
    if ldk_count > 0:
        if ldk_area >= 25:
            add_item(cat18, 180001, prices, 1)  # 5.6kw
        else:
            add_item(cat18, 180002, prices, 1)  # 4.0kw
        ac_count += 1

    # 寝室・子供室用（部屋ごとに面積でサイズ選定）
    bedroom_types = ["主寝室", "寝室", "子供室", "洋室", "書斎", "趣味室"]
    for room in cad_data.rooms:
        if any(bt in room.name for bt in bedroom_types):
            room_area = room.quantities.get("N000001", 0.0)
            if room_area >= 10:
                add_item(cat18, 180003, prices, 1)  # 2.8kw
            else:
                add_item(cat18, 180004, prices, 1)  # 2.2kw
            ac_count += 1

    if ac_count > 0:
        add_item(cat18, 180006, prices, ac_count)   # 無線LANアダプター
    if is_two_story:
        add_item(cat18, 180007, prices, 1)          # 延長配管（2階→1階）

    estimate.warnings.append(
        "エアコンは全居室分を計上しています（LDK＋寝室・子供室等）。設置台数は施主様のご意向により変動します。")

    estimate.categories.append(cat18)

    # ===== 19 太陽光発電システム工事 =====
    cat19 = WorkCategory(no=19, name="太陽光発電システム工事")
    if cond04 == 1:
        add_item(cat19, 190001, prices, 1)
    estimate.categories.append(cat19)

    # ===== 20 造作家具工事 =====
    cat20 = WorkCategory(no=20, name="造作家具工事")
    add_item(cat20, 200001, prices, 1)  # ユーティリティカウンター
    estimate.categories.append(cat20)

    # ===== 21 雑工事 =====
    cat21 = WorkCategory(no=21, name="雑工事")
    add_item(cat21, 210001, prices, total_area)  # 清掃
    floor_count = 2 if is_two_story else 1
    add_item(cat21, 210002, prices, floor_count)  # 天井点検口
    add_item(cat21, 210003, prices, 1)  # 床下点検口
    add_item(cat21, 210004, prices, 2)  # 室内物干し
    estimate.categories.append(cat21)

    # ===== 22 管理費 =====
    cat22 = WorkCategory(no=22, name="管理費")
    mgmt_item = EstimateItem(
        parent_no=22, child_no=1,
        name="管理費", summary="小計×3%/2%", unit="式",
        quantity=1,
        estimate_price=estimate.management_fee_estimate,
        order_price=estimate.management_fee_order,
        estimate_amount=estimate.management_fee_estimate,
        order_amount=estimate.management_fee_order,
    )
    cat22.items.append(mgmt_item)
    estimate.categories.append(cat22)

    # ===== 23 屋外付帯工事 =====
    cat23 = WorkCategory(no=23, name="屋外付帯工事")
    add_item(cat23, 220001, prices, 1)  # 屋外給排水
    add_item(cat23, 220002, prices, 1)  # 共通仮設
    if cond03 == 0:  # IH以外はガス工事
        add_item(cat23, 220003, prices, 1)
    estimate.categories.append(cat23)

    # ===== 24 諸費用 =====
    cat24 = WorkCategory(no=24, name="諸費用")
    add_item(cat24, 230001, prices, 1)  # コーディネート
    add_item(cat24, 230002, prices, 1)  # 設計申請費
    estimate.categories.append(cat24)

    return estimate


def add_item(work_cat: WorkCategory, code: int, prices: dict,
             quantity: float, category: str = ""):
    """明細項目を追加"""
    if code not in prices:
        return

    p = prices[code]
    qty = round(quantity, 2)
    if qty == 0:
        return

    est_amount = int(qty * p['estimate_price'])
    ord_amount = int(qty * p['order_price'])

    item = EstimateItem(
        parent_no=work_cat.no,
        child_no=code,
        name=p['name'],
        summary=p['summary'],
        quantity=qty,
        unit=p['unit'],
        estimate_price=p['estimate_price'],
        order_price=p['order_price'],
        estimate_amount=est_amount,
        order_amount=ord_amount,
        category=category,
    )
    work_cat.items.append(item)


def match_external_fitting_price(fitting: FittingData, prices: dict) -> list:
    """外部建具をマスタとマッチングし、単価コードのリストを返す。

    玄関ドアは電気錠、大型引違い窓は大判ガラス施工費を併せて返す。
    """
    width = fitting.quantities.get("T100010", 0)
    height = fitting.quantities.get("T100012", 0)
    t = fitting.type_name

    # 玄関ドア判定（金属戸・片開/親子・高さ1.9m以上）
    if "戸" in fitting.category and ("片開" in t or "親子" in t) and height >= 1.9:
        # YKKap ヴェナートD30 ＋ 電気錠（標準セット）
        return [70002, 70004]

    # 引違い窓判定（幅×高さでサイズマッチング）
    if "引違" in t:
        if width >= 2.4:
            codes = [70014]  # 25122-2 サポートハンドル
            if height >= 2.2:
                codes.append(70016)  # 大判ガラス施工費
            return codes
        if width >= 1.5:
            if height >= 1.7:
                return [70012]  # 16020
            if height >= 0.7:
                return [70010]  # 16009
            return [70009]      # 16005
        return [70009]

    # FIX窓（幅でサイズマッチング）
    if "Fix" in t or "FIX" in t or "fix" in t:
        if width <= 0.8:
            return [70005]  # 07409
        if width <= 1.2:
            return [70006]  # 11909
        return [70008]      # 16003

    # その他（縦すべり・横すべり等）は同サイズ帯の引違い相当で概算
    if width >= 1.5:
        return [70010]
    return [70009]


def match_internal_fitting_price(fitting: FittingData, prices: dict) -> int:
    """内部建具をマスタとマッチング（標準グレード：LIXIL ラシッサS）"""
    width = fitting.quantities.get("T100010", 0)
    t = fitting.type_name

    # 室内窓（木製FIX窓 → デコマド）
    if "窓" in fitting.category and ("Fix" in t or "FIX" in t or "fix" in t):
        return 80022

    # 折戸（4枚折戸含む）
    if "折戸" in t:
        if width >= 1.5:
            return 80014  # W6尺
        return 80013      # W3尺

    # 片引き・引込み
    if "片引" in t or "引込" in t:
        return 80011

    # 引違い（室内）
    if "引違" in t:
        return 80011

    # 片開き・開き戸
    return 80007


def create_item_from_fitting(price_code: int, prices: dict,
                              count: int, fitting: FittingData,
                              parent_no: int = 0) -> EstimateItem:
    """建具から見積項目を作成"""
    if price_code not in prices:
        return None

    p = prices[price_code]
    return EstimateItem(
        parent_no=parent_no,
        child_no=price_code,
        name=p['name'],
        summary=f"{fitting.material} {fitting.type_name}",
        quantity=count,
        unit=p['unit'],
        estimate_price=p['estimate_price'],
        order_price=p['order_price'],
        estimate_amount=int(count * p['estimate_price']),
        order_amount=int(count * p['order_price']),
    )
