"""ANDPAD工務予算振り分けモジュール - 積算原価から材工分離を行う"""
from dataclasses import dataclass, field
from estimate_calculator import Estimate, EstimateItem, WorkCategory


# 細目工種（ANDPAD発注先）マッピング
VENDOR_MAPPING = {
    # 仮設工事
    10001: {"vendor": "仮設足場", "split": "labor"},
    10002: {"vendor": "仮設工事", "split": "labor"},
    10003: {"vendor": "産廃", "split": "labor"},
    10004: {"vendor": "運搬費", "split": "labor"},
    10005: {"vendor": "仮設工事", "split": "labor"},
    10006: {"vendor": "仮設工事", "split": "labor"},
    10007: {"vendor": "仮設工事", "split": "labor"},
    10008: {"vendor": "仮設工事", "split": "labor"},
    10009: {"vendor": "仮設工事", "split": "labor"},
    10010: {"vendor": "仮設工事", "split": "labor"},

    # 基礎工事
    20001: {"vendor": "基礎工事", "split": "labor"},
    20002: {"vendor": "基礎工事", "split": "labor"},
    20003: {"vendor": "基礎工事", "split": "labor"},
    20004: {"vendor": "基礎工事", "split": "labor"},
    20005: {"vendor": "基礎工事", "split": "labor"},
    20006: {"vendor": "基礎工事", "split": "labor"},
    20007: {"vendor": "基礎工事", "split": "labor"},
    20008: {"vendor": "基礎工事", "split": "labor"},

    # 木工事 - プレカット
    30005: {"vendor": "プレカット", "split": "material"},
    30006: {"vendor": "プレカット", "split": "material"},
    30007: {"vendor": "プレカット", "split": "material"},
    30008: {"vendor": "プレカット", "split": "material"},
    30009: {"vendor": "プレカット", "split": "material"},
    30010: {"vendor": "金物", "split": "material"},
    30012: {"vendor": "プレカット", "split": "material"},

    # 木工事 - 建材
    30001: {"vendor": "建材（標準床材）土橋", "split": "material_labor", "labor_ratio": 0.3},
    30002: {"vendor": "建材（標準床材）土橋", "split": "material_labor", "labor_ratio": 0.3},
    30003: {"vendor": "建材（標準床材）土橋", "split": "material_labor", "labor_ratio": 0.3},
    30004: {"vendor": "建材（特殊床材）", "split": "material_labor", "labor_ratio": 0.3},
    30014: {"vendor": "建材（土橋）", "split": "material"},
    30015: {"vendor": "建材（土橋）", "split": "material"},
    30016: {"vendor": "建材（土橋）", "split": "material"},
    30017: {"vendor": "建材（土橋）", "split": "material"},
    30018: {"vendor": "建材（特殊床材・框）", "split": "material"},
    30019: {"vendor": "建材（特殊床材・框）", "split": "material"},
    30020: {"vendor": "建材（土橋）", "split": "material"},
    30021: {"vendor": "建材（土橋）", "split": "material"},
    30022: {"vendor": "階段手摺金具（カワジュン）", "split": "material"},
    30023: {"vendor": "建材（土橋）", "split": "material"},
    30024: {"vendor": "建材（土橋）", "split": "material"},
    30025: {"vendor": "建材（土橋）", "split": "material"},
    30026: {"vendor": "建材（土橋）", "split": "material"},
    30027: {"vendor": "建材（土橋）", "split": "material"},
    30028: {"vendor": "建材（土橋）", "split": "material"},
    30030: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    30031: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    30032: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    30033: {"vendor": "MIRAIE", "split": "material"},
    30034: {"vendor": "インテリア建具（洋風造作材）中日トーヨー", "split": "material"},
    30035: {"vendor": "インテリア建具（洋風造作材）中日トーヨー", "split": "material"},
    30036: {"vendor": "建材（土橋）", "split": "three_way"},  # クロス/大工/建材で3等分
    30037: {"vendor": "金物", "split": "material"},
    30038: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    30041: {"vendor": "建材（土橋）", "split": "three_way"},
    30042: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    30043: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    30044: {"vendor": "床養生", "split": "labor"},
    30045: {"vendor": "補修工事", "split": "labor"},
    30046: {"vendor": "建材（土橋）", "split": "material"},
    30047: {"vendor": "建材（土橋）", "split": "three_way"},
    30048: {"vendor": "建材（土橋）", "split": "material"},
    30049: {"vendor": "建材（土橋）", "split": "material"},
    30050: {"vendor": "建材（土橋）", "split": "material_labor", "labor_ratio": 0.5},
    30051: {"vendor": "建材（土橋）", "split": "three_way"},
    30054: {"vendor": "建材（土橋）", "split": "three_way"},
    30055: {"vendor": "建材（土橋）", "split": "three_way"},
    30056: {"vendor": "建材（土橋）", "split": "material"},
    30057: {"vendor": "建材（土橋）", "split": "material_labor", "labor_ppu": 6000},

    # 木工事 - 大工手間
    30101: {"vendor": "大工手間", "split": "labor"},
    30102: {"vendor": "大工手間", "split": "labor"},
    30103: {"vendor": "大工手間", "split": "labor"},
    30104: {"vendor": "大工手間", "split": "labor"},
    30105: {"vendor": "大工手間", "split": "labor"},
    30106: {"vendor": "大工手間", "split": "labor"},
    30107: {"vendor": "大工手間", "split": "labor"},
    30108: {"vendor": "大工手間", "split": "labor"},
    30109: {"vendor": "大工手間", "split": "labor"},
    30110: {"vendor": "大工手間", "split": "labor"},
    30111: {"vendor": "大工手間", "split": "labor"},
    30112: {"vendor": "大工手間", "split": "labor"},
    30113: {"vendor": "大工手間", "split": "labor"},
    30114: {"vendor": "大工手間", "split": "labor"},
    30115: {"vendor": "大工手間", "split": "labor"},
    30116: {"vendor": "大工手間", "split": "labor"},
    30117: {"vendor": "大工手間", "split": "labor"},
    30118: {"vendor": "大工手間", "split": "labor"},
    30119: {"vendor": "大工手間", "split": "material_labor", "labor_ratio": 0.5},
    30120: {"vendor": "レッカー", "split": "labor"},
    30128: {"vendor": "大工手間", "split": "labor"},
    30129: {"vendor": "大工手間", "split": "labor"},

    # 含み予算
    30121: {"vendor": "保険・保証", "split": "material"},
    30122: {"vendor": "保険・保証", "split": "material"},
    30123: {"vendor": "保険・保証", "split": "material"},
    30124: {"vendor": "保険・保証", "split": "material"},
    30125: {"vendor": "保険・保証", "split": "material"},
    30126: {"vendor": "保険・保証", "split": "material"},
    30127: {"vendor": "保険・保証", "split": "material"},

    # 断熱工事
    40001: {"vendor": "断熱材・気密シート", "split": "material"},
    40002: {"vendor": "断熱材・気密シート", "split": "material"},
    40003: {"vendor": "断熱材・気密シート", "split": "material"},
    40004: {"vendor": "基礎断熱", "split": "material"},
    40005: {"vendor": "基礎断熱", "split": "material"},
    40006: {"vendor": "基礎断熱", "split": "material"},
    40007: {"vendor": "基礎断熱", "split": "material"},
    40008: {"vendor": "断熱材・気密シート", "split": "material"},
    40009: {"vendor": "シンホリ（ダイライト）", "split": "material"},
    40010: {"vendor": "気密手間", "split": "labor"},
    40011: {"vendor": "気密手間", "split": "labor"},
    40012: {"vendor": "気密手間", "split": "labor"},

    # 屋根工事
    50001: {"vendor": "鋼板屋根", "split": "labor"},
    50002: {"vendor": "瓦屋根", "split": "labor"},
    50003: {"vendor": "鋼板屋根", "split": "labor"},
    50004: {"vendor": "鋼板屋根", "split": "labor"},
    50005: {"vendor": "瓦屋根", "split": "labor"},

    # 板金工事
    60001: {"vendor": "板金工事", "split": "labor"},
    60002: {"vendor": "板金工事", "split": "labor"},
    60003: {"vendor": "板金工事", "split": "labor"},

    # 外部建具
    70001: {"vendor": "サッシ", "split": "material"},
    70002: {"vendor": "サッシ", "split": "material"},
    70003: {"vendor": "サッシ", "split": "material"},
    70004: {"vendor": "サッシ", "split": "material"},
    70005: {"vendor": "サッシ", "split": "material"},
    70006: {"vendor": "サッシ", "split": "material"},
    70007: {"vendor": "サッシ", "split": "material"},
    70008: {"vendor": "サッシ", "split": "material"},
    70009: {"vendor": "サッシ", "split": "material"},
    70010: {"vendor": "サッシ", "split": "material"},
    70011: {"vendor": "サッシ", "split": "material"},
    70012: {"vendor": "サッシ", "split": "material"},
    70013: {"vendor": "サッシ", "split": "material"},
    70014: {"vendor": "サッシ", "split": "material"},
    70015: {"vendor": "サッシ", "split": "material"},
    70016: {"vendor": "サッシ", "split": "labor"},
    70017: {"vendor": "サッシ", "split": "material"},

    # 内部建具
    80001: {"vendor": "可動棚（フィットラック）", "split": "material_labor", "labor_ppu": 8000},
    80002: {"vendor": "可動棚（フィットラック）", "split": "material_labor", "labor_ppu": 8000},
    80003: {"vendor": "可動棚（フィットラック）", "split": "material_labor", "labor_ppu": 8000},
    80004: {"vendor": "可動棚（フィットラック）", "split": "material_labor", "labor_ppu": 8000},
    80005: {"vendor": "可動棚（フィットラック）", "split": "material_labor", "labor_ppu": 8000},
    80007: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80008: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80009: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80010: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80011: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80012: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80013: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80014: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80015: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80016: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80017: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80018: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80019: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80020: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80021: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80022: {"vendor": "インテリア建具（LIXIL）", "split": "material"},
    80023: {"vendor": "造作家具", "split": "material_labor", "labor_ratio": 0.3},
    80024: {"vendor": "造作家具", "split": "material_labor", "labor_ratio": 0.3},

    # 外装工事
    90001: {"vendor": "外装工事", "split": "labor"},
    90002: {"vendor": "外装工事", "split": "labor"},
    90003: {"vendor": "外装工事", "split": "labor"},
    90004: {"vendor": "外装工事", "split": "labor"},
    90005: {"vendor": "外装工事", "split": "labor"},

    # 左官工事
    100002: {"vendor": "左官工事", "split": "labor"},
    100003: {"vendor": "左官工事", "split": "material"},
    100004: {"vendor": "左官工事", "split": "material"},
    100005: {"vendor": "左官工事", "split": "material"},
    100006: {"vendor": "左官工事", "split": "labor"},

    # タイル工事
    110001: {"vendor": "タイル工事", "split": "labor"},

    # 塗装工事
    120001: {"vendor": "内部塗装", "split": "labor"},

    # 内装工事
    130001: {"vendor": "畳", "split": "material"},
    130002: {"vendor": "クロス", "split": "labor"},
    130003: {"vendor": "クロス", "split": "labor"},
    130004: {"vendor": "クロス", "split": "labor"},
    130005: {"vendor": "クロス", "split": "labor"},
    130006: {"vendor": "フロアタイル", "split": "labor"},

    # 住宅設備
    140001: {"vendor": "住設（LIXIL/Panasonic等）", "split": "material"},
    140002: {"vendor": "住設（LIXIL/Panasonic等）", "split": "material"},
    140003: {"vendor": "住設", "split": "material"},
    140004: {"vendor": "住設", "split": "material"},
    140005: {"vendor": "住設", "split": "material"},
    140006: {"vendor": "住設（TOTO/LIXIL/Panasonic）", "split": "material"},
    140007: {"vendor": "住設（TOTO/LIXIL/Panasonic）", "split": "material"},
    140008: {"vendor": "住設（LIXIL）", "split": "material"},
    140009: {"vendor": "住設（TOTO）", "split": "material"},
    140010: {"vendor": "住設（Panasonic）", "split": "material"},
    140011: {"vendor": "住設", "split": "material"},
    140012: {"vendor": "住設", "split": "material"},
    140013: {"vendor": "住設", "split": "material"},
    140014: {"vendor": "住設", "split": "material"},
    140015: {"vendor": "住設", "split": "material"},
    140016: {"vendor": "住設（エコキュート）", "split": "material"},
    140017: {"vendor": "住設（エコキュート）", "split": "material"},
    140018: {"vendor": "住設（LIXIL）", "split": "material"},

    # 給排水
    150001: {"vendor": "給排水設備", "split": "labor"},
    150002: {"vendor": "給排水設備", "split": "labor"},
    150003: {"vendor": "給排水設備", "split": "labor"},

    # 電気
    160001: {"vendor": "電気設備", "split": "labor"},
    160002: {"vendor": "電気設備", "split": "labor"},
    160003: {"vendor": "電気設備", "split": "labor"},
    160004: {"vendor": "電気設備", "split": "labor"},
    160005: {"vendor": "電気設備", "split": "labor"},
    160006: {"vendor": "電気設備", "split": "labor"},
    160007: {"vendor": "電気設備", "split": "labor"},
    160008: {"vendor": "電気設備", "split": "labor"},
    160009: {"vendor": "電気設備", "split": "labor"},
    160010: {"vendor": "電気設備", "split": "labor"},
    160011: {"vendor": "電気設備", "split": "labor"},
    160012: {"vendor": "電気設備", "split": "labor"},
    160013: {"vendor": "電気設備", "split": "labor"},
    160014: {"vendor": "電気設備", "split": "labor"},
    160015: {"vendor": "電気設備", "split": "labor"},
    160016: {"vendor": "電気設備", "split": "material_labor", "labor_ppu": 800},
    160017: {"vendor": "電気設備", "split": "labor"},
    160018: {"vendor": "電気設備", "split": "material"},
    160019: {"vendor": "電気設備", "split": "labor"},
    160020: {"vendor": "電気設備", "split": "material"},
    160021: {"vendor": "電気設備", "split": "material"},
    160022: {"vendor": "電気設備", "split": "labor"},
    160023: {"vendor": "電気設備", "split": "labor"},
    160024: {"vendor": "電気設備", "split": "labor"},
    160025: {"vendor": "電気設備", "split": "labor"},
    160026: {"vendor": "HEMS", "split": "material"},
    160027: {"vendor": "HEMS", "split": "labor"},
    160028: {"vendor": "電気設備", "split": "labor"},

    # 換気
    170001: {"vendor": "換気システム", "split": "material"},
    170002: {"vendor": "換気システム", "split": "material"},
    170003: {"vendor": "換気システム", "split": "material"},
    170004: {"vendor": "換気システム", "split": "material"},
    170005: {"vendor": "換気システム", "split": "material"},
    170006: {"vendor": "換気システム", "split": "material"},
    170007: {"vendor": "換気システム", "split": "labor"},
    170008: {"vendor": "換気システム", "split": "labor"},
    170009: {"vendor": "換気システム", "split": "material"},
    170010: {"vendor": "換気システム", "split": "labor"},
    170011: {"vendor": "換気システム", "split": "labor"},
    170012: {"vendor": "換気システム", "split": "material"},
    170013: {"vendor": "換気システム", "split": "material"},
    170014: {"vendor": "換気システム", "split": "material"},
    170015: {"vendor": "換気システム", "split": "material"},
    170016: {"vendor": "換気システム", "split": "material"},

    # エアコン
    180001: {"vendor": "エアコン（富士通）", "split": "material_labor", "labor_ratio": 0.15},
    180002: {"vendor": "エアコン（富士通）", "split": "material_labor", "labor_ratio": 0.15},
    180003: {"vendor": "エアコン（富士通）", "split": "material_labor", "labor_ratio": 0.15},
    180004: {"vendor": "エアコン（富士通）", "split": "material_labor", "labor_ratio": 0.15},
    180006: {"vendor": "エアコン（富士通）", "split": "material"},
    180007: {"vendor": "エアコン工事", "split": "labor"},
    180008: {"vendor": "エアコン工事", "split": "labor"},

    # 太陽光
    190001: {"vendor": "太陽光（長州産業等）", "split": "material_labor", "labor_ratio": 0.2},

    # 造作家具
    200001: {"vendor": "造作家具", "split": "material_labor", "labor_ratio": 0.3},
    200002: {"vendor": "造作家具", "split": "material_labor", "labor_ratio": 0.3},
    200003: {"vendor": "造作家具", "split": "material_labor", "labor_ratio": 0.3},

    # 雑工事
    210001: {"vendor": "清掃工事", "split": "labor"},
    210002: {"vendor": "建材（土橋）", "split": "material_labor", "labor_ppu": 6000},
    210003: {"vendor": "建材（土橋）", "split": "material_labor", "labor_ppu": 6000},
    210004: {"vendor": "建材（土橋）", "split": "material"},
    210005: {"vendor": "建材（土橋）", "split": "material"},

    # 屋外付帯
    220001: {"vendor": "屋外給排水", "split": "labor"},
    220002: {"vendor": "共通仮設", "split": "labor"},
    220003: {"vendor": "ガス工事", "split": "labor"},

    # 諸費用
    230001: {"vendor": "コーディネート", "split": "labor"},
    230002: {"vendor": "設計申請", "split": "labor"},
}


@dataclass
class ANDPADItem:
    """ANDPAD発注用項目"""
    work_category: str      # 工事場所（大項目）
    vendor: str             # 細目工種（発注先）
    item_name: str          # 名称
    summary: str            # 摘要
    quantity: float = 0.0
    unit: str = ""
    material_cost: int = 0  # 材料費
    labor_cost: int = 0     # 施工費
    total_cost: int = 0     # 発注金額


@dataclass
class ANDPADBudget:
    """ANDPAD工務予算"""
    items: list = field(default_factory=list)

    def by_vendor(self) -> dict:
        """発注先別に集計"""
        result = {}
        for item in self.items:
            vendor = item.vendor
            if vendor not in result:
                result[vendor] = {
                    'material': 0,
                    'labor': 0,
                    'total': 0,
                    'items': [],
                }
            result[vendor]['material'] += item.material_cost
            result[vendor]['labor'] += item.labor_cost
            result[vendor]['total'] += item.total_cost
            result[vendor]['items'].append(item)
        return result

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


def convert_to_andpad(estimate: Estimate) -> ANDPADBudget:
    """見積データからANDPAD工務予算データに変換"""
    budget = ANDPADBudget()

    for category in estimate.categories:
        for item in category.items:
            mapping = VENDOR_MAPPING.get(item.child_no)

            if mapping is None:
                # マッピングがない場合はデフォルト
                andpad_item = ANDPADItem(
                    work_category=category.name,
                    vendor=category.name,
                    item_name=item.name,
                    summary=item.summary,
                    quantity=item.quantity,
                    unit=item.unit,
                    material_cost=item.order_amount,
                    labor_cost=0,
                    total_cost=item.order_amount,
                )
                budget.items.append(andpad_item)
                continue

            vendor = mapping['vendor']
            split_type = mapping['split']
            order_amount = item.order_amount

            material = 0
            labor = 0

            if split_type == "material":
                material = order_amount
                labor = 0
            elif split_type == "labor":
                material = 0
                labor = order_amount
            elif split_type == "material_labor":
                if 'labor_ratio' in mapping:
                    labor = int(order_amount * mapping['labor_ratio'])
                    material = order_amount - labor
                elif 'labor_ppu' in mapping:
                    labor = int(mapping['labor_ppu'] * item.quantity)
                    material = order_amount - labor
                else:
                    material = order_amount // 2
                    labor = order_amount - material
            elif split_type == "three_way":
                # クロス/大工手間/建材で3等分
                third = order_amount // 3
                material = third  # 建材分
                labor = third     # 大工手間分
                # 残りはクロス分として材料に含む
                material += order_amount - third * 3 + third
                labor = third

            andpad_item = ANDPADItem(
                work_category=category.name,
                vendor=vendor,
                item_name=item.name,
                summary=item.summary,
                quantity=item.quantity,
                unit=item.unit,
                material_cost=material,
                labor_cost=labor,
                total_cost=order_amount,
            )
            budget.items.append(andpad_item)

    return budget
