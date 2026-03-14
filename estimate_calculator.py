"""見積計算エンジン - CADデータから概算見積を算出"""
import csv
import math
import os
from dataclasses import dataclass, field
from cad_parser import (CADData, FittingData, LIVING_ROOMS, NON_LIVING_ROOMS,
                         WATER_ROOMS, ENTRANCE_ROOMS, BALCONY_ROOMS)


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
        # 管理費
        for c in self.categories:
            if c.no == 22:
                total += c.estimate_total
            elif c.no >= 23:
                total += c.estimate_total
        return total

    @property
    def total_order_excl_tax(self) -> int:
        total = self.subtotal_order
        for c in self.categories:
            if c.no == 22:
                total += c.order_total
            elif c.no >= 23:
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
                       estimate_no: str = "") -> Estimate:
    """CADデータから概算見積を算出"""
    estimate = Estimate(
        property_name=cad_data.property_name,
        owner_name=owner_name,
        location=location,
        estimate_no=estimate_no,
        total_floor_area_m2=cad_data.total_floor_area(),
        total_floor_area_tsubo=cad_data.total_floor_area_tsubo(),
    )

    # 各面積の事前計算
    total_area = cad_data.total_floor_area()
    total_tsubo = cad_data.total_floor_area_tsubo()
    floor1_area = cad_data.floor_area("1階")
    floor2_area = cad_data.floor_area("2階")
    r_floor_area = cad_data.floor_area("R階")
    floor1_tsubo = floor1_area / 3.3057
    floor2_tsubo = floor2_area / 3.3057
    r_floor_tsubo = r_floor_area / 3.3057

    # 特長条件
    cond01 = cad_data.get_condition("R000051")  # 鋼板屋根
    cond02 = cad_data.get_condition("R000052")  # 準防火地域
    cond03 = cad_data.get_condition("R000053")  # IH採用
    cond04 = cad_data.get_condition("R000054")  # 太陽光発電

    # 条件警告
    if cond01 == 1:
        estimate.warnings.append("鋼板屋根（条件01=1）で計算しています。瓦屋根の場合は「見積修正」タブで屋根工事を修正するか、「フィードバック」タブからご連絡ください。")
    if cond04 == 1:
        estimate.warnings.append("太陽光発電システム工事の金額が未入力です。「見積修正」タブの19.太陽光発電システム工事から金額を入力してください。")
    if cond03 == 0:
        estimate.warnings.append("ガス工事の金額が未入力です。「見積修正」タブの23.屋外付帯工事から金額を入力してください。")

    # 水廻り面積
    water_area_tsubo = cad_data.room_area_by_type("全階", WATER_ROOMS) / 3.3057
    # 玄関ポーチ面積
    entrance_area = cad_data.room_area_by_type("全階", ENTRANCE_ROOMS)
    # バルコニー面積
    balcony_area = cad_data.room_area_by_type("全階", BALCONY_ROOMS)
    has_balcony = balcony_area > 0

    # 入隅数
    corner_count = cad_data.sum_quantity("N000080")
    # 外部入隅数（概算：全体入隅の1/4程度）
    ext_corner_count = max(2, int(corner_count * 0.15))

    # 壁面積（外壁見付面積の概算）
    total_wall = cad_data.sum_quantity("N000120")
    # 外壁面積の概算（全壁面積の約1.5倍を見付面積として使用）
    ext_wall_area = total_wall * 0.8

    # 屋根面積の概算（1階床面積 × 1.3〜1.5）
    roof_area = floor1_area * 1.4

    # トイレ数
    toilet_count = cad_data.room_count("全階", ["トイレ", "便所"])
    # 洗面数
    washroom_count = cad_data.room_count("全階", ["洗面", "脱衣"])

    # 収納数
    storage_count = cad_data.room_count("全階",
        ["クローゼット", "WIC", "収納", "押入", "シューズクローク", "SIC", "納戸"])

    # 居室数
    living_room_count = cad_data.room_count("全階", LIVING_ROOMS)

    # 和室判定
    has_tatami = cad_data.room_count("全階", ["和室", "畳"]) > 0

    # 小屋裏判定
    has_attic = cad_data.room_count("全階", ["小屋裏", "ロフト"]) > 0
    attic_tsubo = cad_data.room_area_by_type("全階", ["小屋裏", "ロフト"]) / 3.3057

    # 階段数
    stair_count = max(1, cad_data.room_count("全階", ["階段"]))

    # 階数判定
    is_two_story = floor2_area > 0

    # 外部建具
    ext_fittings = cad_data.get_external_fittings()
    int_fittings = cad_data.get_internal_fittings()

    # ===== 01 仮設工事 =====
    cat01 = WorkCategory(no=1, name="仮設工事")

    # 外部足場
    scaffold_area = ext_wall_area * 1.15
    add_item(cat01, 10001, prices, scaffold_area)

    # 仮設トイレ
    add_item(cat01, 10002, prices, 1)

    # 廃材処分費
    add_item(cat01, 10003, prices, total_area)

    # 運搬費
    add_item(cat01, 10004, prices, total_area)

    # 駐車場代
    add_item(cat01, 10005, prices, 1)

    # 基礎廻り養生シート
    add_item(cat01, 10006, prices, 1)

    # 安全対策費（面積帯別）
    if total_area <= 120:
        add_item(cat01, 10007, prices, 1)
    elif total_area <= 180:
        add_item(cat01, 10008, prices, 1)
    else:
        add_item(cat01, 10009, prices, 1)

    # 整地費用
    add_item(cat01, 10010, prices, 1)

    estimate.categories.append(cat01)

    # ===== 02 基礎工事 =====
    cat02 = WorkCategory(no=2, name="基礎工事")

    add_item(cat02, 20001, prices, 1)  # 水盛り遣方
    add_item(cat02, 20002, prices, floor1_area)  # ベタ基礎

    # 人通口補強
    if floor1_area <= 60:
        add_item(cat02, 20003, prices, 1)
    else:
        add_item(cat02, 20003, prices, 1)
        add_item(cat02, 20004, prices, floor1_area - 60)

    # 下地土間コンクリート
    add_item(cat02, 20005, prices, entrance_area)

    # ポンプ車
    pump_count = 2 if is_two_story else 1
    add_item(cat02, 20006, prices, pump_count)

    estimate.categories.append(cat02)

    # ===== 03 木工事 =====
    cat03 = WorkCategory(no=3, name="木工事")

    # 1階床フローリング
    add_item(cat03, 30001, prices, floor1_tsubo, category="床工事")

    # 2階床フローリング
    if floor2_area > 0:
        add_item(cat03, 30002, prices, floor2_tsubo, category="床工事")

    # R階床フローリング
    if r_floor_area > 0:
        add_item(cat03, 30003, prices, r_floor_tsubo, category="床工事")

    # 水廻り床
    if water_area_tsubo > 0:
        add_item(cat03, 30004, prices, water_area_tsubo, category="床工事")

    # プレカット
    add_item(cat03, 30005, prices, total_tsubo, category="軸組・羽柄材・PC")

    # 金物費捻出（マイナス）
    add_item(cat03, 30006, prices, 1, category="軸組・羽柄材・PC")

    # 割増A：建物難易度
    add_item(cat03, 30007, prices, ext_corner_count, category="軸組・羽柄材・PC")

    # 割増B：2階比
    if is_two_story:
        diff_tsubo = abs(floor1_tsubo - floor2_tsubo)
        add_item(cat03, 30008, prices, diff_tsubo, category="軸組・羽柄材・PC")

    # 金物費
    add_item(cat03, 30010, prices, 1, category="軸組・羽柄材・PC")

    # 構造追加費
    add_item(cat03, 30012, prices, 1, category="軸組・羽柄材・PC")

    # 下地材
    add_item(cat03, 30014, prices, total_tsubo, category="建材・下地材")
    add_item(cat03, 30015, prices, ext_corner_count, category="建材・下地材")

    # 養生費・補修費捻出
    add_item(cat03, 30017, prices, 1, category="建材・下地材")

    # 玄関框
    add_item(cat03, 30018, prices, 1, category="建材・下地材")

    # 階段
    add_item(cat03, 30021, prices, 1, category="建材・下地材")
    add_item(cat03, 30022, prices, 1, category="建材・下地材")

    # 枕棚+パイプ
    closet_count = max(4, storage_count)
    add_item(cat03, 30023, prices, closet_count, category="建材・下地材")

    # 下地補強
    add_item(cat03, 30024, prices, 1, category="建材・下地材")

    # 鋼板屋根遮音下地
    if cond01 == 1:
        add_item(cat03, 30030, prices, roof_area, category="建材・下地材")

    # 外壁透湿防水シート
    add_item(cat03, 30031, prices, ext_wall_area, category="建材・下地材")

    # 耐力面材
    panel_count = math.ceil(ext_wall_area / (0.91 * 2.73))
    add_item(cat03, 30032, prices, panel_count, category="建材・下地材")

    # 制震ダンパー
    add_item(cat03, 30033, prices, 4, category="建材・下地材")

    # 窓枠
    ext_fitting_count = len(ext_fittings)
    add_item(cat03, 30034, prices, ext_fitting_count, category="建材・下地材")

    # 化粧巾木
    baseboard_length = cad_data.sum_quantity("N000130")
    baseboard_count = math.ceil(baseboard_length / 3.64) if baseboard_length > 0 else int(total_tsubo * 1.2)
    add_item(cat03, 30035, prices, baseboard_count, category="建材・下地材")

    # 床養生費
    add_item(cat03, 30044, prices, 1, category="建材・下地材")
    # 補修費
    add_item(cat03, 30045, prices, 1, category="建材・下地材")

    # 換気部材
    eave_length = math.sqrt(floor1_area) * 4 * 0.7
    add_item(cat03, 30038, prices, eave_length, category="建材・下地材")

    # 大工手間 基本
    add_item(cat03, 30101, prices, total_tsubo, category="大工手間")

    # 外周入隅
    add_item(cat03, 30102, prices, ext_corner_count, category="大工手間")

    # ベランダ・ポーチ
    if has_balcony or entrance_area > 0:
        porch_area = balcony_area + entrance_area
        add_item(cat03, 30103, prices, porch_area, category="大工手間")

    # 下屋
    if is_two_story and floor1_area > floor2_area:
        add_item(cat03, 30106, prices, floor1_area - floor2_area, category="大工手間")

    # 外部耐力面材貼り
    add_item(cat03, 30107, prices, total_area, category="大工手間")

    # 中段・枕棚取付
    add_item(cat03, 30110, prices, max(4, storage_count), category="大工手間")

    # クローゼット・収納造作
    extra_storage = max(0, storage_count - 4)
    if extra_storage > 0:
        add_item(cat03, 30111, prices, extra_storage, category="大工手間")

    # 床ガラリ取付
    add_item(cat03, 30115, prices, living_room_count, category="大工手間")

    # 省令準耐火手間
    add_item(cat03, 30117, prices, total_area, category="大工手間")

    # 建て方レッカー費
    add_item(cat03, 30120, prices, 1, category="大工手間")

    # 含み予算
    add_item(cat03, 30121, prices, 1, category="【含み予算】")
    if total_area >= 100:
        add_item(cat03, 30122, prices, total_area - 100, category="【含み予算】")
    add_item(cat03, 30123, prices, 1, category="【含み予算】")
    add_item(cat03, 30124, prices, 1, category="【含み予算】")
    add_item(cat03, 30125, prices, 1, category="【含み予算】")
    add_item(cat03, 30126, prices, 1, category="【含み予算】")
    add_item(cat03, 30127, prices, 1, category="【含み予算】")

    estimate.categories.append(cat03)

    # ===== 04 断熱工事 =====
    cat04 = WorkCategory(no=4, name="断熱工事")

    # 壁断熱材
    wall_insulation_area = ext_wall_area * 1.1
    add_item(cat04, 40001, prices, wall_insulation_area)

    # 天井断熱材
    ceiling_insulation_area = floor1_area if not is_two_story else floor2_area
    add_item(cat04, 40002, prices, ceiling_insulation_area)

    # 2階直下天井吸音材
    if is_two_story:
        add_item(cat04, 40003, prices, floor2_area)

    # 基礎断熱
    foundation_perimeter = math.sqrt(floor1_area) * 4
    foundation_insulation = math.ceil(foundation_perimeter / 1.82)
    add_item(cat04, 40004, prices, foundation_insulation)

    # 気密シート
    add_item(cat04, 40008, prices, 2)

    # 気密テープ
    add_item(cat04, 40009, prices, 3)

    # 気密シート貼り手間
    add_item(cat04, 40010, prices, total_area * 0.6)

    # 気流止め
    add_item(cat04, 40011, prices, total_area * 0.6)

    estimate.categories.append(cat04)

    # ===== 05 屋根工事 =====
    cat05 = WorkCategory(no=5, name="屋根工事")

    if cond01 == 1:  # 鋼板屋根
        add_item(cat05, 50001, prices, roof_area)  # 小屋裏換気(鋼板用)
        add_item(cat05, 50003, prices, 1)           # 荷揚げ費
        add_item(cat05, 50004, prices, roof_area)   # SGL鋼板
    else:  # 瓦屋根
        add_item(cat05, 50002, prices, roof_area)   # 小屋裏換気(瓦用)
        add_item(cat05, 50005, prices, roof_area)   # 平板瓦葺

    estimate.categories.append(cat05)

    # ===== 06 板金・樋工事 =====
    cat06 = WorkCategory(no=6, name="板金・樋工事")

    gutter_length = math.sqrt(floor1_area) * 2 * 1.2
    add_item(cat06, 60001, prices, gutter_length)  # 軒樋
    downspout_count = max(2, int(gutter_length / 15))
    add_item(cat06, 60002, prices, downspout_count)  # 集水器
    downspout_length = downspout_count * (6 if is_two_story else 3.5)
    add_item(cat06, 60003, prices, downspout_length)  # 竪樋

    estimate.categories.append(cat06)

    # ===== 07 外部建具工事 =====
    cat07 = WorkCategory(no=7, name="外部建具工事")

    # 外部建具は個別に計上
    for fitting in ext_fittings:
        count = int(fitting.quantities.get("T100001", 1))
        width = fitting.quantities.get("T100010", 0)
        height = fitting.quantities.get("T100012", 0)

        # 建具タイプに基づいて適切な単価コードを選定
        price_code = match_external_fitting_price(fitting, prices)
        if price_code:
            item = create_item_from_fitting(price_code, prices, count, fitting)
            if item:
                cat07.items.append(item)

    # 建具がマッチしなかった場合のフォールバック
    if not cat07.items:
        # 玄関ドア
        add_item(cat07, 70001, prices, 1)
        # 標準サッシ（概算）
        ext_count = max(1, cad_data.get_fitting_count("金属") - 1)
        add_item(cat07, 70010, prices, ext_count)

    estimate.categories.append(cat07)

    # ===== 08 内部建具工事 =====
    cat08 = WorkCategory(no=8, name="内部建具工事")

    for fitting in int_fittings:
        count = int(fitting.quantities.get("T100001", 1))
        price_code = match_internal_fitting_price(fitting, prices)
        if price_code:
            item = create_item_from_fitting(price_code, prices, count, fitting)
            if item:
                cat08.items.append(item)

    # フォールバック
    if not cat08.items:
        door_count = max(5, cad_data.get_fitting_count("木製"))
        add_item(cat08, 80007, prices, door_count)

    # 可動棚
    if storage_count > 0:
        add_item(cat08, 80001, prices, min(storage_count, 4))

    estimate.categories.append(cat08)

    # ===== 09 外装工事 =====
    cat09 = WorkCategory(no=9, name="外装工事")

    add_item(cat09, 90004, prices, ext_wall_area)   # サイディング
    corner_length = (3.0 if is_two_story else 2.7) * ext_corner_count
    add_item(cat09, 90001, prices, corner_length)    # 出隅コーナー
    add_item(cat09, 90003, prices, ext_wall_area)    # 残材処理
    add_item(cat09, 90002, prices, eave_length)      # 破風・鼻隠し
    # 軒天
    eave_area = eave_length * 0.6
    add_item(cat09, 90005, prices, eave_area)

    estimate.categories.append(cat09)

    # ===== 10 左官工事 =====
    cat10 = WorkCategory(no=10, name="左官工事")

    add_item(cat10, 100002, prices, 1)  # タイル下地
    add_item(cat10, 100003, prices, 1)  # CTバリヤ下地
    add_item(cat10, 100004, prices, 1)  # CTバリヤ本塗
    add_item(cat10, 100005, prices, 1)  # CTバリヤトップ
    # 基礎モルタル
    foundation_wall_area = foundation_perimeter * 0.45
    add_item(cat10, 100006, prices, foundation_wall_area)

    estimate.categories.append(cat10)

    # ===== 11 タイル工事 =====
    cat11 = WorkCategory(no=11, name="タイル工事")
    add_item(cat11, 110001, prices, entrance_area)
    estimate.categories.append(cat11)

    # ===== 12 塗装工事 =====
    cat12 = WorkCategory(no=12, name="塗装工事")
    # 和室がある場合や木部仕上げがある場合
    if has_tatami or cad_data.room_count("全階", ["和室"]) > 0:
        add_item(cat12, 120001, prices, 1)
    estimate.categories.append(cat12)

    # ===== 13 内装工事 =====
    cat13 = WorkCategory(no=13, name="内装工事")

    # クロス工事（一般部）
    add_item(cat13, 130002, prices, total_tsubo)

    # 小屋裏クロス
    if has_attic:
        add_item(cat13, 130003, prices, attic_tsubo)

    # 収納追加クロス
    extra_closet = max(0, storage_count - 6)
    if extra_closet > 0:
        add_item(cat13, 130004, prices, extra_closet)

    # 畳
    if has_tatami:
        tatami_area = cad_data.room_area_by_type("全階", ["和室", "畳"])
        tatami_count = math.ceil(tatami_area / 0.81)  # 半帖換算
        add_item(cat13, 130001, prices, tatami_count)

    estimate.categories.append(cat13)

    # ===== 14 住宅設備機器工事 =====
    cat14 = WorkCategory(no=14, name="住宅設備機器工事")

    add_item(cat14, 140001, prices, 1)  # キッチン

    # UB (面積に応じてサイズ選定)
    bath_area = cad_data.room_area_by_type("全階", ["浴室"])
    if bath_area >= 3.3:
        add_item(cat14, 140006, prices, 1)  # UB 1620
    else:
        add_item(cat14, 140007, prices, 1)  # UB 1616

    # トイレ
    if toilet_count >= 1:
        add_item(cat14, 140008, prices, 1)  # タンクレストイレ
    if toilet_count >= 2:
        add_item(cat14, 140009, prices, toilet_count - 1)  # 2台目以降

    # 造作洗面
    add_item(cat14, 140011, prices, 1)

    # タオル掛け・ペーパーホルダー
    add_item(cat14, 140013, prices, toilet_count)
    add_item(cat14, 140014, prices, toilet_count)

    # 洗濯パン
    add_item(cat14, 140015, prices, 1)

    # エコキュート
    if total_area >= 120:
        add_item(cat14, 140016, prices, 1)  # 460L
    else:
        add_item(cat14, 140017, prices, 1)  # 370L

    estimate.categories.append(cat14)

    # ===== 15 給排水設備工事 =====
    cat15 = WorkCategory(no=15, name="給排水設備工事")

    add_item(cat15, 150001, prices, 1)  # 基本
    # 追加配管
    extra_plumbing = max(0, (toilet_count + washroom_count + 1) - 6)
    if extra_plumbing > 0:
        add_item(cat15, 150002, prices, extra_plumbing)

    estimate.categories.append(cat15)

    # ===== 16 電気設備工事 =====
    cat16 = WorkCategory(no=16, name="電気設備工事")

    add_item(cat16, 160001, prices, 1)  # 仮設電気

    # スイッチ・コンセント（居室数ベースで概算）
    switch_count = living_room_count + toilet_count + washroom_count + 3
    add_item(cat16, 160002, prices, switch_count)  # 片切スイッチ
    add_item(cat16, 160003, prices, max(3, living_room_count))  # 3路4路

    # 電灯配線
    light_count = living_room_count * 3 + toilet_count + washroom_count + 5
    add_item(cat16, 160004, prices, light_count)

    # コンセント
    outlet_count = living_room_count * 3 + toilet_count + washroom_count + 6
    add_item(cat16, 160005, prices, outlet_count)

    # E付コンセント
    add_item(cat16, 160006, prices, toilet_count + 4)  # UB等
    add_item(cat16, 160007, prices, 2)  # レンジ・食洗機
    if cond03 == 1:  # IH
        add_item(cat16, 160008, prices, 1)

    # エアコン用コンセント
    ac_count = living_room_count + (1 if is_two_story else 0)
    add_item(cat16, 160010, prices, ac_count)

    # TV配線
    add_item(cat16, 160011, prices, max(2, living_room_count))

    # 防水コンセント
    add_item(cat16, 160013, prices, 2)

    # 火災警報器
    alarm_count = living_room_count + 2
    add_item(cat16, 160016, prices, alarm_count)
    add_item(cat16, 160017, prices, alarm_count)

    # 幹線・分電盤
    add_item(cat16, 160023, prices, 1)

    # 太陽光対応
    if cond04 == 1:
        add_item(cat16, 160024, prices, 1)
        add_item(cat16, 160026, prices, 1)  # HEMS
        add_item(cat16, 160027, prices, 1)

    add_item(cat16, 160028, prices, 1)  # 電力会社申請

    estimate.categories.append(cat16)

    # ===== 17 換気システム工事 =====
    cat17 = WorkCategory(no=17, name="換気システム工事")

    # 面積帯別換気本体
    if total_area <= 63:
        add_item(cat17, 170001, prices, 1)
    elif total_area <= 103:
        add_item(cat17, 170002, prices, 1)
    elif total_area <= 123:
        add_item(cat17, 170003, prices, 1)
    elif total_area <= 162:
        add_item(cat17, 170004, prices, 1)
    else:
        add_item(cat17, 170005, prices, 1)

    # ダクト材
    add_item(cat17, 170006, prices, total_tsubo)

    # 施工費
    add_item(cat17, 170007, prices, 1)

    # 電源工事
    vent_unit_count = 2 if total_area > 123 else 1
    add_item(cat17, 170008, prices, vent_unit_count)

    # 局所換気
    local_vent_count = toilet_count + 1
    add_item(cat17, 170009, prices, local_vent_count)
    add_item(cat17, 170010, prices, local_vent_count)
    add_item(cat17, 170011, prices, local_vent_count)

    # 屋外フード
    hood_count = local_vent_count + vent_unit_count * 2
    if cond02 == 1:  # 準防火
        add_item(cat17, 170014, prices, hood_count)
    else:
        add_item(cat17, 170013, prices, hood_count)

    estimate.categories.append(cat17)

    # ===== 18 エアコン工事 =====
    cat18 = WorkCategory(no=18, name="エアコン工事")

    # LDK用（大型）
    ldk_count = cad_data.room_count("全階", ["LDK", "居間", "リビング"])
    if ldk_count > 0:
        ldk_area = cad_data.room_area_by_type("全階", ["LDK", "居間", "リビング"])
        if ldk_area >= 30:
            add_item(cat18, 180001, prices, 1)  # 5.6kw
        else:
            add_item(cat18, 180002, prices, 1)  # 4.0kw

    # 寝室・子供室用
    bedroom_count = cad_data.room_count("全階",
        ["主寝室", "寝室", "子供室", "洋室", "書斎"])
    for _ in range(bedroom_count):
        add_item(cat18, 180004, prices, 1)  # 2.2kw

    # 無線LANアダプター
    total_ac = ldk_count + bedroom_count
    add_item(cat18, 180006, prices, total_ac)

    estimate.categories.append(cat18)

    # ===== 19 太陽光発電システム工事 =====
    cat19 = WorkCategory(no=19, name="太陽光発電システム工事")
    if cond04 == 1:
        add_item(cat19, 190001, prices, 1)
    estimate.categories.append(cat19)

    # ===== 20 造作家具工事 =====
    cat20 = WorkCategory(no=20, name="造作家具工事")
    add_item(cat20, 200001, prices, 1)  # カウンター
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
    if cond03 == 0:  # ガス
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


def match_external_fitting_price(fitting: FittingData, prices: dict) -> int:
    """外部建具をマスタとマッチング"""
    code = fitting.code
    width = fitting.quantities.get("T100010", 0)
    height = fitting.quantities.get("T100012", 0)

    # コードからサイズ情報を抽出
    size_str = ""
    parts = code.split('-')
    if len(parts) >= 3:
        w_mm = int(width * 1000) if width > 0 else 0
        h_mm = int(height * 100) if height > 0 else 0

    # 玄関ドア判定
    if "片開" in fitting.type_name and height > 2.0:
        if "LIXIL" in str(fitting.material) or "アルミ" in fitting.material:
            return 70001
        return 70002

    # 引違い窓判定
    if "引違" in fitting.type_name:
        w_code = int(width * 100)
        h_code = int(height * 100)
        # サイズに基づくマッチング
        if w_code >= 250:
            return 70015
        elif w_code >= 160 and h_code >= 20:
            return 70012
        elif w_code >= 160 and h_code >= 9:
            return 70010
        elif w_code >= 160:
            return 70009
        else:
            return 70010

    # FIX窓
    if "Fix" in fitting.type_name:
        return 70007

    # デフォルト
    return 70010


def match_internal_fitting_price(fitting: FittingData, prices: dict) -> int:
    """内部建具をマスタとマッチング"""
    code = fitting.code
    width = fitting.quantities.get("T100010", 0)
    height = fitting.quantities.get("T100012", 0)

    # 折戸
    if "折戸" in fitting.type_name or "4枚折戸" in fitting.type_name:
        if width >= 1.5:
            return 80014  # W6尺
        return 80013  # W3尺

    # 片引き
    if "片引き" in fitting.type_name:
        return 80011

    # 片開き
    if "片開" in fitting.type_name:
        if height >= 2.1:
            return 80007  # 標準ドア
        return 80007

    return 80007


def create_item_from_fitting(price_code: int, prices: dict,
                              count: int, fitting: FittingData) -> EstimateItem:
    """建具から見積項目を作成"""
    if price_code not in prices:
        return None

    p = prices[price_code]
    return EstimateItem(
        parent_no=0,
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
