"""CAD数量データ（ARCHITREND ZERO）パーサー"""
import codecs
import re
import io
from dataclasses import dataclass, field


@dataclass
class RoomData:
    floor: str  # "1階", "2階", "R階" etc.
    name: str   # "LDK", "トイレ" etc.
    quantities: dict = field(default_factory=dict)  # code -> value


@dataclass
class FittingData:
    category: str    # "金属/戸", "木製/戸" etc.
    type_name: str   # "片開", "引違い" etc.
    material: str    # "アルミ製", "樹脂製" etc.
    code: str        # "AD-1-08823" etc.
    quantities: dict = field(default_factory=dict)


@dataclass
class PropertyCondition:
    code: str
    name: str
    value: float


@dataclass
class CADData:
    property_number: str = ""
    property_name: str = ""
    drawing_name: str = ""
    designer: str = ""
    date: str = ""
    structure: str = ""
    rooms: list = field(default_factory=list)
    fittings: list = field(default_factory=list)
    conditions: list = field(default_factory=list)

    def get_condition(self, code: str) -> float:
        for c in self.conditions:
            if c.code == code:
                return c.value
        return 0.0

    def total_floor_area(self) -> float:
        """全階の床面積合計"""
        total = 0.0
        for room in self.rooms:
            total += room.quantities.get("N000001", 0.0)
        return total

    def floor_area(self, floor: str) -> float:
        """指定階の床面積合計"""
        total = 0.0
        for room in self.rooms:
            if room.floor == floor:
                total += room.quantities.get("N000001", 0.0)
        return total

    def room_area_by_type(self, floor: str, room_types: list) -> float:
        """指定階の指定タイプの部屋面積合計"""
        total = 0.0
        for room in self.rooms:
            if floor != "全階" and room.floor != floor:
                continue
            for rt in room_types:
                if rt in room.name:
                    total += room.quantities.get("N000001", 0.0)
                    break
        return total

    def room_count(self, floor: str, room_types: list) -> int:
        """指定階の指定タイプの部屋数"""
        count = 0
        for room in self.rooms:
            if floor != "全階" and room.floor != floor:
                continue
            for rt in room_types:
                if rt in room.name:
                    count += 1
                    break
        return count

    def sum_quantity(self, code: str, floor: str = "全階", room_types: list = None) -> float:
        """数量コードの合計を取得"""
        total = 0.0
        for room in self.rooms:
            if floor != "全階" and room.floor != floor:
                continue
            if room_types:
                match = False
                for rt in room_types:
                    if rt in room.name:
                        match = True
                        break
                if not match:
                    continue
            total += room.quantities.get(code, 0.0)
        return total

    def get_fitting_count(self, fitting_type: str = None) -> int:
        """建具の数を取得"""
        count = 0
        for f in self.fittings:
            if fitting_type is None or fitting_type in f.category:
                count += int(f.quantities.get("T100001", 0))
        return count

    def get_external_fittings(self) -> list:
        """外部建具を取得"""
        return [f for f in self.fittings if "金属" in f.category]

    def get_internal_fittings(self) -> list:
        """内部建具を取得"""
        return [f for f in self.fittings if "木製" in f.category]

    def total_floor_area_tsubo(self) -> float:
        """延床面積（坪）"""
        return self.total_floor_area() / 3.3057

    def total_wall_area(self) -> float:
        """壁仕上面積合計"""
        return self.sum_quantity("N000120")

    def total_ceiling_area(self) -> float:
        """天井面積合計"""
        return self.sum_quantity("N000003")

    def get_floor_names(self) -> list:
        """存在する階名を取得"""
        floors = set()
        for room in self.rooms:
            floors.add(room.floor)
        order = {"1階": 1, "2階": 2, "3階": 3, "R階": 4}
        return sorted(floors, key=lambda x: order.get(x, 99))


# 居室判定用の部屋タイプリスト
LIVING_ROOMS = ["LDK", "居間", "台所", "食堂", "主寝室", "子供室", "洋室", "和室",
                "書斎", "趣味室", "寝室", "リビング", "ダイニング", "キッチン"]
NON_LIVING_ROOMS = ["玄関", "ホール", "廊下", "階段室", "階段", "便所", "トイレ",
                     "浴室", "脱衣室", "脱衣", "洗面所", "洗面", "収納", "押入",
                     "クローゼット", "WIC", "車庫", "インナーガレージ", "ガレージ",
                     "小屋裏", "ロフト", "納戸", "ユーティリティ", "パントリー",
                     "SIC", "シューズクローク"]
WATER_ROOMS = ["便所", "トイレ", "浴室", "脱衣室", "脱衣", "洗面所", "洗面",
               "ユーティリティ"]
ENTRANCE_ROOMS = ["玄関", "ポーチ"]
BALCONY_ROOMS = ["バルコニー", "ベランダ", "テラス"]


def parse_cad_data(content: str) -> CADData:
    """CAD数量データをパースする"""
    data = CADData()
    lines = content.strip().split('\n')

    if not lines:
        return data

    # ヘッダー行のパース
    header = parse_csv_line(lines[0])
    if len(header) >= 6:
        data.property_number = header[0]
        data.property_name = header[1]
        data.drawing_name = header[2]
        data.designer = header[3]
        data.date = header[4]
        data.structure = header[5]

    current_section = None
    current_room = None
    current_fitting = None
    i = 1

    while i < len(lines):
        line = lines[i].strip()

        if not line:
            i += 1
            continue

        # セクション判定
        if line.startswith('"部屋データ"') or line == '部屋データ':
            current_section = "rooms"
            i += 1
            continue
        elif line.startswith('"建具集計マスタ"') or line == '建具集計マスタ':
            current_section = "fittings"
            if current_room:
                data.rooms.append(current_room)
                current_room = None
            i += 1
            continue
        elif line.startswith('"物件条件"') or line == '物件条件':
            current_section = "conditions"
            if current_fitting:
                data.fittings.append(current_fitting)
                current_fitting = None
            i += 1
            continue

        if current_section == "rooms":
            # 部屋ヘッダーの判定
            room_match = re.match(r'^"?(\d+階|R階|PH階)-<(.+?)>"?$', line)
            if room_match:
                if current_room:
                    data.rooms.append(current_room)
                current_room = RoomData(
                    floor=room_match.group(1),
                    name=room_match.group(2),
                )
                i += 1
                continue

            # 数量行のパース
            if current_room and line.startswith('"'):
                parts = parse_csv_line(line)
                if len(parts) >= 3:
                    code = parts[0]
                    try:
                        value = float(parts[2])
                        current_room.quantities[code] = value
                    except (ValueError, IndexError):
                        pass

        elif current_section == "fittings":
            # 建具ヘッダーの判定
            fitting_match = re.match(r'^"?(.+?)/(.+?)/(.+?)-(.+)"?$', line)
            if fitting_match:
                if current_fitting:
                    data.fittings.append(current_fitting)
                current_fitting = FittingData(
                    category=f"{fitting_match.group(1)}/{fitting_match.group(2)}",
                    type_name=fitting_match.group(4).split('-')[0] if '-' in fitting_match.group(4) else fitting_match.group(4),
                    material=fitting_match.group(3),
                    code=fitting_match.group(4),
                )
                i += 1
                continue

            # 建具数量行
            if current_fitting and line.startswith('"'):
                parts = parse_csv_line(line)
                if len(parts) >= 3:
                    code = parts[0]
                    try:
                        value = float(parts[2])
                        current_fitting.quantities[code] = value
                    except (ValueError, IndexError):
                        pass

        elif current_section == "conditions":
            if line.startswith('"'):
                parts = parse_csv_line(line)
                if len(parts) >= 3:
                    code = parts[0]
                    name = parts[1]
                    try:
                        value = float(parts[2])
                        data.conditions.append(PropertyCondition(
                            code=code, name=name, value=value
                        ))
                    except (ValueError, IndexError):
                        pass

        i += 1

    # 最後のデータを追加
    if current_room:
        data.rooms.append(current_room)
    if current_fitting:
        data.fittings.append(current_fitting)

    return data


def parse_csv_line(line: str) -> list:
    """CSVライクな行をパース（ダブルクォート対応）"""
    result = []
    current = ""
    in_quotes = False

    for char in line:
        if char == '"':
            in_quotes = not in_quotes
        elif char == ',' and not in_quotes:
            result.append(current.strip())
            current = ""
        else:
            current += char

    result.append(current.strip())
    return result


def load_cad_file(file_path: str) -> CADData:
    """ファイルパスからCADデータを読み込む"""
    with codecs.open(file_path, 'r', encoding='cp932') as f:
        content = f.read()
    return parse_cad_data(content)


def load_cad_from_bytes(file_bytes: bytes) -> CADData:
    """バイト列からCADデータを読み込む（Streamlitのファイルアップロード用）"""
    # CP932でデコード試行、失敗したらUTF-8
    try:
        content = file_bytes.decode('cp932')
    except UnicodeDecodeError:
        try:
            content = file_bytes.decode('utf-8')
        except UnicodeDecodeError:
            content = file_bytes.decode('shift_jis', errors='replace')
    return parse_cad_data(content)
