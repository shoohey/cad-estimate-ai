"""PDF図面からの概算見積 (検証用モジュール / CAD→概算見積もり生成AI)

AIがPDF図面から読み取った数量 (ai_client.extract_quantities_from_pdf の戻り値) を
cad_parser.CADData に組み立て、既存の見積計算エンジン
(estimate_calculator.calculate_estimate) をそのまま通して概算見積を作る。
CAD数量データ (TXT) 由来の見積と並べて「PDFだけでどこまで精度が出るか」を検証する用途。

PDFから読めない集計値 (施工床・軒樋長 など) は 延床面積 × 比率 で推定する。
比率は calibration/ の8物件 (CAD実測値) から derive_ratios_from_calibration() で算出した
平均値を定数 RATIO_TO_TOTAL_FLOOR_AREA に埋め込んでいる。推定した項目は必ず
CADData.parse_warnings に記録する (見積の warnings にもそのまま引き継がれる)。

依存: cad_parser / estimate_calculator / ai_client.parse_number
"""

from __future__ import annotations

import glob
import os
import re
import statistics
from typing import Any

from ai_client import parse_number
from cad_parser import (CADData, FittingData, PropertyCondition, RoomData,
                        load_cad_file, _norm)
from estimate_calculator import Estimate, calculate_estimate

# --- 定数 ---------------------------------------------------------------

CALIB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "calibration")
CALIB_GLOB = "*【数量データ】*.TXT"

# 延床面積からの比率で推定する集計コード (code -> 日本語名)
ESTIMATED_AGG_NAMES = {
    "H000402": "建築面積",
    "B000581": "外周長",
    "Y000181": "屋根面積",
    "Y000540": "軒樋長",
    "F000140": "外壁面積",
    "H000405": "施工床面積",
}

# 延床面積に対する各集計値の比率 (値 ÷ 延床面積 H000403)。
#
# 由来: calibration/*【数量データ】*.TXT 8物件 (1階建 3 / 2階建 5) を
#       cad_parser.load_cad_file で読み、derive_ratios_from_calibration() で
#       算出した平均値 (2026-10-26 実測、小数4桁)。
#       建築面積・外周長・屋根面積・外壁面積は階数で大きく変わるため
#       1階建 / 2階建 (3階建以上は2階建を流用) で分けている。
#       軒樋長 (Y000540) は8物件中1物件しか出力されていないため、
#       他の物件は軒先長 (Y000151) で代用した (estimate_calculator も
#       軒樋長が無い場合は軒先長で代用している)。
# 更新: calibration 物件を追加したら derive_ratios_from_calibration() を実行して差し替える。
RATIO_TO_TOTAL_FLOOR_AREA: dict[int, dict[str, float]] = {
    1: {
        "H000402": 1.0291,
        "B000581": 0.4500,
        "Y000181": 1.2783,
        "Y000540": 0.2251,
        "F000140": 1.6031,
        "H000405": 1.0584,
    },
    2: {
        "H000402": 0.6129,
        "B000581": 0.6108,
        "Y000181": 0.8461,
        "Y000540": 0.2359,
        "F000140": 1.9158,
        "H000405": 1.0524,
    },
}

# 2階建で階別面積が読めない場合の 1階床面積 ÷ 延床面積 (同上・2階建5物件の平均)
FLOOR1_SHARE_TWO_STORY = 0.5714

# 玄関ドアの寸法が読めない場合の標準寸法 (m)。種類の判定 (玄関ドア=電気錠付セット) が
# 単価マッチングに効くため、寸法未取得でも玄関ドアとして計上されるようにする
DEFAULT_ENTRANCE_DOOR_W_M = 0.9
DEFAULT_ENTRANCE_DOOR_H_M = 2.3

# 延床面積と階別面積合計・部屋面積合計の食い違いを警告する閾値
AREA_MISMATCH_TOLERANCE = 0.05   # 階別面積合計 (5%)
ROOM_SUM_MISMATCH_TOLERANCE = 0.15   # 部屋面積合計 (15%: バルコニー等の扱い差を許容)


# --- 内部ヘルパー: 値の正規化 -------------------------------------------

_DIGITS_RE = re.compile(r"\d+")
_KANJI_FLOOR = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5"}


def _text(value: Any) -> str:
    """文字列に正規化する (None / bool → 空文字、全角英数は半角に)。"""
    if value is None or isinstance(value, bool):
        return ""
    return _norm(str(value))


def _positive(value: Any, label: str, warnings: list[str]) -> float | None:
    """正の数値を返す。値があるのに読めない / 0以下なら警告を残して None。"""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    num = parse_number(value)
    if num is None:
        warnings.append(f"{label}「{value}」を数値として読めないため未取得として扱いました")
        return None
    if num <= 0:
        warnings.append(f"{label}の値 {num:g} が0以下のため未取得として扱いました")
        return None
    return num


def normalize_floor_label(value: Any) -> str | None:
    """階の表記を cad_parser の階名 ("1階" / "2階" / "R階" / "PH階") に揃える。

    "1F" / "１階" / "一階" / 1 → "1階"、"RF" / "屋上" → "R階"、"PH" → "PH階"。
    読めない場合は None。
    """
    if value is None or isinstance(value, bool):
        return None
    text = _norm(str(value))
    if not text:
        return None
    for kanji, digit in _KANJI_FLOOR.items():
        text = text.replace(kanji, digit)
    upper = text.upper()
    if upper.startswith("PH") or "ペントハウス" in text:
        return "PH階"
    if upper.startswith("R") or "屋上" in text:
        return "R階"
    m = _DIGITS_RE.search(text)
    if m:
        n = int(m.group(0))
        if upper.startswith("B") or "地下" in text:
            return f"B{n}階"
        return f"{n}階"
    return None


# --- 内部ヘルパー: 建具の分類 -------------------------------------------

_ENTRANCE_WORDS = ("玄関",)
_BACK_DOOR_WORDS = ("勝手口",)
_INTERNAL_WORDS = ("内部", "室内", "内装", "木製")
_WINDOW_TYPES = (
    # (キーワード群, cad_parser の type_name 相当)
    (("引違", "引き違", "掃き出", "掃出", "テラス"), "引違い"),
    (("fix", "嵌め殺", "嵌殺", "はめ殺"), "Fix"),
    (("縦すべり", "縦滑り"), "縦すべり"),
    (("横すべり", "横滑り", "すべり出し", "滑り出し"), "横すべり"),
    (("上げ下げ", "上下"), "上げ下げ"),
    (("突き出し", "突出し"), "突出し"),
    (("ルーバー", "ジャロジー"), "ルーバー"),
    (("片開",), "片開"),
)
_DOOR_TYPES = (
    (("親子",), "親子"),
    (("両開",), "両開"),
    (("折戸", "折れ戸"), "折戸"),
    (("引込", "引き込"), "引込み"),
    (("片引", "吊戸", "吊り戸", "上吊"), "片引き"),
    (("引違", "引き違"), "引違い"),
    (("片開", "開き", "ドア"), "片開"),
)


def _match_type(text: str, table) -> str | None:
    lowered = text.lower()
    for keywords, type_name in table:
        if any(k.lower() in lowered for k in keywords):
            return type_name
    return None


def classify_fitting(kind: Any, name: Any) -> tuple[str, str, str]:
    """PDF抽出の建具 (kind, name) を (category, type_name, material) に分類する。

    cad_parser.FittingData の表記に合わせる:
      外部窓     → ("金属/窓", 引違い/Fix/縦すべり/..., "樹脂製")
      玄関ドア   → ("金属/戸", 片開 or 親子, "アルミ製")
      勝手口     → ("金属/戸", "勝手口", "アルミ製")  ※標準サッシ相当で概算
      内部建具   → ("木製/戸" or 室内窓は "木製/窓", 片開/片引き/折戸/..., "木製")
    """
    text = f"{_text(kind)} {_text(name)}".strip()
    lowered = text.lower()
    if any(w in text for w in _ENTRANCE_WORDS):
        return "金属/戸", ("親子" if "親子" in text else "片開"), "アルミ製"
    if any(w in text for w in _BACK_DOOR_WORDS):
        return "金属/戸", "勝手口", "アルミ製"
    if any(w in text for w in _INTERNAL_WORDS):
        if "窓" in text:
            return "木製/窓", (_match_type(text, _WINDOW_TYPES) or "Fix"), "木製"
        return "木製/戸", (_match_type(text, _DOOR_TYPES) or "片開"), "木製"
    if "窓" in text or "サッシ" in text or "fix" in lowered or "外部" in text:
        return "金属/窓", (_match_type(text, _WINDOW_TYPES) or (_text(name) or "窓")), "樹脂製"
    # kind が不明で名称が戸っぽいものは内部建具扱い (外部ドアは玄関/勝手口で拾う)
    if any(w in text for w in ("戸", "ドア", "扉")):
        return "木製/戸", (_match_type(text, _DOOR_TYPES) or "片開"), "木製"
    return "金属/窓", (_match_type(text, _WINDOW_TYPES) or (_text(name) or "窓")), "樹脂製"


# --- 比率の算出 (calibration 実測) --------------------------------------

def derive_ratios_from_calibration(calib_dir: str | None = None) -> dict[str, Any]:
    """calibration/ のCAD数量データから「集計値 ÷ 延床面積」の比率を実測する。

    RATIO_TO_TOTAL_FLOOR_AREA / FLOOR1_SHARE_TWO_STORY の由来を再現するための関数。
    テストでも呼び、埋め込み定数が実測と乖離していないことを確認する。

    Returns:
        {
          "n_properties": int,
          "files": [ファイル名, ...],
          "by_stories": {1: {code: {"mean","min","max","n"}}, 2: {...}},
          "floor1_share_two_story": {"mean","min","max","n"},
        }

    Raises:
        FileNotFoundError: calibration ディレクトリに数量データが無い場合。
    """
    calib_dir = calib_dir or CALIB_DIR
    files = sorted(glob.glob(os.path.join(calib_dir, CALIB_GLOB)))
    if not files:
        raise FileNotFoundError(f"calibration の数量データが見つかりません: {calib_dir}")

    samples: dict[int, dict[str, list[float]]] = {
        1: {c: [] for c in ESTIMATED_AGG_NAMES}, 2: {c: [] for c in ESTIMATED_AGG_NAMES}}
    floor1_shares: list[float] = []
    for path in files:
        cad = load_cad_file(path)
        tfa = cad.total_floor_area()
        if tfa <= 0:
            continue
        two_story = cad.agg("B001212") > 0
        stories = 2 if two_story else 1
        if two_story:
            floor1_shares.append(cad.agg("B001211") / tfa)
        for code in ESTIMATED_AGG_NAMES:
            value = cad.agg(code)
            if code == "Y000540" and value <= 0:
                value = cad.agg("Y000151")   # 軒樋長が無い物件は軒先長で代用
            if value > 0:
                samples[stories][code].append(value / tfa)

    def _stats(values: list[float]) -> dict[str, float | int]:
        if not values:
            return {"mean": 0.0, "min": 0.0, "max": 0.0, "n": 0}
        return {"mean": round(statistics.mean(values), 4), "min": round(min(values), 4),
                "max": round(max(values), 4), "n": len(values)}

    return {
        "n_properties": len(files),
        "files": [os.path.basename(f) for f in files],
        "by_stories": {s: {c: _stats(v) for c, v in codes.items()} for s, codes in samples.items()},
        "floor1_share_two_story": _stats(floor1_shares),
    }


# --- CADData の組み立て -------------------------------------------------

def _resolve_stories(extracted: dict, floor_areas: dict[str, float],
                     rooms: list[RoomData], warnings: list[str]) -> int:
    """階数を決める: floors → 階別面積のキー → 部屋の階 → 既定 (2階建・警告)。"""
    floors = parse_number(extracted.get("floors"))
    if floors is not None and floors >= 1:
        return int(floors)
    numbered = [int(m.group(0)) for label in list(floor_areas) + [r.floor for r in rooms]
                for m in [_DIGITS_RE.search(label or "")] if m]
    if numbered:
        return max(numbered)
    warnings.append("階数はPDFから読めず2階建と仮定しました")
    return 2


def _build_rooms(extracted: dict, warnings: list[str]) -> list[RoomData]:
    raw_rooms = extracted.get("rooms")
    if raw_rooms is None:
        return []
    if not isinstance(raw_rooms, list):
        warnings.append("部屋一覧がリスト形式でないため部屋データ無しとして扱いました")
        return []
    rooms: list[RoomData] = []
    skipped = 0
    no_floor = 0
    no_area = 0
    for entry in raw_rooms:
        if not isinstance(entry, dict):
            skipped += 1
            continue
        name = _text(entry.get("name"))
        if not name:
            skipped += 1
            continue
        floor = normalize_floor_label(entry.get("floor"))
        if floor is None:
            floor = "1階"
            no_floor += 1
        area = _positive(entry.get("area_m2"), f"部屋「{name}」の面積", warnings)
        if area is None:
            area = 0.0
            no_area += 1
        # N000001=床面積 / N000100=部屋面積 (estimate_calculator が参照する2コード)
        rooms.append(RoomData(floor=floor, name=name,
                              quantities={"N000001": area, "N000100": area}))
    if skipped:
        warnings.append(f"部屋一覧のうち{skipped}件は名称が無い・形式不正のため除外しました")
    if no_floor:
        warnings.append(f"部屋{no_floor}室は階がPDFから読めず1階として扱いました")
    if no_area:
        warnings.append(f"部屋{no_area}室は面積がPDFから読めず0㎡として扱いました (部屋数の計上のみ)")
    return rooms


def _build_fittings(extracted: dict, warnings: list[str]) -> list[FittingData]:
    raw = extracted.get("fittings")
    if raw is None:
        return []
    if not isinstance(raw, list):
        warnings.append("建具一覧がリスト形式でないため建具無しとして扱いました")
        return []
    fittings: list[FittingData] = []
    skipped = 0
    no_dims = 0
    for idx, entry in enumerate(raw, start=1):
        if not isinstance(entry, dict):
            skipped += 1
            continue
        kind, name = entry.get("kind"), entry.get("name")
        if not (_text(kind) or _text(name)):
            skipped += 1
            continue
        category, type_name, material = classify_fitting(kind, name)
        label = f"建具{idx}「{_text(name) or _text(kind)}」"
        count = parse_number(entry.get("count"))
        if count is None or count < 1:
            if entry.get("count") is not None:
                warnings.append(f"{label}の数量「{entry.get('count')}」が不正のため1として扱いました")
            count = 1
        width_mm = _positive(entry.get("width_mm"), f"{label}の幅", warnings)
        height_mm = _positive(entry.get("height_mm"), f"{label}の高さ", warnings)
        # CADの建具寸法 (T100010 幅 / T100012 高さ) は m 単位
        quantities: dict[str, float] = {"T100001": float(int(count))}
        if width_mm is not None:
            quantities["T100010"] = width_mm / 1000.0
        if height_mm is not None:
            quantities["T100012"] = height_mm / 1000.0
        is_entrance = category == "金属/戸" and type_name in ("片開", "親子")
        if is_entrance:
            if "T100010" not in quantities:
                quantities["T100010"] = DEFAULT_ENTRANCE_DOOR_W_M
            if "T100012" not in quantities:
                quantities["T100012"] = DEFAULT_ENTRANCE_DOOR_H_M
                warnings.append(
                    f"{label}の高さがPDFから読めず標準{DEFAULT_ENTRANCE_DOOR_H_M}mとして玄関ドアで計上しました")
        elif width_mm is None or height_mm is None:
            no_dims += 1
        fittings.append(FittingData(
            category=category, type_name=type_name, material=material,
            code=f"{type_name}-PDF-{idx}", quantities=quantities))
    if skipped:
        warnings.append(f"建具一覧のうち{skipped}件は種類・名称が無い・形式不正のため除外しました")
    if no_dims:
        warnings.append(
            f"建具{no_dims}件は寸法がPDFから読めず最小サイズ帯の単価で計上しました")
    return fittings


def _build_conditions(extracted: dict, warnings: list[str]) -> list[PropertyCondition]:
    """屋根の記載から物件条件 (R000051 鋼板屋根 / R000031 切妻) を立てる。"""
    roof = extracted.get("roof") if isinstance(extracted.get("roof"), dict) else {}
    roof_type = _text(roof.get("type"))
    conditions: list[PropertyCondition] = []
    if not roof_type:
        warnings.append("屋根の種類がPDFから読めず瓦屋根 (条件01=0) で計算しました")
        return conditions
    metal = any(w in roof_type for w in ("ガルバ", "鋼板", "金属", "板金", "スレート", "コロニアル"))
    tile = "瓦" in roof_type
    if metal and not tile:
        conditions.append(PropertyCondition(code="R000051", name="鋼板屋根", value=1.0))
    elif not tile:
        warnings.append(f"屋根材「{roof_type}」から瓦/鋼板を判定できず瓦屋根 (条件01=0) で計算しました")
    if "切妻" in roof_type:
        conditions.append(PropertyCondition(code="R000031", name="切妻屋根", value=1.0))
    return conditions


def build_cad_from_extraction(extracted: Any) -> CADData:
    """PDF抽出結果 (dict) から CADData を組み立てる。

    - rooms → RoomData (N000001=床面積 / N000100=部屋面積)
    - fittings → FittingData (T100001=数量 / T100010=幅m / T100012=高さm)
    - aggregates: H000403 延床 / B001211・B001212 階別 / H000402 建築面積 /
      B000581 外周長 / Y000181 屋根面積 / Y000540 軒樋長 / F000140 外壁面積 / H000405 施工床
    - PDFから読めなかった集計値は 延床 × RATIO_TO_TOTAL_FLOOR_AREA で推定し、
      その旨を parse_warnings に記録する

    入力が None / dict 以外 / 値の型不正 / 負数でも例外にせず、warnings に記録して進める。
    """
    cad = CADData()
    warnings = cad.parse_warnings
    if extracted is None:
        warnings.append("PDF抽出結果が空 (None) のため、数量無しの見積になります")
        extracted = {}
    elif not isinstance(extracted, dict):
        warnings.append(
            f"PDF抽出結果が辞書形式ではありません ({type(extracted).__name__}) のため、数量無しの見積になります")
        extracted = {}
    elif not extracted:
        warnings.append("PDF抽出結果が空のため、数量無しの見積になります")

    cad.property_name = _text(extracted.get("property_name")) or "PDF図面からの概算"
    cad.drawing_name = "PDF図面 (AI読み取り)"
    cad.structure = _text(extracted.get("structure"))

    cad.rooms = _build_rooms(extracted, warnings)
    cad.fittings = _build_fittings(extracted, warnings)
    cad.conditions = _build_conditions(extracted, warnings)

    # --- 階別面積 ---
    floor_areas: dict[str, float] = {}
    raw_fa = extracted.get("floor_areas")
    if isinstance(raw_fa, dict):
        for label, value in raw_fa.items():
            floor = normalize_floor_label(label)
            area = _positive(value, f"階別面積「{label}」", warnings)
            if floor and area is not None:
                floor_areas[floor] = floor_areas.get(floor, 0.0) + area
    elif raw_fa is not None:
        warnings.append("階別面積が {階: 面積} 形式でないため未取得として扱いました")

    stories = _resolve_stories(extracted, floor_areas, cad.rooms, warnings)
    if stories >= 3:
        warnings.append(f"{stories}階建は検証対象外のため、推定比率は2階建の値を流用しました")

    # --- 延床面積 ---
    room_sum = cad.room_floor_area_sum()
    tfa = _positive(extracted.get("total_floor_area_m2"), "延床面積", warnings)
    if tfa is None:
        if floor_areas:
            tfa = sum(floor_areas.values())
            warnings.append(f"延床面積はPDFから読めず階別面積の合計 {tfa:.2f}㎡ で代用しました")
        elif room_sum > 0:
            tfa = room_sum
            warnings.append(f"延床面積はPDFから読めず部屋面積の合計 {tfa:.2f}㎡ で代用しました")
        else:
            tfa = 0.0
            warnings.append("延床面積・階別面積・部屋面積のいずれもPDFから読めないため、面積に基づく項目は0で計上されます")
    else:
        if floor_areas and abs(sum(floor_areas.values()) - tfa) > tfa * AREA_MISMATCH_TOLERANCE:
            warnings.append(
                f"延床面積 {tfa:.2f}㎡ と階別面積の合計 {sum(floor_areas.values()):.2f}㎡ が"
                f"{AREA_MISMATCH_TOLERANCE:.0%}超食い違っています (延床面積を優先)")
        if room_sum > 0 and abs(room_sum - tfa) > tfa * ROOM_SUM_MISMATCH_TOLERANCE:
            warnings.append(
                f"延床面積 {tfa:.2f}㎡ と部屋面積の合計 {room_sum:.2f}㎡ が"
                f"{ROOM_SUM_MISMATCH_TOLERANCE:.0%}超食い違っています (部屋別の内訳精度に注意)")

    agg = cad.aggregates
    if tfa > 0:
        agg["H000403"] = tfa
        # 階別 (B001211=1階 / B001212=2階)
        f1, f2 = floor_areas.get("1階"), floor_areas.get("2階")
        if stories <= 1:
            if f1 is None:
                f1 = tfa
            f2 = None
        else:
            if f1 is None and f2 is None:
                f1 = round(tfa * FLOOR1_SHARE_TWO_STORY, 2)
                f2 = round(tfa - f1, 2)
                warnings.append(
                    f"階別面積はPDFから読めず延床×比率で推定しました "
                    f"(1階 {f1:.2f}㎡ / 2階 {f2:.2f}㎡、1階比率 {FLOOR1_SHARE_TWO_STORY})")
            elif f1 is None:
                f1 = max(0.0, round(tfa - f2, 2))
                warnings.append(f"1階床面積はPDFから読めず延床−2階 = {f1:.2f}㎡ で推定しました")
            elif f2 is None:
                f2 = max(0.0, round(tfa - f1, 2))
                warnings.append(f"2階床面積はPDFから読めず延床−1階 = {f2:.2f}㎡ で推定しました")
        agg["B001211"] = f1
        if f2 is not None and f2 > 0:
            agg["B001212"] = f2

        # PDFに記載があれば実測値、無ければ 延床 × 比率
        ratios = RATIO_TO_TOTAL_FLOOR_AREA[2 if stories >= 2 else 1]
        roof = extracted.get("roof") if isinstance(extracted.get("roof"), dict) else {}
        direct = {
            "H000402": _positive(extracted.get("building_area_m2"), "建築面積", warnings),
            "B000581": _positive(extracted.get("perimeter_m"), "外周長", warnings),
            "Y000181": _positive(roof.get("area_m2"), "屋根面積", warnings),
            "F000140": _positive(extracted.get("exterior_wall_area_m2"), "外壁面積", warnings),
            "Y000540": None,   # 軒樋長はPDFに書かれないため常に推定
            "H000405": None,   # 施工床面積も同上
        }
        for code, name in ESTIMATED_AGG_NAMES.items():
            value = direct.get(code)
            if value is None:
                value = round(tfa * ratios[code], 2)
                warnings.append(f"{name}はPDFから読めず延床×比率で推定 ({value:g}、比率 {ratios[code]})")
            agg[code] = value
    else:
        # 延床0: 面積依存の集計値は入れない (estimate_calculator 側の0対策に任せる)
        pass

    return cad


# --- 見積の算出・比較 ---------------------------------------------------

def estimate_from_pdf_extraction(extracted: Any, prices: dict,
                                 **calc_opts: Any) -> tuple[Estimate, CADData, list[str]]:
    """PDF抽出結果から概算見積を計算する。

    Args:
        extracted: ai_client.extract_quantities_from_pdf の戻り値 (dict)。
        prices: estimate_calculator.load_unit_prices の単価マスタ。
        **calc_opts: calculate_estimate にそのまま渡すオプション
            (include_adjustment / include_solar / safety_rate / include_custom_furniture 等)。

    Returns:
        (Estimate, CADData, warnings): warnings は PDF→CAD 変換時の警告
        (推定した項目・読めなかった値)。Estimate.warnings にも同じ内容が含まれる。
    """
    cad = build_cad_from_extraction(extracted)
    estimate = calculate_estimate(cad, prices, **calc_opts)
    return estimate, cad, list(cad.parse_warnings)


def compare_with_estimate(pdf_est: Estimate, base_est: Estimate) -> list[dict[str, Any]]:
    """工事種別ごとに PDF由来見積と基準見積 (TXT由来等) の見積金額を並べる。

    Returns:
        [{"name": 工事種別名, "pdf_total": int, "base_total": int, "diff": pdf−base}, ...]
        工事種別番号順。片方にしか無い種別は無い側を0として含める。
    """
    rows: dict[str, dict[str, Any]] = {}
    order: dict[str, int] = {}
    for est, key in ((pdf_est, "pdf_total"), (base_est, "base_total")):
        for cat in (getattr(est, "categories", None) or []):
            name = getattr(cat, "name", "") or ""
            row = rows.setdefault(name, {"name": name, "pdf_total": 0, "base_total": 0})
            row[key] += int(getattr(cat, "estimate_total", 0) or 0)
            order.setdefault(name, int(getattr(cat, "no", 0) or 0))
    result = []
    for name in sorted(rows, key=lambda n: (order.get(n, 0), n)):
        row = rows[name]
        row["diff"] = row["pdf_total"] - row["base_total"]
        result.append(row)
    return result
