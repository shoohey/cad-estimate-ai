"""PDF図面→概算見積 (検証機能) の回帰テスト

対象: ai_client.extract_quantities_from_pdf / normalize_extraction / parse_number、
      pdf_estimate.build_cad_from_extraction / estimate_from_pdf_extraction /
      compare_with_estimate / derive_ratios_from_calibration。
実APIは呼ばない (requests.post を monkeypatch)。Streamlit secrets も読まない (_get_secret を差し替え)。
"""
import io
import json
import os

import pytest

import ai_client
import pdf_estimate
from ai_client import AIClientError, EXTRACTION_KEYS
from cad_parser import CADData
from estimate_calculator import Estimate, WorkCategory, EstimateItem, load_unit_prices
from pdf_estimate import (build_cad_from_extraction, estimate_from_pdf_extraction,
                          compare_with_estimate, derive_ratios_from_calibration,
                          classify_fitting, normalize_floor_label,
                          RATIO_TO_TOTAL_FLOOR_AREA, FLOOR1_SHARE_TWO_STORY, ESTIMATED_AGG_NAMES)

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PRICES_CSV = os.path.join(ROOT, "tankamaster_updated.csv")

# 2階建・部屋6つ・建具5件のサンプル抽出結果 (extract_quantities_from_pdf の戻り値相当)
SAMPLE = {
    "property_name": "サンプル邸新築工事",
    "structure": "在来木造",
    "floors": 2,
    "total_floor_area_m2": 105.0,
    "floor_areas": {"1階": 55.0, "2階": 50.0},
    "building_area_m2": 58.0,
    "perimeter_m": 32.0,
    "rooms": [
        {"name": "LDK", "floor": "1階", "area_m2": 30.0},
        {"name": "玄関", "floor": "1階", "area_m2": 3.0},
        {"name": "浴室", "floor": "1階", "area_m2": 3.3},
        {"name": "洗面脱衣室", "floor": "1階", "area_m2": 3.3},
        {"name": "トイレ", "floor": "1F", "area_m2": 1.6},
        {"name": "主寝室", "floor": "2階", "area_m2": 13.0},
    ],
    "fittings": [
        {"kind": "玄関ドア", "name": "片開きドア", "width_mm": 920, "height_mm": 2330, "count": 1, "location": "玄関"},
        {"kind": "窓", "name": "引違い窓", "width_mm": 1650, "height_mm": 2000, "count": 2, "location": "LDK"},
        {"kind": "窓", "name": "FIX窓", "width_mm": 600, "height_mm": 900, "count": 3, "location": None},
        {"kind": "内部建具", "name": "片引き戸", "width_mm": 1400, "height_mm": 2000, "count": 4, "location": None},
        {"kind": "勝手口", "name": "勝手口ドア", "width_mm": 800, "height_mm": 1800, "count": 1, "location": "キッチン"},
    ],
    "roof": {"type": "切妻 ガルバリウム鋼板", "pitch": "4寸", "area_m2": None},
    "exterior_wall_area_m2": None,
    "notes": ["建具表の寸法は内法"],
    "confidence": "medium",
    "source_pages": [1, 2],
}


@pytest.fixture(scope="module")
def prices():
    if not os.path.exists(PRICES_CSV):
        pytest.skip("tankamaster_updated.csv が無い")
    return load_unit_prices(PRICES_CSV)


# ======================================================================
# pdf_estimate: CADData の組み立て
# ======================================================================

class TestBuildCadFromSample:
    def test_areas_match_extraction(self):
        cad = build_cad_from_extraction(SAMPLE)
        assert isinstance(cad, CADData)
        assert cad.property_name == "サンプル邸新築工事"
        assert cad.structure == "在来木造"
        assert cad.total_floor_area() == pytest.approx(105.0)
        assert cad.floor_area("1階") == pytest.approx(55.0)
        assert cad.floor_area("2階") == pytest.approx(50.0)
        assert cad.aggregates["H000403"] == pytest.approx(105.0)
        assert cad.aggregates["B001211"] == pytest.approx(55.0)
        assert cad.aggregates["B001212"] == pytest.approx(50.0)
        # 部屋面積合計 = 抽出した6室の合計
        assert cad.room_floor_area_sum() == pytest.approx(30 + 3 + 3.3 + 3.3 + 1.6 + 13)
        assert len(cad.rooms) == 6
        assert [r.floor for r in cad.rooms] == ["1階"] * 5 + ["2階"]   # "1F" → "1階"
        for r in cad.rooms:
            assert r.quantities["N000001"] == r.quantities["N000100"]

    def test_direct_values_are_not_estimated(self):
        cad = build_cad_from_extraction(SAMPLE)
        assert cad.aggregates["H000402"] == pytest.approx(58.0)   # 建築面積はPDF値
        assert cad.aggregates["B000581"] == pytest.approx(32.0)   # 外周長はPDF値
        joined = "\n".join(cad.parse_warnings)
        assert "建築面積はPDFから読めず" not in joined
        assert "外周長はPDFから読めず" not in joined
        # 読めなかった屋根面積・外壁面積・軒樋長・施工床は推定され、警告に残る
        for name in ("屋根面積", "外壁面積", "軒樋長", "施工床面積"):
            assert f"{name}はPDFから読めず延床×比率で推定" in joined
        ratios = RATIO_TO_TOTAL_FLOOR_AREA[2]
        assert cad.aggregates["Y000181"] == pytest.approx(105.0 * ratios["Y000181"], abs=0.01)
        assert cad.aggregates["F000140"] == pytest.approx(105.0 * ratios["F000140"], abs=0.01)
        assert cad.aggregates["Y000540"] == pytest.approx(105.0 * ratios["Y000540"], abs=0.01)
        assert cad.aggregates["H000405"] == pytest.approx(105.0 * ratios["H000405"], abs=0.01)

    def test_fittings_mapped_to_cad_codes(self):
        cad = build_cad_from_extraction(SAMPLE)
        assert len(cad.fittings) == 5
        door = cad.fittings[0]
        assert door.category == "金属/戸" and door.type_name == "片開" and door.material == "アルミ製"
        # mm → m
        assert door.quantities["T100010"] == pytest.approx(0.92)
        assert door.quantities["T100012"] == pytest.approx(2.33)
        assert door.quantities["T100001"] == 1
        win = cad.fittings[1]
        assert win.category == "金属/窓" and win.type_name == "引違い" and win.quantities["T100001"] == 2
        fix = cad.fittings[2]
        assert fix.type_name == "Fix" and fix.quantities["T100001"] == 3
        inner = cad.fittings[3]
        assert inner.category == "木製/戸" and inner.type_name == "片引き" and inner.quantities["T100001"] == 4
        back = cad.fittings[4]
        assert back.category == "金属/戸" and back.type_name == "勝手口"
        assert cad.get_fitting_count("金属") == 1 + 2 + 3 + 1
        assert cad.get_fitting_count("木製") == 4

    def test_roof_conditions(self):
        cad = build_cad_from_extraction(SAMPLE)
        assert cad.get_condition("R000051") == 1.0    # ガルバ → 鋼板屋根
        assert cad.get_condition("R000031") == 1.0    # 切妻
        tile = build_cad_from_extraction({**SAMPLE, "roof": {"type": "寄棟 瓦", "pitch": None, "area_m2": 90}})
        assert tile.get_condition("R000051") == 0.0
        assert tile.get_condition("R000031") == 0.0
        assert tile.aggregates["Y000181"] == pytest.approx(90.0)
        assert "屋根面積はPDFから読めず" not in "\n".join(tile.parse_warnings)

    def test_estimate_total_positive(self, prices):
        est, cad, warns = estimate_from_pdf_extraction(
            SAMPLE, prices, include_adjustment=True, include_solar=False,
            safety_rate=0.05, include_custom_furniture=False)
        assert isinstance(est, Estimate)
        assert est.total_estimate_excl_tax > 0
        assert est.total_floor_area_m2 == pytest.approx(105.0)
        assert est.property_name == "サンプル邸新築工事"
        # 外部建具は PDF の建具から個別計上される (フォールバックではない)
        cat07 = next(c for c in est.categories if c.no == 7)
        names = [i.name for i in cat07.items]
        assert any("玄関ドア" in n for n in names)
        assert any("電気錠" in n for n in names)
        assert sum(i.quantity for i in cat07.items if "FIX" in i.name) == 3
        # 概算安全率 (25 概算調整) がオプションで渡っている
        assert any(c.no == 25 for c in est.categories)
        # 変換時の警告は見積の warnings にも引き継がれる
        assert warns == list(cad.parse_warnings)
        assert all(w in est.warnings for w in warns)


class TestBuildCadTotalOnly:
    def test_all_missing_aggregates_are_estimated_and_warned(self, prices):
        est, cad, warns = estimate_from_pdf_extraction({"total_floor_area_m2": 100.0}, prices)
        assert est.total_estimate_excl_tax > 0
        assert cad.total_floor_area() == pytest.approx(100.0)
        joined = "\n".join(warns)
        # 階数不明 → 2階建仮定、階別面積は比率で分割
        assert "階数はPDFから読めず2階建と仮定" in joined
        assert "階別面積はPDFから読めず延床×比率で推定" in joined
        assert cad.aggregates["B001211"] == pytest.approx(100.0 * FLOOR1_SHARE_TWO_STORY, abs=0.01)
        assert cad.aggregates["B001211"] + cad.aggregates["B001212"] == pytest.approx(100.0, abs=0.02)
        # 推定した集計値が全て列挙されている
        for code, name in ESTIMATED_AGG_NAMES.items():
            assert f"{name}はPDFから読めず延床×比率で推定" in joined, name
            assert cad.aggregates[code] == pytest.approx(100.0 * RATIO_TO_TOTAL_FLOOR_AREA[2][code], abs=0.01)
        assert "屋根の種類がPDFから読めず" in joined

    def test_single_story_uses_single_story_ratios(self):
        cad = build_cad_from_extraction({"total_floor_area_m2": 90.0, "floors": 1})
        assert cad.aggregates["B001211"] == pytest.approx(90.0)
        assert "B001212" not in cad.aggregates
        assert cad.aggregates["H000402"] == pytest.approx(90.0 * RATIO_TO_TOTAL_FLOOR_AREA[1]["H000402"], abs=0.01)
        assert "2階建と仮定" not in "\n".join(cad.parse_warnings)

    def test_partial_floor_areas_are_completed(self):
        cad = build_cad_from_extraction({"total_floor_area_m2": 100.0, "floors": 2, "floor_areas": {"1階": 60.0}})
        assert cad.aggregates["B001211"] == pytest.approx(60.0)
        assert cad.aggregates["B001212"] == pytest.approx(40.0)
        assert "2階床面積はPDFから読めず延床−1階" in "\n".join(cad.parse_warnings)

    def test_total_derived_from_floor_areas_or_rooms(self):
        cad = build_cad_from_extraction({"floor_areas": {"1F": 50.0, "2F": 45.0}})
        assert cad.total_floor_area() == pytest.approx(95.0)
        assert "階別面積の合計" in "\n".join(cad.parse_warnings)
        cad2 = build_cad_from_extraction({"rooms": [{"name": "LDK", "floor": "1階", "area_m2": 20.0},
                                                     {"name": "洋室", "floor": "2階", "area_m2": 10.0}]})
        assert cad2.total_floor_area() == pytest.approx(30.0)
        assert "部屋面積の合計" in "\n".join(cad2.parse_warnings)
        assert cad2.aggregates["B001211"] + cad2.aggregates["B001212"] == pytest.approx(30.0, abs=0.02)

    def test_mismatch_between_total_and_floor_areas_is_warned(self):
        cad = build_cad_from_extraction({"total_floor_area_m2": 100.0, "floor_areas": {"1階": 70.0, "2階": 60.0}})
        assert cad.total_floor_area() == pytest.approx(100.0)   # 延床を優先
        assert "食い違っています" in "\n".join(cad.parse_warnings)


class TestBuildCadBadInput:
    @pytest.mark.parametrize("bad", [{}, None, "junk", 123, []])
    def test_no_exception_and_warned(self, bad, prices):
        cad = build_cad_from_extraction(bad)
        assert isinstance(cad, CADData)
        assert cad.parse_warnings, "警告が記録されること"
        assert cad.rooms == [] and cad.fittings == []
        assert "H000403" not in cad.aggregates
        est, _, warns = estimate_from_pdf_extraction(bad, prices)
        assert isinstance(est, Estimate)
        assert warns == cad.parse_warnings

    def test_negative_and_string_values(self):
        cad = build_cad_from_extraction({
            "total_floor_area_m2": -5,
            "floors": "two",
            "floor_areas": {"1階": "abc", "2階": -1},
            "building_area_m2": "不明",
            "rooms": "x",
            "fittings": [1, None, {"kind": "窓", "count": "abc", "width_mm": -100, "height_mm": "高さ不明"}],
            "roof": "ガルバ",
        })
        joined = "\n".join(cad.parse_warnings)
        assert "延床面積の値 -5 が0以下" in joined
        assert "階別面積「1階」「abc」を数値として読めない" in joined
        assert "階別面積「2階」の値 -1 が0以下" in joined
        assert "部屋一覧がリスト形式でない" in joined
        assert "建具一覧のうち2件" in joined
        assert "数量「abc」が不正のため1" in joined
        assert "幅の値 -100 が0以下" in joined
        assert "高さ「高さ不明」を数値として読めない" in joined
        assert len(cad.fittings) == 1 and cad.fittings[0].quantities["T100001"] == 1
        assert "T100010" not in cad.fittings[0].quantities
        # 延床が無いので面積依存の集計値は入れない
        assert "H000403" not in cad.aggregates
        assert "いずれもPDFから読めない" in joined

    def test_rooms_with_missing_floor_or_area(self):
        cad = build_cad_from_extraction({
            "total_floor_area_m2": 80, "floors": 1,
            "rooms": [{"name": "LDK"}, {"name": "", "area_m2": 5}, "x", {"name": "玄関", "floor": None, "area_m2": "3.0㎡"}],
        })
        assert [r.name for r in cad.rooms] == ["LDK", "玄関"]
        assert all(r.floor == "1階" for r in cad.rooms)
        assert cad.rooms[1].quantities["N000001"] == pytest.approx(3.0)
        joined = "\n".join(cad.parse_warnings)
        assert "部屋2室は階がPDFから読めず1階" in joined
        assert "部屋1室は面積がPDFから読めず0㎡" in joined
        assert "部屋一覧のうち2件" in joined

    def test_entrance_door_without_dimensions_still_matches_entrance(self, prices):
        est, cad, warns = estimate_from_pdf_extraction(
            {"total_floor_area_m2": 100, "fittings": [{"kind": "玄関ドア", "name": None, "count": 1}]}, prices)
        door = cad.fittings[0]
        assert door.quantities["T100012"] == pytest.approx(pdf_estimate.DEFAULT_ENTRANCE_DOOR_H_M)
        assert any("標準2.3mとして玄関ドアで計上" in w for w in warns)
        cat07 = next(c for c in est.categories if c.no == 7)
        assert any("玄関ドア" in i.name for i in cat07.items)


class TestHelpers:
    @pytest.mark.parametrize("value,expected", [
        ("1階", "1階"), ("1F", "1階"), ("１階", "1階"), ("一階", "1階"), (2, "2階"), ("2.0", "2階"),
        ("2階建", "2階"), ("RF", "R階"), ("屋上", "R階"), ("PH", "PH階"), ("B1", "B1階"), ("地下1階", "B1階"),
        ("", None), (None, None), ("小屋裏", None), (True, None),
    ])
    def test_normalize_floor_label(self, value, expected):
        assert normalize_floor_label(value) == expected

    @pytest.mark.parametrize("kind,name,expected", [
        ("玄関ドア", "親子ドア", ("金属/戸", "親子", "アルミ製")),
        ("玄関ドア", None, ("金属/戸", "片開", "アルミ製")),
        ("勝手口", "勝手口ドア", ("金属/戸", "勝手口", "アルミ製")),
        ("窓", "掃き出し窓", ("金属/窓", "引違い", "樹脂製")),
        ("窓", "縦すべり出し窓", ("金属/窓", "縦すべり", "樹脂製")),
        ("窓", "はめ殺し窓", ("金属/窓", "Fix", "樹脂製")),
        ("サッシ", "上げ下げ窓", ("金属/窓", "上げ下げ", "樹脂製")),
        ("内部建具", "折れ戸", ("木製/戸", "折戸", "木製")),
        ("内部建具", "引込み戸", ("木製/戸", "引込み", "木製")),
        ("内部建具", "片開きドア", ("木製/戸", "片開", "木製")),
        ("内部建具", "室内窓 FIX", ("木製/窓", "Fix", "木製")),
        (None, "吊戸", ("木製/戸", "片引き", "木製")),
        (None, None, ("金属/窓", "窓", "樹脂製")),
    ])
    def test_classify_fitting(self, kind, name, expected):
        assert classify_fitting(kind, name) == expected


class TestCompareWithEstimate:
    @staticmethod
    def _est(cats):
        est = Estimate()
        for no, name, amount in cats:
            cat = WorkCategory(no=no, name=name)
            cat.items.append(EstimateItem(parent_no=no, child_no=1, name="x", quantity=1,
                                          estimate_amount=amount, order_amount=amount))
            est.categories.append(cat)
        return est

    def test_rows_and_diff(self):
        pdf = self._est([(1, "仮設工事", 100), (2, "基礎工事", 300), (25, "概算調整", 10)])
        base = self._est([(2, "基礎工事", 250), (1, "仮設工事", 120), (3, "木工事", 500)])
        rows = compare_with_estimate(pdf, base)
        assert [r["name"] for r in rows] == ["仮設工事", "基礎工事", "木工事", "概算調整"]
        assert rows[0] == {"name": "仮設工事", "pdf_total": 100, "base_total": 120, "diff": -20}
        assert rows[1]["diff"] == 50
        assert rows[2] == {"name": "木工事", "pdf_total": 0, "base_total": 500, "diff": -500}
        assert rows[3] == {"name": "概算調整", "pdf_total": 10, "base_total": 0, "diff": 10}
        assert set(rows[0]) == {"name", "pdf_total", "base_total", "diff"}

    def test_empty(self):
        assert compare_with_estimate(Estimate(), Estimate()) == []
        assert compare_with_estimate(None, Estimate()) == []

    def test_real_pipeline(self, prices):
        pdf_est, _, _ = estimate_from_pdf_extraction(SAMPLE, prices)
        base_est, _, _ = estimate_from_pdf_extraction({**SAMPLE, "total_floor_area_m2": 120.0,
                                                       "floor_areas": {"1階": 60, "2階": 60}}, prices)
        rows = compare_with_estimate(pdf_est, base_est)
        assert len(rows) == len(pdf_est.categories)
        assert sum(r["pdf_total"] for r in rows) == sum(c.estimate_total for c in pdf_est.categories)
        assert sum(r["base_total"] for r in rows) == sum(c.estimate_total for c in base_est.categories)
        assert all(r["diff"] == r["pdf_total"] - r["base_total"] for r in rows)


class TestCalibrationRatios:
    @pytest.fixture(scope="class")
    def ratios(self):
        try:
            return derive_ratios_from_calibration()
        except FileNotFoundError:
            pytest.skip("calibration ディレクトリ (顧客実データ) が無い環境")

    def test_ratios_in_plausible_range(self, ratios):
        assert ratios["n_properties"] >= 1
        bounds = {
            "H000402": (0.4, 1.2),    # 建築面積/延床
            "B000581": (0.2, 1.0),    # 外周長(m)/延床(㎡)
            "Y000181": (0.3, 1.5),    # 屋根面積/延床
            "Y000540": (0.1, 0.5),    # 軒樋長(m)/延床(㎡)
            "F000140": (1.0, 2.5),    # 外壁面積/延床
            "H000405": (1.0, 1.2),    # 施工床/延床
        }
        for stories, codes in ratios["by_stories"].items():
            for code, (lo, hi) in bounds.items():
                st = codes[code]
                if st["n"] == 0:
                    continue
                assert lo <= st["min"] <= st["mean"] <= st["max"] <= hi, (stories, code, st)
        share = ratios["floor1_share_two_story"]
        if share["n"]:
            assert 0.4 <= share["min"] <= share["mean"] <= share["max"] <= 0.9

    def test_embedded_constants_match_calibration(self, ratios):
        """埋め込み定数の更新忘れ検知: 実測平均と 0.02 以上ズレたら定数を差し替える"""
        for stories, codes in RATIO_TO_TOTAL_FLOOR_AREA.items():
            for code, embedded in codes.items():
                st = ratios["by_stories"][stories][code]
                if st["n"] == 0:
                    continue
                assert abs(st["mean"] - embedded) < 0.02, (stories, code, st, embedded)
        assert abs(ratios["floor1_share_two_story"]["mean"] - FLOOR1_SHARE_TWO_STORY) < 0.02


# ======================================================================
# ai_client: 数値正規化・スキーマ正規化
# ======================================================================

class TestParseNumber:
    @pytest.mark.parametrize("value,expected", [
        (115.3, 115.3), (105, 105.0), ("115.30", 115.3), ("115.30㎡", 115.3), ("１１５．３", 115.3),
        ("2,700", 2700.0), ("W2700", 2700.0), ("約 4.5 m", 4.5), ("-5", -5.0), ("2階建", 2.0),
        (None, None), (True, None), (False, None), ("", None), ("abc", None), ([1], None), ({"a": 1}, None),
        (float("nan"), None), (float("inf"), None),
    ])
    def test_parse_number(self, value, expected):
        assert ai_client.parse_number(value) == expected

    def test_helpers(self):
        assert ai_client._to_positive("0") is None and ai_client._to_positive(-3) is None
        assert ai_client._to_positive("3.5㎡") == 3.5
        assert ai_client._to_int("2階建") == 2 and ai_client._to_int(2.6) == 3 and ai_client._to_int("x") is None
        assert ai_client._to_text("  a ") == "a" and ai_client._to_text("") is None and ai_client._to_text(False) is None


class TestNormalizeExtraction:
    def test_all_keys_present_for_empty_input(self):
        for raw in ({}, None, "junk", []):
            out = ai_client.normalize_extraction(raw)
            assert tuple(out) == EXTRACTION_KEYS
            assert out["rooms"] == [] and out["fittings"] == [] and out["floor_areas"] == {}
            assert out["confidence"] == "low" and out["source_pages"] == []
            assert out["roof"] == {"type": None, "pitch": None, "area_m2": None}
            assert out["total_floor_area_m2"] is None and out["floors"] is None

    def test_values_are_coerced(self):
        out = ai_client.normalize_extraction({
            "property_name": " サンプル邸 ", "floors": "2階建", "total_floor_area_m2": "105.30㎡",
            "floor_areas": {"1階": "55.5", "2階": -1, "": 3},
            "building_area_m2": "不明", "perimeter_m": 0,
            "rooms": [{"name": "LDK", "floor": "1階", "area_m2": "30"}, {"name": "", "area_m2": 3}, "x",
                      {"name": "納戸", "area_m2": "abc"}],
            "fittings": [{"kind": "窓", "name": "引違い窓", "width_mm": "1,650", "height_mm": 2000, "count": "2"},
                         {"kind": None, "name": None}, {"name": "片開き", "count": -1}, 5],
            "roof": {"type": "切妻", "area_m2": "abc"},
            "notes": ["a", "", None, 3],
            "confidence": "HIGH",
            "source_pages": [2, "1", 0, -3, "x", 2],
        })
        assert out["property_name"] == "サンプル邸"
        assert out["floors"] == 2
        assert out["total_floor_area_m2"] == 105.3
        assert out["floor_areas"] == {"1階": 55.5}
        assert out["building_area_m2"] is None and out["perimeter_m"] is None
        assert out["rooms"] == [{"name": "LDK", "floor": "1階", "area_m2": 30.0},
                                {"name": "納戸", "floor": None, "area_m2": None}]
        assert out["fittings"] == [
            {"kind": "窓", "name": "引違い窓", "width_mm": 1650.0, "height_mm": 2000.0, "count": 2, "location": None},
            {"kind": None, "name": "片開き", "width_mm": None, "height_mm": None, "count": 1, "location": None},
        ]
        assert out["roof"] == {"type": "切妻", "pitch": None, "area_m2": None}
        assert out["notes"] == ["a", "3"]
        assert out["confidence"] == "high"
        assert out["source_pages"] == [1, 2]

    def test_floor_areas_list_form(self):
        out = ai_client.normalize_extraction({"floor_areas": [{"floor": "1階", "area_m2": 50}, {"floor": "2階", "area": 40}, "x"]})
        assert out["floor_areas"] == {"1階": 50.0, "2階": 40.0}


# ======================================================================
# ai_client.extract_quantities_from_pdf (requests.post を差し替え)
# ======================================================================

def _make_pdf(pages: int = 1) -> bytes:
    from pypdf import PdfWriter
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=595, height=842)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


class _FakeResponse:
    def __init__(self, text: str, status: int = 200):
        self.status_code = status
        self._text = text
        self.text = text

    def json(self):
        if self.status_code >= 300:
            return {"error": {"message": self._text}}
        return {"content": [{"type": "text", "text": self._text}], "stop_reason": "end_turn"}


class _FakePost:
    """requests.post の代替。応答テキストを順番に返し、呼び出し内容を記録する"""

    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def __call__(self, url, json=None, headers=None, timeout=None, **kw):
        self.calls.append({"url": url, "json": json, "headers": headers, "timeout": timeout})
        item = self.responses.pop(0)
        if isinstance(item, Exception):
            raise item
        if isinstance(item, tuple):
            return _FakeResponse(item[1], item[0])
        return _FakeResponse(item)


GOOD_JSON = json.dumps({
    "property_name": "サンプル邸", "structure": "在来木造", "floors": 2, "total_floor_area_m2": 105.3,
    "floor_areas": {"1階": 55.3, "2階": 50.0}, "building_area_m2": None, "perimeter_m": "32.5",
    "rooms": [{"name": "LDK", "floor": "1階", "area_m2": 30.1}],
    "fittings": [{"kind": "窓", "name": "引違い窓", "width_mm": 1650, "height_mm": 2000, "count": 2, "location": "LDK"}],
    "roof": {"type": "切妻", "pitch": "4寸", "area_m2": None}, "exterior_wall_area_m2": None,
    "notes": ["面積表より"], "confidence": "high", "source_pages": [1],
}, ensure_ascii=False)


@pytest.fixture
def fake_env(monkeypatch):
    """APIキーあり・モデル既定。Streamlit secrets は読まない"""
    monkeypatch.setattr(ai_client, "_get_secret", lambda key: {"ANTHROPIC_API_KEY": "sk-test"}.get(key, ""))


def _install(monkeypatch, responses) -> _FakePost:
    fake = _FakePost(responses)
    monkeypatch.setattr(ai_client.requests, "post", fake)
    return fake


class TestExtractQuantitiesFromPdf:
    def test_success_and_request_shape(self, monkeypatch, fake_env):
        pdf = _make_pdf(2)
        fake = _install(monkeypatch, [GOOD_JSON])
        out = ai_client.extract_quantities_from_pdf(pdf, "plan.pdf")
        assert tuple(out) == EXTRACTION_KEYS
        assert out["property_name"] == "サンプル邸" and out["floors"] == 2
        assert out["total_floor_area_m2"] == 105.3 and out["perimeter_m"] == 32.5
        assert out["floor_areas"] == {"1階": 55.3, "2階": 50.0}
        assert out["fittings"][0]["count"] == 2 and out["confidence"] == "high"
        assert len(fake.calls) == 1
        call = fake.calls[0]
        assert call["url"] == ai_client.ANTHROPIC_ENDPOINT
        assert call["headers"]["anthropic-version"] == ai_client.ANTHROPIC_VERSION
        assert call["headers"]["x-api-key"] == "sk-test"
        assert call["timeout"] == ai_client.PDF_TIMEOUT_SEC
        body = call["json"]
        assert body["model"] == ai_client.DEFAULT_MODEL
        assert body["max_tokens"] == ai_client.MAX_TOKENS
        assert "W2700" in body["system"] and "null" in body["system"]
        content = body["messages"][0]["content"]
        assert body["messages"][0]["role"] == "user"
        assert content[0]["type"] == "document"
        assert content[0]["source"] == {"type": "base64", "media_type": "application/pdf",
                                        "data": __import__("base64").standard_b64encode(pdf).decode("ascii")}
        assert "\n" not in content[0]["source"]["data"]
        assert content[1]["type"] == "text" and "plan.pdf" in content[1]["text"]

    def test_hints_are_appended_only_when_given(self, monkeypatch, fake_env):
        pdf = _make_pdf()
        fake = _install(monkeypatch, [GOOD_JSON, GOOD_JSON, GOOD_JSON])
        ai_client.extract_quantities_from_pdf(pdf, "a.pdf", hints="2階建。延床は面積表の値を採用 ")
        text = fake.calls[0]["json"]["messages"][0]["content"][1]["text"]
        assert text.endswith("## 担当者からの補足\n2階建。延床は面積表の値を採用")
        ai_client.extract_quantities_from_pdf(pdf, "a.pdf", hints="   ")
        ai_client.extract_quantities_from_pdf(pdf, "a.pdf")
        for call in fake.calls[1:]:
            assert "担当者からの補足" not in call["json"]["messages"][0]["content"][1]["text"]

    def test_code_fence_and_preamble(self, monkeypatch, fake_env):
        pdf = _make_pdf()
        _install(monkeypatch, ["```json\n" + GOOD_JSON + "\n```"])
        assert ai_client.extract_quantities_from_pdf(pdf, "a.pdf")["property_name"] == "サンプル邸"
        _install(monkeypatch, ["読み取り結果は以下です。\n" + GOOD_JSON + "\n以上"])
        assert ai_client.extract_quantities_from_pdf(pdf, "a.pdf")["floors"] == 2

    def test_retry_once_on_invalid_json(self, monkeypatch, fake_env):
        pdf = _make_pdf()
        fake = _install(monkeypatch, ["これはJSONではありません", GOOD_JSON])
        out = ai_client.extract_quantities_from_pdf(pdf, "a.pdf")
        assert out["property_name"] == "サンプル邸"
        assert len(fake.calls) == 2
        first_text = fake.calls[0]["json"]["messages"][0]["content"][1]["text"]
        retry_text = fake.calls[1]["json"]["messages"][0]["content"][1]["text"]
        assert retry_text.startswith(first_text) and "JSONとして解析できませんでした" in retry_text
        # 再試行でもPDFは同じものを送る
        assert fake.calls[1]["json"]["messages"][0]["content"][0] == fake.calls[0]["json"]["messages"][0]["content"][0]

    def test_json_array_is_rejected_then_retry(self, monkeypatch, fake_env):
        fake = _install(monkeypatch, ["[1, 2, 3]", GOOD_JSON])
        assert ai_client.extract_quantities_from_pdf(_make_pdf(), "a.pdf")["floors"] == 2
        assert len(fake.calls) == 2

    def test_both_invalid_raises(self, monkeypatch, fake_env):
        fake = _install(monkeypatch, ["ダメ", "{not json"])
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(_make_pdf(), "a.pdf")
        assert "JSONとして解析できませんでした" in str(exc.value)
        assert len(fake.calls) == 2

    def test_missing_api_key(self, monkeypatch):
        monkeypatch.setattr(ai_client, "_get_secret", lambda key: "")
        fake = _install(monkeypatch, [GOOD_JSON])
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(_make_pdf(), "a.pdf")
        assert "ANTHROPIC_API_KEY" in str(exc.value)
        assert fake.calls == []

    def test_size_limit(self, monkeypatch, fake_env):
        fake = _install(monkeypatch, [GOOD_JSON])
        too_big = b"%PDF-1.4\n" + b"\0" * ai_client.PDF_MAX_BYTES
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(too_big, "big.pdf")
        assert "32MB" in str(exc.value) and "big.pdf" in str(exc.value)
        assert fake.calls == []

    def test_page_limit(self, monkeypatch, fake_env):
        fake = _install(monkeypatch, [GOOD_JSON])
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(_make_pdf(ai_client.PDF_MAX_PAGES + 1), "many.pdf")
        assert "101 ページ" in str(exc.value) and "100 ページ" in str(exc.value)
        assert fake.calls == []
        # 上限ちょうどは通る
        assert ai_client.extract_quantities_from_pdf(_make_pdf(ai_client.PDF_MAX_PAGES), "ok.pdf")["floors"] == 2
        assert len(fake.calls) == 1

    def test_empty_and_broken_pdf(self, monkeypatch, fake_env):
        fake = _install(monkeypatch, [GOOD_JSON])
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(b"", "empty.pdf")
        assert "空" in str(exc.value)
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(b"this is not a pdf at all", "broken.pdf")
        assert "PDFとして読み取れません" in str(exc.value)
        assert fake.calls == []

    def test_http_error_and_timeout(self, monkeypatch, fake_env):
        _install(monkeypatch, [(500, "server error")])
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(_make_pdf(), "a.pdf")
        assert "HTTP 500" in str(exc.value)
        _install(monkeypatch, [ai_client.requests.Timeout()])
        with pytest.raises(AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(_make_pdf(), "a.pdf")
        assert "タイムアウト" in str(exc.value) and str(ai_client.PDF_TIMEOUT_SEC) in str(exc.value)

    def test_existing_text_call_still_works(self, monkeypatch, fake_env):
        """_call_anthropic の後方互換: 文字列 content・テキスト用タイムアウト"""
        fake = _install(monkeypatch, ["[]"])
        assert ai_client.parse_voice_command("エアコンを2台に", "{}") == []
        assert fake.calls[0]["json"]["messages"][0]["content"] == (
            "## 現在の見積サマリ (JSON)\n{}\n\n## 音声コマンド (文字起こし)\nエアコンを2台に\n\n"
            "上記の音声コマンドを見積への変更指示として構造化し、JSON配列で出力してください。")
        assert fake.calls[0]["timeout"] == ai_client.TEXT_TIMEOUT_SEC
