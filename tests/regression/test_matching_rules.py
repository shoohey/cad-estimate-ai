"""実績比較の突合ルール（2026-10-26 未突合分析の反映）

R1 内訳展開: 提出見積の親小計(N-M)を外し内訳(N-M-K)を明細にする（種別合計は残差行で保つ）
R2 明細の無い種別（管理費）は種別合計を擬似明細にして突合
R3 正規化: 先頭記号・括弧付記（親グループ名/日付）・「・」「%」の除去
R4 順位付け: 名称＋摘要の完全一致を最優先、数字コード不一致は順位を下げる
"""
import os

import pytest

import excel_diff
from cad_parser import load_cad_file
from estimate_calculator import calculate_estimate, load_unit_prices

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SAMPLE_XLSX = os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx")
SAMPLE_TXT = os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")


@pytest.fixture(scope="module")
def actual():
    with open(SAMPLE_XLSX, "rb") as f:
        return excel_diff.parse_actual_estimate(f.read())


@pytest.fixture(scope="module")
def diff(actual):
    prices = load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv"))
    cad = load_cad_file(SAMPLE_TXT)
    est = calculate_estimate(cad, prices, safety_rate=0.05)
    return excel_diff.diff_estimates(est, actual)


class TestR1DetailExpansion:
    def test_children_replace_parent_and_totals_hold(self, actual):
        cats = {c["name"]: c for c in actual["categories"]}
        kasetsu = cats["仮設工事"]
        names = [i["name"] for i in kasetsu["items"]]
        # 親「安全対策費 387,300」は外れ、内訳の駐車場代などが明細になる
        assert "駐車場代" in names and "基礎廻り養生ｼｰﾄ" in names
        assert not [i for i in kasetsu["items"] if i["name"] == "安全対策費" and i["amount"] == 387300]
        assert any(i["name"] == "安全対策費" and i["amount"] == 142900 for i in kasetsu["items"])
        assert kasetsu["groups"] and kasetsu["groups"][0]["name"] == "安全対策費"
        # 種別合計＝明細合計（残差は「（内訳外）」行で保たれる）
        for c in actual["categories"]:
            if c["items"]:
                assert sum(i["amount"] for i in c["items"]) == c["total"], c["name"]
        assert any("内訳" in w for w in actual["warnings"])

    def test_children_carry_group(self, actual):
        cats = {c["name"]: c for c in actual["categories"]}
        moku = cats["木工事"]
        groups = {i.get("group") for i in moku["items"]}
        assert {"軸組・羽柄材・PC", "建材・下地材", "大工手間"} <= groups

    def test_no_double_count_against_summary(self, actual):
        assert actual["total_excl_tax"] == 32701000
        assert sum(c["total"] for c in actual["categories"]) + actual["rounding_adjustment"] == 32701000


class TestMatchRate:
    def test_sample_property_money_match_rate(self, diff):
        matched = sum(m["actual_amount"] for d in diff["item_diffs"].values() for m in d["matched"])
        total = matched + sum(x["amount"] for d in diff["item_diffs"].values() for x in d["actual_only"])
        assert matched / total >= 0.85, f"金額突合率 {matched / total:.1%}"

    def test_child_matches_child_not_parent(self, diff):
        m = [m for m in diff["item_diffs"]["仮設工事"]["matched"] if m["ai_name"] == "安全対策費"]
        assert m and m[0]["actual_amount"] == 142900 and m[0]["diff"] == 0

    def test_pseudo_item_for_management_fee(self, diff):
        kanri = [cat for cat in diff["item_diffs"] if "管理費" in cat]
        assert kanri, "管理費の種別が無い"
        m = diff["item_diffs"][kanri[0]]["matched"]
        assert m and m[0]["via_pseudo"] is True
        assert m[0]["actual_amount"] == next(cd["actual_total"] for cd in diff["category_diffs"] if cd["name"] == kanri[0])

    def test_adjustment_rows_match_within_group(self, diff):
        moku = diff["item_diffs"]["木工事"]
        adj = [m for m in moku["matched"] if "33調整分" in m["ai_name"]]
        assert len(adj) == 2, [m["ai_name"] for m in moku["matched"] if "調整" in m["ai_name"]]
        assert {m["actual_name"] for m in adj} == {"33調整分"}


class TestR3Normalize:
    @pytest.mark.parametrize("text,groups,expected", [
        ("●33調整分（大工手間）", ("大工手間",), "33調整分"),
        ("●33調整分（大工手間）", (), "33調整分大工手間"),      # 相手の group が無ければ括弧は残す
        ("上框・見切(挽板･無垢床用)", (), "上框見切挽板無垢床用"),  # 意味のある括弧は残す（「・」は除去）
        ("太陽光発電システム工事（見積　令和7年10月10日）", (), "太陽光発電システム工事"),
        ("33％調整費", (), "33調整費"),
        ("ﾍﾞﾀ基礎", (), "ベタ基礎"),
        ("床組無し(内部土間床)", (), "床組無し内部土間床"),
        ("", (), ""), (None, (), ""),
    ])
    def test_normalize(self, text, groups, expected):
        assert excel_diff._normalize_name(text, *groups) == expected


class TestR4Ranking:
    def test_exact_name_summary_beats_alias(self):
        ai = [{'name': 'ﾀﾀﾐ', 'summary': '半帖', 'quantity': 9, 'unit': '枚', 'amount': 128700}]
        ac = [{'name': 'ﾀﾀﾐ', 'summary': '薄畳変更追加分', 'quantity': 1, 'unit': '枚', 'amount': 38700},
              {'name': 'ﾀﾀﾐ', 'summary': '半帖', 'quantity': 9, 'unit': '枚', 'amount': 128700}]
        out = excel_diff._match_items(ai, ac)
        assert out['matched'][0]['actual_amount'] == 128700

    def test_code_mismatch_ranks_below_code_match(self):
        ai = [{'name': '単体引違い窓（HH）16005', 'summary': '', 'quantity': 1, 'unit': '箇所', 'amount': 50000},
              {'name': '単体引違い窓（HH）16524', 'summary': '', 'quantity': 1, 'unit': '箇所', 'amount': 60000}]
        ac = [{'name': '★単体引違い窓（HH）16524', 'summary': '', 'quantity': 1, 'unit': '箇所', 'amount': 61000}]
        out = excel_diff._match_items(ai, ac)
        assert out['matched'][0]['ai_name'].endswith('16524')

    def test_annex_rows_lose_ties(self):
        ai = [{'name': 'ｸﾛｽ', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 100000}]
        ac = [{'name': 'ｸﾛｽ', 'summary': '追加分', 'quantity': 1, 'unit': '式', 'amount': 5000},
              {'name': 'ｸﾛｽ', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 98000}]
        out = excel_diff._match_items(ai, ac)
        assert out['matched'][0]['actual_amount'] == 98000
