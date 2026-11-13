"""2026-11-13 打合せ: 「まとめて計上」の登録（複数行の合計が実見積では1行）

中村様の指摘:
- 実見積では 駐車場代・整地費用・基礎廻り養生シート をお客様に見せず、安全対策費にまとめて計上している
- 画面では AI の 安全対策費＋この3つ の合計が、実見積の 安全対策費 と一致する
- チェックした後に「この案件のみ／全案件共通」の二択しかなく、細かい内容（どう処理しているか）を登録できなかった

→ AI の複数行（または実績の複数行）を、相手側の1行に「まとめて計上」として登録できるようにし、
  以後の実績比較ではその組を1行にまとめて突き合わせる。
"""
import copy
import os
from types import SimpleNamespace

import pytest

import excel_diff
from cad_parser import load_cad_file
from estimate_calculator import calculate_estimate, load_unit_prices
from ui_helpers import (
    DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY, DETAIL_KIND_DIFF, DETAIL_KIND_SAME,
    MERGE_LABEL, ALIAS_SCOPE_PROJECT, ALIAS_SCOPE_GLOBAL,
    category_detail_rows, detail_kind_counts, detail_header,
    validate_merge_selection, make_merge_entry, merge_rules_from_library,
    merge_library_add, merge_library_remove, merge_library_set_scope,
    suggest_merge_candidates,
)

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SAMPLE_XLSX = os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx")
SAMPLE_TXT = os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")

SAFETY_AI = [("安全対策費", "120㎡以下"), ("駐車場代", "※駐車スペースある場合もマイナスしない"),
             ("基礎廻り養生ｼｰﾄ", ""), ("整地費用", "※整地不要の場合もマイナスしない")]


def _old_style_actual(actual: dict) -> dict:
    """本番の既存案件と同じ形: 安全対策費の内訳（隠し行）が無く、合計 387,300 の1行だけ"""
    old = copy.deepcopy(actual)
    k = next(c for c in old["categories"] if c["name"] == "仮設工事")
    rest = [i for i in k["items"] if i.get("group") != "安全対策費"]
    total = sum(int(i["amount"]) for i in k["items"] if i.get("group") == "安全対策費")
    k["items"] = [{"name": "安全対策費", "summary": "", "quantity": 1.0, "unit": "式",
                   "price": total, "amount": total, "no": "1-1"}] + rest
    k["groups"] = []
    return old


@pytest.fixture(scope="module")
def actual_new():
    with open(SAMPLE_XLSX, "rb") as f:
        return excel_diff.parse_actual_estimate(f.read())


@pytest.fixture(scope="module")
def actual_old(actual_new):
    return _old_style_actual(actual_new)


@pytest.fixture(scope="module")
def estimate():
    prices = load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv"))
    return calculate_estimate(load_cad_file(SAMPLE_TXT), prices, safety_rate=0.05)


def _safety_entry(scope=ALIAS_SCOPE_GLOBAL, project_id="", note="お客様に見せないため安全対策費に含める"):
    return make_merge_entry([{"name": n, "summary": s} for n, s in SAFETY_AI],
                            [{"name": "安全対策費", "summary": ""}],
                            "仮設工事", "古居邸", "2026-11-13 10:00", scope=scope, project_id=project_id, note=note)


# ---------------------------------------------------------------- 再現（登録が無いと分かれて見える）
def test_without_rule_safety_cost_is_split(estimate, actual_old):
    d = excel_diff.diff_estimates(estimate, actual_old)["item_diffs"]["仮設工事"]
    safety = [m for m in d["matched"] if m["actual_name"] == "安全対策費"]
    assert safety and safety[0]["diff"] == -244400
    assert {x["name"] for x in d["ai_only"]} >= {"駐車場代", "基礎廻り養生ｼｰﾄ", "整地費用"}


# ---------------------------------------------------------------- 登録するとまとめて突き合う
def test_merge_rule_combines_rows(estimate, actual_old):
    rules = merge_rules_from_library([_safety_entry()], "sample_furui")
    diff = excel_diff.diff_estimates(estimate, actual_old, merges=rules)
    d = diff["item_diffs"]["仮設工事"]
    merged = [m for m in d["matched"] if m.get("via_merge")]
    assert len(merged) == 1
    m = merged[0]
    assert m["ai_amount"] == 387300 and m["actual_amount"] == 387300 and m["diff"] == 0
    assert {x["name"] for x in m["merge"]["ai_members"]} == {n for n, _ in SAFETY_AI}
    assert [x["name"] for x in m["merge"]["actual_members"]] == ["安全対策費"]
    assert m["merge"]["note"] == "お客様に見せないため安全対策費に含める"
    assert not {x["name"] for x in d["ai_only"]} & {"駐車場代", "基礎廻り養生ｼｰﾄ", "整地費用", "安全対策費"}
    # 金額の合計は変わらない（種別合計＝明細の積み上げ）
    cat = next(c for c in diff["category_diffs"] if c["name"] == "仮設工事")
    assert sum(x["ai_amount"] for x in d["matched"]) + sum(x["amount"] for x in d["ai_only"]) == cat["ai_total"]
    assert diff["total_ai"] == excel_diff.diff_estimates(estimate, actual_old)["total_ai"]


def test_rule_not_applied_when_actual_itemizes_rows(estimate, actual_new):
    """内訳（駐車場代など）を別の行で持っている実績には、まとめを当てない（全案件共通の登録で正しい1対1を崩さない）"""
    base = excel_diff.diff_estimates(estimate, actual_new)["item_diffs"]["仮設工事"]
    rules = merge_rules_from_library([_safety_entry()], "sample_furui")
    d = excel_diff.diff_estimates(estimate, actual_new, merges=rules)["item_diffs"]["仮設工事"]
    assert not [m for m in d["matched"] if m.get("via_merge")]
    assert sorted((m["ai_name"], m["actual_name"], m["diff"]) for m in d["matched"]) == \
        sorted((m["ai_name"], m["actual_name"], m["diff"]) for m in base["matched"])


def test_area_band_summary_difference_still_merges(estimate, actual_old):
    """安全対策費の摘要（面積帯）は物件ごとに変わる。登録時と違っても、同じ名称が1行だけなら当てる"""
    e = make_merge_entry([{"name": "安全対策費", "summary": "180㎡以下"}, {"name": "駐車場代", "summary": ""},
                          {"name": "基礎廻り養生ｼｰﾄ", "summary": ""}, {"name": "整地費用", "summary": ""}],
                         [{"name": "安全対策費", "summary": ""}], "仮設工事", "別邸", "2026-11-13 10:00")
    d = excel_diff.diff_estimates(estimate, actual_old, merges=merge_rules_from_library([e], "x"))
    merged = [m for m in d["item_diffs"]["仮設工事"]["matched"] if m.get("via_merge")]
    assert merged and merged[0]["ai_amount"] == 387300


def test_other_project_rule_not_used(estimate, actual_old):
    lib = [_safety_entry(scope=ALIAS_SCOPE_PROJECT, project_id="other_pid")]
    assert merge_rules_from_library(lib, "sample_furui") == []
    assert len(merge_rules_from_library(lib, "other_pid")) == 1


def test_rule_for_other_category_not_used(estimate, actual_old):
    e = make_merge_entry([{"name": n, "summary": s} for n, s in SAFETY_AI], [{"name": "安全対策費", "summary": ""}],
                         "基礎工事", "古居邸", "2026-11-13 10:00")
    d = excel_diff.diff_estimates(estimate, actual_old, merges=merge_rules_from_library([e], "p"))
    assert not [m for m in d["item_diffs"]["仮設工事"]["matched"] if m.get("via_merge")]


def _fake_estimate(items, cat="屋外付帯工事"):
    its = [SimpleNamespace(name=n, summary=s, quantity=1, unit="式", estimate_amount=a, category="")
           for n, s, a in items]
    c = SimpleNamespace(no=21, name=cat, items=its, estimate_total=sum(a for _, _, a in items))
    return SimpleNamespace(categories=[c], total_estimate_excl_tax=c.estimate_total)


def _fake_actual(items, cat="屋外付帯工事"):
    its = [{"name": n, "summary": s, "quantity": 1.0, "unit": "式", "price": a, "amount": a} for n, s, a in items]
    return {"total_excl_tax": sum(a for _, _, a in items),
            "categories": [{"no": 21, "name": cat, "total": sum(a for _, _, a in items), "items": its, "groups": []}]}


def test_reverse_direction_one_ai_row_to_two_actual_rows():
    est = _fake_estimate([("給排水引込", "", 500000), ("雨水排水", "", 100000)])
    act = _fake_actual([("給水引込工事", "", 300000), ("排水引込工事", "", 200000), ("雨水排水", "", 100000)])
    e = make_merge_entry([{"name": "給排水引込", "summary": ""}],
                         [{"name": "給水引込工事", "summary": ""}, {"name": "排水引込工事", "summary": ""}],
                         "屋外付帯工事", "邸", "2026-11-13 10:00")
    d = excel_diff.diff_estimates(est, act, merges=merge_rules_from_library([e], "p"))["item_diffs"]["屋外付帯工事"]
    merged = [m for m in d["matched"] if m.get("via_merge")]
    assert len(merged) == 1 and merged[0]["ai_amount"] == merged[0]["actual_amount"] == 500000
    assert not d["ai_only"] and not d["actual_only"]


def test_partial_members_still_merge_and_report_missing():
    """登録した行の一部がこの物件のAI見積に無いときは、ある行だけでまとめ、無い行を知らせる"""
    est = _fake_estimate([("安全対策費", "", 142900), ("駐車場代", "", 171500)], cat="仮設工事")
    act = _fake_actual([("安全対策費", "", 314400)], cat="仮設工事")
    rules = merge_rules_from_library([_safety_entry()], "p")
    d = excel_diff.diff_estimates(est, act, merges=rules)["item_diffs"]["仮設工事"]
    m = [x for x in d["matched"] if x.get("via_merge")][0]
    assert m["ai_amount"] == 314400 and m["diff"] == 0
    assert set(m["merge"]["missing"]) == {"基礎廻り養生ｼｰﾄ", "整地費用"}


def test_merge_with_mismatched_total_shows_difference():
    est = _fake_estimate([("安全対策費", "", 142900), ("駐車場代", "", 171500)], cat="仮設工事")
    act = _fake_actual([("安全対策費", "", 400000)], cat="仮設工事")
    d = excel_diff.diff_estimates(est, act, merges=merge_rules_from_library([_safety_entry()], "p"))
    m = [x for x in d["item_diffs"]["仮設工事"]["matched"] if x.get("via_merge")][0]
    assert m["diff"] == 314400 - 400000
    rows = category_detail_rows(d["item_diffs"]["仮設工事"])
    assert rows[0]["区分"] == DETAIL_KIND_DIFF and rows[0]["扱い"].startswith(MERGE_LABEL)


# ---------------------------------------------------------------- 画面の行（扱い・件数・見出し）
def test_detail_rows_show_merge_label(estimate, actual_old):
    diff = excel_diff.diff_estimates(estimate, actual_old, merges=merge_rules_from_library([_safety_entry()], "p"))
    rows = category_detail_rows(diff["item_diffs"]["仮設工事"])
    merged = [r for r in rows if str(r.get("扱い", "")).startswith(MERGE_LABEL)]
    assert len(merged) == 1
    r = merged[0]
    assert r["区分"] == DETAIL_KIND_SAME and bool(r["登録済"]) is True
    assert "駐車場代" in r["AI品名"] and "安全対策費" in r["AI品名"] and r["実績品名"] == "安全対策費"
    counts = detail_kind_counts(rows)
    assert counts[MERGE_LABEL] == 1
    cd = next(c for c in diff["category_diffs"] if c["name"] == "仮設工事")
    assert f"{MERGE_LABEL} 1" in detail_header(cd, counts)


# ---------------------------------------------------------------- チェックした行から登録する
def _row(kind, ai=("", ""), ac=("", ""), ai_amt=0, ac_amt=0, label=""):
    return {"区分": kind, "扱い": label, "AI品名": ai[0], "AI摘要": ai[1], "AI金額": ai_amt,
            "実績品名": ac[0], "実績摘要": ac[1], "実績金額": ac_amt, "差額(AI−実績)": ai_amt - ac_amt}


def test_validate_merge_selection_from_diff_and_ai_only_rows():
    sel = [_row(DETAIL_KIND_DIFF, ai=("安全対策費", "120㎡以下"), ac=("安全対策費", ""), ai_amt=142900, ac_amt=387300),
           _row(DETAIL_KIND_AI_ONLY, ai=("駐車場代", "※駐車スペース"), ai_amt=171500),
           _row(DETAIL_KIND_AI_ONLY, ai=("基礎廻り養生ｼｰﾄ", ""), ai_amt=22900),
           _row(DETAIL_KIND_AI_ONLY, ai=("整地費用", ""), ai_amt=50000)]
    err, ai, ac = validate_merge_selection(sel)
    assert err is None
    assert [m["name"] for m in ai] == ["安全対策費", "駐車場代", "基礎廻り養生ｼｰﾄ", "整地費用"]
    assert ac == [{"name": "安全対策費", "summary": ""}]


@pytest.mark.parametrize("sel, needle", [
    ([], "チェック"),
    ([_row(DETAIL_KIND_AI_ONLY, ai=("A", "")), _row(DETAIL_KIND_AI_ONLY, ai=("B", ""))], "実績側"),
    ([_row(DETAIL_KIND_ACTUAL_ONLY, ac=("A", "")), _row(DETAIL_KIND_ACTUAL_ONLY, ac=("B", ""))], "AI側"),
    ([_row(DETAIL_KIND_AI_ONLY, ai=("A", "")), _row(DETAIL_KIND_ACTUAL_ONLY, ac=("B", ""))], "同じ商品"),
    ([_row(DETAIL_KIND_DIFF, ai=("A", ""), ac=("A", ""))], "3行以上"),
    ([_row(DETAIL_KIND_SAME, ai=("A", ""), ac=("A", "")), _row(DETAIL_KIND_AI_ONLY, ai=("B", "")),
      _row(DETAIL_KIND_ACTUAL_ONLY, ac=("C", ""))], "一致"),
    ([_row(DETAIL_KIND_DIFF, ai=("A", ""), ac=("A", ""), label="まとめ（AI 2行・実績 1行）"),
      _row(DETAIL_KIND_AI_ONLY, ai=("B", "")), _row(DETAIL_KIND_ACTUAL_ONLY, ac=("C", ""))], "登録済み"),
])
def test_validate_merge_selection_errors(sel, needle):
    err, _, _ = validate_merge_selection(sel)
    assert err and needle in err, err


def test_entry_identity_and_library_ops():
    e1 = _safety_entry()
    lib, added = merge_library_add([], e1)
    assert added and len(lib) == 1
    # 並べ方が違っても同じ組は二重登録しない
    e2 = make_merge_entry([{"name": n, "summary": s} for n, s in reversed(SAFETY_AI)],
                          [{"name": "安全対策費", "summary": ""}], "仮設工事", "別邸", "2026-11-13 11:00")
    lib2, added2 = merge_library_add(lib, e2)
    assert not added2 and len(lib2) == 1
    lib3 = merge_library_set_scope(lib2, e1, ALIAS_SCOPE_PROJECT, "sample_furui")
    assert lib3[0]["scope"] == ALIAS_SCOPE_PROJECT and lib3[0]["project_id"] == "sample_furui"
    assert merge_library_remove(lib3, e2) == []
    # 他の担当者が先に足した登録は消さない（最新の一覧に足す）
    other = make_merge_entry([{"name": "X", "summary": ""}, {"name": "Y", "summary": ""}],
                             [{"name": "Z", "summary": ""}], "木工事", "邸", "2026-11-13 12:00")
    lib4, _ = merge_library_add([other], e1)
    assert len(lib4) == 2


def test_entry_keeps_note_and_rejects_bad_scope():
    e = make_merge_entry([{"name": "A", "summary": ""}, {"name": "B", "summary": ""}], [{"name": "C", "summary": ""}],
                         "仮設工事", "邸", "t", scope="weird", note="  お客様に見せない  ")
    assert e["scope"] == ALIAS_SCOPE_GLOBAL and e["note"] == "お客様に見せない"


# ---------------------------------------------------------------- 合計が一致する組み合わせの候補
def test_suggest_finds_safety_combination(estimate, actual_old):
    diff = excel_diff.diff_estimates(estimate, actual_old)
    sugg = suggest_merge_candidates(diff["item_diffs"]["仮設工事"])
    assert sugg, "候補が出ない"
    s = sugg[0]
    assert {m["name"] for m in s["ai"]} == {n for n, _ in SAFETY_AI}
    assert [m["name"] for m in s["actual"]] == ["安全対策費"]
    assert s["ai_total"] == s["actual_total"] == 387300


def test_suggest_skips_excluded_rows(estimate, actual_old):
    """案件特有・除外として登録済みの行は候補に入れない"""
    diff = excel_diff.diff_estimates(estimate, actual_old)
    sugg = suggest_merge_candidates(diff["item_diffs"]["仮設工事"],
                                   exclude={("駐車場代", "※駐車スペースある場合もマイナスしない")})
    assert not [s for s in sugg if "駐車場代" in {m["name"] for m in s["ai"]}]


def test_suggest_not_shown_after_merge_registered(estimate, actual_old):
    diff = excel_diff.diff_estimates(estimate, actual_old,
                                     merges=merge_rules_from_library([_safety_entry()], "p"))
    assert suggest_merge_candidates(diff["item_diffs"]["仮設工事"]) == []


def _diff_of(matched=(), ai_only=(), actual_only=()):
    return {"matched": [{"ai_name": a, "actual_name": b, "ai_summary": "", "actual_summary": "",
                         "ai_amount": x, "actual_amount": y, "diff": x - y, "ai_order": o1, "actual_order": o2}
                        for a, b, x, y, o1, o2 in matched],
            "ai_only": [{"name": n, "summary": "", "amount": a, "order": o} for n, a, o in ai_only],
            "actual_only": [{"name": n, "summary": "", "amount": a, "order": o} for n, a, o in actual_only]}


def test_suggest_requires_named_pair_and_adjacent_rows():
    # 突き合っている行が無い（AIにしかない2行の合計＝実績にしかない1行）だけでは候補にしない
    d = _diff_of(ai_only=[("A", 100, 0), ("B", 200, 1)], actual_only=[("C", 300, 0)])
    assert suggest_merge_candidates(d) == []
    # 突き合っている行＋隣の行なら候補
    d = _diff_of(matched=[("安全対策費", "安全対策費", 100, 300, 5, 0)], ai_only=[("駐車場代", 200, 4)])
    assert len(suggest_merge_candidates(d)) == 1
    # 合計は合っても、見積の中で離れた行どうしは偶然の一致とみなして出さない
    d = _diff_of(matched=[("安全対策費", "安全対策費", 100, 300, 5, 0)], ai_only=[("笠木", 200, 40)])
    assert suggest_merge_candidates(d) == []
    # 0円の行・1行どうし（差額0）は対象外
    d = _diff_of(matched=[("X", "X", 100, 100, 1, 1)], ai_only=[("Z", 0, 2)])
    assert suggest_merge_candidates(d) == []


def test_suggest_reverse_direction():
    d = _diff_of(matched=[("給排水引込", "給水引込工事", 500000, 300000, 3, 7)],
                 actual_only=[("排水引込工事", 200000, 8)])
    s = suggest_merge_candidates(d)
    assert len(s) == 1 and [m["name"] for m in s[0]["actual"]] == ["給水引込工事", "排水引込工事"]


def test_suggest_no_coincidences_on_calibration_properties():
    """校正8物件（内訳まで書いた実績）では候補を出さない。本番と同じ形（安全対策費1行）にすると、
    出る候補は安全対策費のまとめだけ（2026-11-13: 条件を付ける前は偶然の組が20件以上出た）"""
    from calibration_check import iter_pairs, DEFAULT_SAFETY_RATE
    prices = load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv"))
    safety_hits = 0
    for pair in iter_pairs():
        cad_file, xlsx = pair[2], pair[3]
        with open(xlsx, "rb") as f:
            act = excel_diff.parse_actual_estimate(f.read())
        est = calculate_estimate(load_cad_file(cad_file), prices,
                                 include_solar=any("太陽光" in c["name"] for c in act["categories"]),
                                 safety_rate=DEFAULT_SAFETY_RATE)
        for label, a in (("new", act), ("old", _old_style_actual(act) if any(
                c["name"] == "仮設工事" and any(i.get("group") == "安全対策費" for i in c["items"])
                for c in act["categories"]) else None)):
            if a is None:
                continue
            for cat, idf in excel_diff.diff_estimates(est, a)["item_diffs"].items():
                for s in suggest_merge_candidates(idf):
                    assert label == "old" and cat == "仮設工事", (xlsx, label, cat, s)
                    assert [m["name"] for m in s["actual"]] == ["安全対策費"]
                    safety_hits += 1
    assert safety_hits >= 5


def test_suggest_is_bounded_for_large_categories():
    """明細の多い工事種別でも重くならない（候補の数を絞る）"""
    import time
    d = _diff_of(matched=[(f"M{i}", f"M{i}", 1000 + i, 9000 + i * 13, i, i) for i in range(40)],
                 ai_only=[(f"A{i}", 1000 + i * 37, 100 + i) for i in range(60)],
                 actual_only=[(f"B{i}", 5000 + i * 41, 100 + i) for i in range(60)])
    t = time.time()
    suggest_merge_candidates(d)
    assert time.time() - t < 1.5


# ---------------------------------------------------------------- 反証（11/13 夜・独立検証）で見つかった4件
def test_extra_same_name_row_does_not_drop_rule():
    """AI側に登録と同名の行（安全対策費の追加行など）がもう1つあっても、登録どおりの行でまとめ、余りは AIにしかない に残す"""
    est = _fake_estimate([("安全対策費", "120㎡以下", 142900), ("駐車場代", "", 171500), ("基礎廻り養生ｼｰﾄ", "", 22900),
                          ("整地費用", "", 50000), ("安全対策費", "追加", 10000)], cat="仮設工事")
    act = _fake_actual([("安全対策費", "", 387300)], cat="仮設工事")
    d = excel_diff.diff_estimates(est, act, merges=merge_rules_from_library([_safety_entry()], "p"))["item_diffs"]["仮設工事"]
    merged = [m for m in d["matched"] if m.get("via_merge")]
    assert len(merged) == 1 and merged[0]["diff"] == 0
    assert [(x["name"], x["summary"], x["amount"]) for x in d["ai_only"]] == [("安全対策費", "追加", 10000)]


def test_suggest_not_limited_by_large_rows_far_away():
    """金額の大きい無関係な行が何行あっても、隣り合う正しい組は候補に出る（上位14行で切らない）"""
    big = [(f"大きい行{i}", 1_000_000 + i, 100 + i) for i in range(20)]
    d = _diff_of(matched=[("安全対策費", "安全対策費", 142900, 387300, 5, 0)],
                 ai_only=[("駐車場代", 171500, 6), ("基礎廻り養生ｼｰﾄ", 22900, 7), ("整地費用", 50000, 8)] + big)
    s = suggest_merge_candidates(d)
    assert len(s) == 1 and {m["name"] for m in s[0]["ai"]} == {"安全対策費", "駐車場代", "基礎廻り養生ｼｰﾄ", "整地費用"}


def test_merge_note_target_is_single_side_row():
    """合計が合わない「まとめ」行の差異理由は、1行側（実績の安全対策費）に紐づける（解除しても理由が迷子にならない）"""
    from ui_helpers import merge_note_target
    m = {"ai_name": "安全対策費 ＋ 駐車場代", "actual_name": "安全対策費", "ai_summary": "", "actual_summary": "",
         "via_merge": True, "merge": {"ai_members": [{"name": "安全対策費", "summary": "120㎡以下"}, {"name": "駐車場代", "summary": ""}],
                                      "actual_members": [{"name": "安全対策費", "summary": ""}]}}
    assert merge_note_target(m) == ("安全対策費", "")
    m2 = {"ai_name": "給排水引込", "actual_name": "給水引込工事 ＋ 排水引込工事", "ai_summary": "一式", "actual_summary": "",
          "via_merge": True, "merge": {"ai_members": [{"name": "給排水引込", "summary": "一式"}],
                                       "actual_members": [{"name": "給水引込工事", "summary": ""}, {"name": "排水引込工事", "summary": ""}]}}
    assert merge_note_target(m2) == ("給排水引込", "一式")
    # 両側とも複数行なら、合成名のまま（今の表示と同じ）
    m3 = {"ai_name": "A ＋ B", "actual_name": "C ＋ D", "ai_summary": "", "actual_summary": "", "via_merge": True,
          "merge": {"ai_members": [{"name": "A", "summary": ""}, {"name": "B", "summary": ""}],
                    "actual_members": [{"name": "C", "summary": ""}, {"name": "D", "summary": ""}]}}
    assert merge_note_target(m3) == ("A ＋ B", "")
