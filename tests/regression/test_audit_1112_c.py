"""2026-11-12 監査で確認された不足の再現テスト（第3群: 設計部の単価マスタを見積に使う）

[13] 図面上の可動棚などのマスタ項目を見積に入れる方法が手入力しかない
[14] 類似候補から「同じ商品」登録しても、同じ表記どうしの無意味な登録ができてしまう
[15] 設計部マスタ保存後、「アプリ内 単価マスタ」が 0 行と表示される
[16] 取込プレビューが英語の列名で、どの列を販価・原価に採ったか分からない
"""
import json
import os

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401
from ui_helpers import design_breakdown_rows, design_item_to_new_item, design_master_preview, validate_alias_pair

DESIGN_ROWS = [
    {'code': '4-KD-001', 'category': '内部建具', 'group': '4.内部建具工事', 'name': '可動棚5枚　W4.5尺・ﾌｨｯﾄﾗｯｸ（SC419a）',
     'summary': '材工', 'unit': 'ヶ所', 'order_price': 52000, 'estimate_price': 77500, 'row_no': 10, 'cost_kind': '材工',
     'breakdown_of': ''},
    {'code': '4-KD-001-1', 'category': '内部建具', 'group': '4.内部建具工事', 'name': '内訳①：[材] 可動棚 材料',
     'summary': '', 'unit': 'ヶ所', 'order_price': 32000, 'estimate_price': 47700, 'row_no': 11, 'cost_kind': '材',
     'breakdown_of': '4-KD-001'},
    {'code': '4-KD-001-2', 'category': '内部建具', 'group': '4.内部建具工事', 'name': '内訳②：[工] 可動棚 取付',
     'summary': '', 'unit': 'ヶ所', 'order_price': 20000, 'estimate_price': 29800, 'row_no': 12, 'cost_kind': '工事費',
     'breakdown_of': '4-KD-001'},
    {'code': '2-KS-002', 'category': '基礎工事', 'group': '2.構造オプション工事', 'name': '深基礎', 'summary': '高低差450mm以下',
     'unit': 'm', 'order_price': 8000, 'estimate_price': 12000, 'row_no': 8, 'cost_kind': '材工', 'breakdown_of': ''},
]


def test_preview_uses_japanese_columns_and_states_price_columns():
    rows = design_master_preview(DESIGN_ROWS)
    assert set(rows[0].keys()) >= {'コード', '工事種別', '名称', '摘要', '単位', '見積単価（販価）', '発注単価（設計原価）', '材工区分', '内訳の親'}
    assert rows[1]['内訳の親'] == '4-KD-001'


def test_breakdown_rows_and_new_item():
    bd = design_breakdown_rows(DESIGN_ROWS, '4-KD-001')
    assert [r['name'] for r in bd] == ['内訳①：[材] 可動棚 材料', '内訳②：[工] 可動棚 取付']
    sg = {'name': DESIGN_ROWS[0]['name'], 'summary': '材工', 'unit': 'ヶ所', 'estimate_price': 77500, 'order_price': 52000,
          'code': '4-KD-001'}
    ni = design_item_to_new_item(sg, 6, 2)
    assert ni == {'cat_no': 6, 'name': sg['name'], 'summary': '材工', 'quantity': 2.0, 'unit': 'ヶ所',
                  'estimate_price': 77500, 'order_price': 52000, 'source': '設計部単価マスタ 4-KD-001'}
    # 設計原価が無い項目は粗利33%で原価を入れる
    assert design_item_to_new_item({**sg, 'order_price': 0}, 6, 1)['order_price'] == 51925
    assert design_item_to_new_item(sg, 6, 0)['quantity'] == 1.0


def test_validate_alias_pair():
    assert validate_alias_pair('深基礎', '高低差450mm以下', '深基礎', '高低差450mm以下')     # 同じ表記どうしは無意味
    assert validate_alias_pair('', '', '深基礎', '') is not None
    assert validate_alias_pair('基礎 深基礎', '', '深基礎', '高低差450mm以下') is None


def _with_design_master(app_dir):
    with open(os.path.join(app_dir, "projects", "design_price_master.json"), "w", encoding="utf-8") as f:
        json.dump({'rows': DESIGN_ROWS, 'file': 'design.xlsx', 'at': '2026-11-12'}, f, ensure_ascii=False)


def test_app_master_count_with_design_master(app_dir):
    _with_design_master(app_dir)
    at = _boot(app_dir)
    m = [x for x in at.metric if x.label == "アプリ内 単価マスタ"][0]
    assert int(str(m.value).split()[0]) >= 200, m.value
    d = [x for x in at.metric if x.label == "設計部 単価マスタ"][0]
    assert str(d.value).startswith("4")
    os.remove(os.path.join(app_dir, "projects", "design_price_master.json"))


def test_add_design_master_item_from_edit_tab(app_dir):
    _with_design_master(app_dir)
    at = _boot(app_dir)
    [t for t in at.text_input if t.key == "dm_search"][0].input("可動棚").run()
    assert not at.exception, [str(e.value) for e in at.exception]
    pick = [s for s in at.selectbox if s.key == "dm_pick"][0]
    assert any("可動棚" in str(o) for o in pick.options)
    assert not any("内訳" in str(o) for o in pick.options), "内訳行が候補に出ている"
    [n for n in at.number_input if n.key == "dm_qty"][0].set_value(2.0)
    [b for b in at.button if b.key == "dm_add"][0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    items = at.session_state["proj_sample_furui_mods"]["new_items"]
    added = [i for i in items if "可動棚" in i["name"]]
    assert added and added[0]["estimate_price"] == 77500 and added[0]["order_price"] == 52000
    assert added[0]["quantity"] == 2.0
    os.remove(os.path.join(app_dir, "projects", "design_price_master.json"))


def test_add_design_candidate_from_actual_comparison(app_dir):
    with open(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx"), "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    # 「実績にしかない」行と同じ名前の項目を設計部マスタに入れておく（候補が出る前提を作る）
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    est = calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                             load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")), safety_rate=0.05)
    diff = excel_diff.diff_estimates(est, actual)
    extra = []
    for cat, d in diff["item_diffs"].items():
        for x in d["actual_only"]:
            extra.append({'code': f"T-{len(extra)}", 'category': cat, 'group': '', 'name': x['name'],
                          'summary': x.get('summary', ''), 'unit': x.get('unit') or '式', 'order_price': 10000,
                          'estimate_price': 15000, 'row_no': 100 + len(extra), 'cost_kind': '材工', 'breakdown_of': ''})
    with open(os.path.join(app_dir, "projects", "design_price_master.json"), "w", encoding="utf-8") as f:
        json.dump({'rows': DESIGN_ROWS + extra, 'file': 'design.xlsx', 'at': '2026-11-12'}, f, ensure_ascii=False)
    at = _boot(app_dir)
    import re
    headers = [e.label for e in at.expander if "AI " in e.label and "実績 " in e.label and "差額 " in e.label]
    idx = next(i for i, h in enumerate(headers) if int(re.search(r"実績のみ (\d+)", h).group(1)) > 0)
    toggles = [c for c in at.checkbox if c.key and c.key.startswith("cand_toggle_sample_furui_")]
    toggles[idx].check().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    add_btns = [b for b in at.button if b.key and b.key.startswith("cand_add_sample_furui_")]
    assert add_btns, "設計部マスタの候補を見積に追加するボタンが無い"
    before = len(at.session_state["proj_sample_furui_mods"]["new_items"])
    add_btns[0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    assert len(at.session_state["proj_sample_furui_mods"]["new_items"]) == before + 1
    os.remove(os.path.join(app_dir, "projects", "design_price_master.json"))
    os.remove(os.path.join(app_dir, "projects", "sample_furui_actual.json"))
