"""2026-11-12 修正の反証で確認された問題の再現テスト（第5群）"""
import json
import os

import pytest

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401
from ui_helpers import (DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY, alias_pairs_from_library, category_detail_rows,
                        design_category_no, make_alias_entry, parse_money_jp, proposal_number, proposal_value,
                        search_design_master, withdrawn_note_keys)


# ---------------------------------------------------------------- 数字の読み方（直しすぎの戻し）
@pytest.mark.parametrize("ctype,text,expected", [
    ("quantity", "軒樋 L=12m", 12), ("quantity", "長さ12m", 12), ("quantity", "手すり 長さ 3.6m", 3.6),
    ("price", "UB：400,000円", 400000), ("price", "IH：98,000円", 98000), ("quantity", "WC：2箇所", 2),
    ("quantity", "TKS05305JA を 2 に", 2), ("price", "XCH1601RWS 98,000", 98000), ("price", "TCF8GM 98000", 98000),
    ("add", "1,500,000", 1500000), ("spec", "食洗機 1,500,000", 1500000),
    # 直した内容は保つ
    ("spec", "TOTO サザナ 1616サイズ", 0), ("spec", "I型2550 に変更", 0), ("quantity", "カップボード W2700 1台", 1),
    ("quantity", "2700mmのカップボードを2台", 2), ("quantity", "幅2700のカップボード 1台", 1), ("price", "514,300円", 514300),
])
def test_numbers(ctype, text, expected):
    assert proposal_number(ctype, text, None) == expected


def test_zero_is_a_value_not_missing():
    assert proposal_value("quantity", "数量を0に", None) == 0.0
    assert proposal_value("quantity", "0本", None) == 0.0
    assert proposal_value("quantity", "仕様の相談", None) is None
    assert proposal_value("spec", "TOTO サザナ 1616サイズ", None) is None
    assert parse_money_jp("UB：400,000円") == 400000.0


# ---------------------------------------------------------------- 同じ商品の照合
def test_name_only_registration_hits_only_the_summaryless_row():
    entry = make_alias_entry({'AI品名': '打合せ造作', 'AI摘要': ''}, {'実績品名': '造作工事', '実績摘要': ''}, '木工事', 'A邸', 'd')
    ai = [{'name': '打合せ造作', 'summary': 'カウンター', 'quantity': 1, 'unit': '式', 'amount': 30000},
          {'name': '打合せ造作', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 50000}]
    ac = [{'name': '造作工事', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 31000}]
    out = excel_diff._match_items(ai, ac, alias_pairs_from_library([entry]))
    assert len(out['matched']) == 1 and out['matched'][0]['ai_summary'] == ''


def test_registered_pair_still_works_on_other_property_with_different_summary():
    # A邸で摘要付きで登録（全案件共通）→ B邸では摘要が違うが、同じ名称の行が1つだけなら突き合う
    entry = make_alias_entry({'AI品名': '標準ｸﾛｰｾﾞｯﾄﾄﾞｱ 折戸', 'AI摘要': 'LIXIL ﾗｼｯｻS'},
                             {'実績品名': '標準ﾄﾞｱ RAD ﾗｲﾝ枠', '実績摘要': 'LIXIL ﾗﾌｨｽ'}, '内部建具工事', 'A邸', 'd')
    ai = [{'name': '標準ｸﾛｰｾﾞｯﾄﾄﾞｱ 折戸', 'summary': '木製 4枚折戸', 'quantity': 1, 'unit': '箇所', 'amount': 40800}]
    ac = [{'name': '標準ﾄﾞｱ RAD ﾗｲﾝ枠', 'summary': '材工 W900', 'quantity': 1, 'unit': '箇所', 'amount': 41000}]
    out = excel_diff._match_items(ai, ac, alias_pairs_from_library([entry]))
    assert len(out['matched']) == 1 and out['matched'][0]['via_alias'] is True


def test_name_fallback_not_used_when_name_is_ambiguous():
    entry = make_alias_entry({'AI品名': '打合せ造作', 'AI摘要': 'カウンター'}, {'実績品名': '造作工事', '実績摘要': '棚'}, '木工事', 'A邸', 'd')
    ai = [{'name': '打合せ造作', 'summary': 'ニッチ', 'quantity': 1, 'unit': '式', 'amount': 30000},
          {'name': '打合せ造作', 'summary': '収納', 'quantity': 1, 'unit': '式', 'amount': 50000}]
    ac = [{'name': '造作工事', 'summary': '別', 'quantity': 1, 'unit': '式', 'amount': 31000}]
    out = excel_diff._match_items(ai, ac, alias_pairs_from_library([entry]))
    assert not any(m.get('via_alias') for m in out['matched'])


# ---------------------------------------------------------------- 差異理由・案件特有の表示
def test_withdrawn_keys_include_emptied_reason():
    old = {'a': {'reason': 'x', 'category': '数量の違い'}, 'b': {'reason': 'y'}, 'c': {'reason': 'z'}}
    new = {'a': {'reason': '', 'category': '数量の違い', 'scope': 'global'}, 'c': {'reason': 'z', 'scope': 'project'}}
    assert sorted(withdrawn_note_keys(old, new)) == ['a', 'b', 'c']


def test_special_label_respects_summary():
    item_diff = {'matched': [], 'ai_only': [{'name': '人通口補強', 'summary': '60㎡超過分', 'amount': 8188},
                                            {'name': '人通口補強', 'summary': '別仕様', 'amount': 1000}], 'actual_only': []}
    rows = category_detail_rows(item_diff, {('人通口補強', '60㎡超過分'): '除外: 地域・条例の違い'})
    labels = {r['AI摘要']: r['扱い'] for r in rows}
    assert labels['60㎡超過分'].startswith('除外') and labels['別仕様'] == ''
    # 旧形式（品名だけのキー）も読める
    rows = category_detail_rows(item_diff, {'人通口補強': '案件特有'})
    assert all(r['扱い'] == '案件特有' for r in rows)


# ---------------------------------------------------------------- 設計部マスタ
def test_design_category_no():
    pairs = [(1, '仮設工事'), (3, '木工事'), (6, '内部建具工事'), (14, '住宅設備機器工事'), (7, '外部建具工事'), (21, '雑工事')]
    assert design_category_no({'category': '可動棚', 'group': '4.内部建具工事'}, pairs) == 6
    assert design_category_no({'category': 'キッチン', 'group': '7.住宅設備工事'}, pairs) == 14
    assert design_category_no({'category': 'サッシ', 'group': '3.外部建具工事'}, pairs) == 7
    assert design_category_no({'category': '謎', 'group': ''}, pairs) == 21


def test_zero_price_headers_excluded_from_search():
    rows = [{'name': '●可動棚＝棚柱四隅タイプ', 'summary': '', 'unit': '', 'estimate_price': 0, 'order_price': 0, 'breakdown_of': ''},
            {'name': '可動棚5枚', 'summary': '材工', 'unit': 'ヶ所', 'estimate_price': 77500, 'order_price': 52000, 'breakdown_of': ''}]
    assert [r['name'] for r in search_design_master(rows, '可動棚')] == ['可動棚5枚']


def test_added_design_item_keeps_andpad_split():
    import app
    from andpad_converter import convert_to_andpad
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    from ui_helpers import design_item_to_new_item
    import master_lookup
    with open(os.path.join(ROOT, "tankamaster_updated.csv"), "rb") as f:
        tanka, _ = master_lookup.parse_price_master(f.read(), "tankamaster_updated.csv")
    est = calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                             load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")))
    cat = next(c for c in est.categories if '内部建具' in c.name)
    sg = {'name': '可動棚5枚　W4.5尺・ﾌｨｯﾄﾗｯｸ（SC419a）', 'summary': '材工', 'unit': 'ヶ所', 'estimate_price': 77500,
          'order_price': 52000, 'code': '4-KD-001', 'category': '可動棚', 'cost_kind': '材工'}
    ni = design_item_to_new_item(sg, cat.no, 1, tanka_rows=tanka)
    est2 = app.apply_modifications(est, {'overrides': {}, 'deleted': set(), 'new_items': [ni]})
    budget = convert_to_andpad(est2)
    item = next(i for i in budget.items if i.item_name == sg['name'])
    assert any('可動棚' in a.vendor for a in item.allocations), [(a.vendor, a.kind) for a in item.allocations]


# ---------------------------------------------------------------- フィードバック
def test_delete_confirmation_does_not_carry_over(app_dir):
    fb_path = os.path.join(app_dir, "projects", "sample_furui_feedbacks.json")
    with open(fb_path, "w", encoding="utf-8") as f:
        json.dump([{"category": "操作性", "body": "1件目", "timestamp": "2026-08-04 17:25", "name": "t", "status": "未対応",
                    "feedback_id": "fb-1", "attachments": []},
                   {"category": "操作性", "body": "2件目", "timestamp": "2026-08-04 17:31", "name": "t", "status": "未対応",
                    "feedback_id": "fb-2", "attachments": []}], f, ensure_ascii=False)
    at = _boot(app_dir)
    oks = [c for c in at.checkbox if c.key and "_delok" in c.key]
    assert len(oks) == 2
    oks[0].check().run()
    dels = [b for b in at.button if b.key and b.key.endswith("_del")]
    dels[0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    dels = [b for b in at.button if b.key and b.key.endswith("_del")]
    assert len(dels) == 1 and dels[0].disabled is True, "残った問い合わせの削除ボタンが確認なしで押せる"
    os.remove(fb_path)


def test_legacy_feedback_same_time_and_body_do_not_crash(app_dir):
    fb_path = os.path.join(app_dir, "projects", "sample_furui_feedbacks.json")
    with open(fb_path, "w", encoding="utf-8") as f:
        json.dump([{"category": "操作性", "body": "同じ", "timestamp": "2026-07-01 10:00", "status": "未対応"},
                   {"category": "その他", "body": "同じ", "timestamp": "2026-07-01 10:00", "status": "未対応"}], f, ensure_ascii=False)
    at = _boot(app_dir)
    assert not at.exception, [str(e.value) for e in at.exception]
    os.remove(fb_path)


# ---------------------------------------------------------------- PDFタブ
def test_pdf_edits_survive_widget_reset_and_floor_labels(app_dir):
    raw = {'total_floor_area_m2': 100.0, 'floors': 2, 'floor_areas': {'一階': 61.0, '二階': 39.0},
           'rooms': [{'name': 'LDK', 'floor': '1階', 'area_m2': 30.0}], 'fittings': []}
    edited = json.loads(json.dumps(raw))
    edited['total_floor_area_m2'] = 150.0
    edited['rooms'][0]['area_m2'] = 33.0
    at = _boot(app_dir)
    # 他の案件を見ていた間に入力欄の状態が消えた状態（直した値は data にだけある）
    at.session_state["proj_sample_furui_pdf_extracted"] = {"file": "t.pdf", "raw": raw, "data": edited}
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    data = at.session_state["proj_sample_furui_pdf_extracted"]["data"]
    assert data['total_floor_area_m2'] == 150.0 and data['rooms'][0]['area_m2'] == 33.0
    assert data['floor_areas'] == {'1階': 61.0, '2階': 39.0}


def test_pdf_basement_is_not_first_floor(app_dir):
    raw = {'total_floor_area_m2': 120.0, 'floors': 2, 'floor_areas': {'地下1階': 20.0, '1階': 60.0, '2階': 40.0},
           'rooms': [], 'fittings': []}
    at = _boot(app_dir)
    at.session_state["proj_sample_furui_pdf_extracted"] = {"file": "b.pdf", "raw": raw, "data": raw}
    at.run()
    data = at.session_state["proj_sample_furui_pdf_extracted"]["data"]
    assert data['floor_areas']['1階'] == 60.0 and data['floor_areas'].get('B1階') == 20.0
