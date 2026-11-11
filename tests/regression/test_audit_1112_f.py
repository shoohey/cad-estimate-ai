"""2026-11-12 反証2回目で確認された問題の再現テスト（第6群）"""
import json
import os

import pytest

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401
from ui_helpers import (alias_pairs_from_library, category_detail_rows, clean_editor_records, design_category_no,
                        design_item_to_new_item, find_feedback_index, library_keys_to_withdraw, make_alias_entry,
                        merge_proposal_edits, parse_amount_jp, parse_special_note_key, pdf_floor_areas,
                        proposal_delta, proposal_from_row, proposal_number, proposal_table_values, proposal_value,
                        resolve_note_key, search_design_master, special_note_key)

LABEL_TO_CTYPE = {'単価変更': 'price', '数量変更': 'quantity', '追加': 'add', '削除': 'remove', '仕様変更': 'spec'}


def _estimate():
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    return calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                              load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")))


# ---------------------------------------------------------------- [0] 増減額は単価そのものにしない
@pytest.mark.parametrize("ctype,text,delta", [
    ("price", "2万円アップ", 20000), ("spec", "グレードダウン △50,000円", -50000),
    ("spec", "エコキュート 460L 差額 +85,000円", 85000), ("spec", "カップボード W2700→W1800 -50,000円", -50000),
    ("price", "▲3,000円", -3000), ("spec", "減額 5万円", -50000),
])
def test_delta_is_not_absolute(ctype, text, delta):
    assert proposal_value(ctype, text, None) is None
    assert proposal_delta(ctype, text) == delta
    assert proposal_number(ctype, text, None) == 0


def test_delta_markers_that_are_not_deltas():
    assert proposal_delta("spec", "グレードアップ 398,000円") is None
    assert proposal_value("spec", "グレードアップ 398,000円", None) == 398000
    assert proposal_value("add", "+98,000円", None) == 98000      # 追加の「+」は金額そのもの
    assert proposal_delta("price", "371,500円") is None


def test_delta_applied_to_current_price():
    import app
    est = _estimate()
    cat, idx, item = next((c, i, it) for c in est.categories for i, it in enumerate(c.items)
                          if it.estimate_price > 100000)
    mods = {}
    app._apply_ai_proposal(est, mods, {'change_type': 'spec', 'category': cat.name, 'item_name': item.name,
                                       'proposed': f"{item.name} 差額 +85,000円", 'amount': 0})
    assert mods['overrides'][f"{cat.no}_{idx}"]['estimate_price'] == item.estimate_price + 85000
    mods = {}
    app._apply_ai_proposal(est, mods, {'change_type': 'price', 'category': cat.name, 'item_name': item.name,
                                       'proposed': "△50,000円", 'amount': 0})
    assert mods['overrides'][f"{cat.no}_{idx}"]['estimate_price'] == item.estimate_price - 50000


# ---------------------------------------------------------------- [2] 漢数字の位取り
@pytest.mark.parametrize("text,value", [
    ("9万8千円", 98000), ("51万4300円", 514300), ("10万5000円", 105000), ("1万2千5百円", 12500),
    ("98千円", 98000), ("8.5万円", 85000), ("37万8千円", 378000), ("51万4,300円", 514300),
])
def test_kanji_amounts(text, value):
    assert proposal_value("price", text, None) == value
    assert parse_amount_jp(text) == value


def test_man_followed_by_count_is_not_merged():
    assert proposal_value("add", "10万 2台", None) == 100000
    assert proposal_value("quantity", "10万円のものを2台", None) == 2


# ---------------------------------------------------------------- [3] 変更前→変更後は後ろを採用
@pytest.mark.parametrize("ctype,text,value", [
    ("price", "UB 26万→371,500円", 371500), ("price", "360,000円→514,300円", 514300),
    ("quantity", "3ヶ所→4ヶ所", 4), ("quantity", "30.5㎡→32.0㎡", 32), ("quantity", "手すり 3.6m→4.5m", 4.5),
    ("quantity", "現状3ヶ所を4ヶ所に", 4), ("price", "51万4300円から60万円に", 600000),
    ("quantity", "手すりを2箇所追加", 2), ("quantity", "数量を0に", 0),
])
def test_after_change_value(ctype, text, value):
    assert proposal_value(ctype, text, None) == value


# ---------------------------------------------------------------- [4] 単価変更の単位なし数字
@pytest.mark.parametrize("text,value", [
    ("2階 280000", 280000), ("エコキュート 370L 450000", 450000), ("100V 3000", 3000),
    ("I型2550 514300", 514300), ("サザナ 1616サイズ 371500", 371500), ("2台 98000", 98000),
    ("30.5㎡ 単価 3500", 3500), ("UB 371500", 371500), ("WC 128000", 128000), ("514,300", 514300),
])
def test_price_bare_number(text, value):
    assert proposal_value("price", text, None) == value


def test_price_bare_number_needs_a_candidate():
    assert proposal_value("price", "2階 エコキュート 370L", None) is None


# ---------------------------------------------------------------- [5][6][7][9] 数量の単位
@pytest.mark.parametrize("text,value", [
    ("2階 大工 3人工", 3), ("1階 コンクリート 12.5㎥", 12.5), ("1F 土間コン 12.5m3", 12.5), ("2階 照明 5灯", 5),
    ("2階 手すり 3.6M", 3.6), ("2階 防水 3.5M2", 3.5), ("2階 床 30.5", 30.5), ("2階 WC 1箇所", 1),
    ("UB 2台", 2), ("IH 1台", 1), ("WC 2箇所", 2), ("ＵＢ １台", 1), ("棚板 W900 x3枚", 3), ("コンセント x5ヶ所", 5),
    ("高さ2.4m 3本", 3), ("幅1.8m 3本", 3), ("間口2.7m 1台", 1), ("L=4m 10本", 10), ("W 1.8m 2枚", 2),
    ("照明 LED5台", 5), ("2F5台", 5), ("窓枠 高さ2.0m 10本", 10),
    # 前回直したものは保つ
    ("軒樋 L=12m", 12), ("長さ12m", 12), ("TKS05305JA を 2 に", 2), ("カップボード W2700 1台", 1),
])
def test_quantity_units(text, value):
    assert proposal_value("quantity", text, None) == value


# ---------------------------------------------------------------- [1][8] 表の数字（数量は小数・内容だけ直したら読み直す）
def test_quantity_keeps_decimals_in_table():
    p = {'change_type': 'quantity', 'category': '板金・樋工事', 'item_name': '軒樋', 'proposed': '軒樋 30.5m'}
    init = proposal_table_values(p)
    assert init['数量'] == 30.5
    row = {'反映': True, '工事種別': '板金・樋工事', '品名': '軒樋', '変更の種類': '数量変更',
           '変更後の内容': '軒樋 30.5m', '金額(円)': init['金額(円)'], '数量': init['数量'], '根拠': ''}
    out = proposal_from_row(p, row, LABEL_TO_CTYPE)
    assert proposal_value(out['change_type'], out['proposed'], out['amount']) == 30.5


def test_edited_text_is_reread_when_number_untouched():
    p = {'change_type': 'quantity', 'category': 'x', 'item_name': 'ﾀｵﾙ掛け', 'proposed': '3ヶ所'}
    init = proposal_table_values(p)
    row = {'反映': True, '工事種別': 'x', '品名': 'ﾀｵﾙ掛け', '変更の種類': '数量変更', '変更後の内容': '5ヶ所',
           '金額(円)': init['金額(円)'], '数量': init['数量'], '根拠': ''}
    out = proposal_from_row(p, row, LABEL_TO_CTYPE)
    assert proposal_value(out['change_type'], out['proposed'], out['amount']) == 5
    # 種類を単価変更に直した
    row2 = dict(row, **{'変更の種類': '単価変更', '変更後の内容': '7,200円'})
    out2 = proposal_from_row(p, row2, LABEL_TO_CTYPE)
    assert proposal_value(out2['change_type'], out2['proposed'], out2['amount']) == 7200
    # 数字を自分で直したときはその数字
    row3 = dict(row, **{'数量': 6.0})
    out3 = proposal_from_row(p, row3, LABEL_TO_CTYPE)
    assert proposal_value(out3['change_type'], out3['proposed'], out3['amount']) == 6


def test_select_all_keeps_untouched_number_as_automatic():
    p = {'change_type': 'price', 'category': 'x', 'item_name': 'UB', 'proposed': '371,500円'}
    init = proposal_table_values(p)
    rows = [{'反映': False, '工事種別': 'x', '品名': 'UB', '変更の種類': '単価変更', '変更後の内容': '398,000円',
             '金額(円)': init['金額(円)'], '数量': init['数量'], '根拠': ''}]
    merged = merge_proposal_edits([p], rows, LABEL_TO_CTYPE, apply_all=True)
    assert proposal_table_values(merged[0])['金額(円)'] == 398000


# ---------------------------------------------------------------- [10][11][12][14] 差異理由のキー
def test_special_key_with_empty_summary_is_distinct_from_legacy():
    assert special_note_key('木工事', '床', '') != '木工事|床'
    assert parse_special_note_key(special_note_key('木工事', '床', '')) == ('木工事', '床', '')
    rows = category_detail_rows(
        {'matched': [], 'actual_only': [],
         'ai_only': [{'name': '床', 'summary': '', 'amount': 100, 'quantity': 1, 'unit': '式'},
                     {'name': '床', 'summary': '塗装費のみ', 'amount': 50, 'quantity': 1, 'unit': '式'}]},
        {('床', ''): '案件特有'})
    by = {(r['AI品名'], r['AI摘要']): r['扱い'] for r in rows}
    assert by[('床', '')] == '案件特有' and by[('床', '塗装費のみ')] == ''


def test_library_withdraw_keys_use_name_key_and_keep_other_rows():
    new_notes = {special_note_key('木工事', '床', '塗装費のみ'): {'category': 'その他', 'reason': '理由', 'scope': 'global'}}
    withdrawn = [special_note_key('木工事', '床', 'オーク'), special_note_key('基礎工事', '人通口', '')]
    keys = library_keys_to_withdraw(withdrawn, new_notes)
    assert keys == ['基礎工事|人通口'], keys


def test_reason_table_key_resolution():
    """差異理由の表は、読んだキーに書く（摘要付きの登録を品名だけのキーへ複製しない）"""
    reg = special_note_key('木工事', '床ｶﾞﾗﾘ変更', '樹脂→オーク')
    notes = {reg: {'category': '案件特有', 'reason': 'x', 'scope': 'project'}}
    # 登録した行（金額差の行は AI摘要・実績摘要のどちらで登録されていても見つける）
    assert resolve_note_key(notes, '木工事', '床ｶﾞﾗﾘ変更', ['別', '樹脂→オーク']) == reg
    # 登録していない同じ品名の行は、自分の摘要の新しいキー
    assert resolve_note_key(notes, '木工事', '床ｶﾞﾗﾘ変更', ['塗装費のみ']) == special_note_key('木工事', '床ｶﾞﾗﾘ変更', '塗装費のみ')
    # 旧形式（品名だけ）で登録済みならそのキー
    assert resolve_note_key({'木工事|床': {'reason': 'y'}}, '木工事', '床', ['a']) == '木工事|床'


# ---------------------------------------------------------------- [13] 名称だけの組が完全一致を奪わない
def test_name_fallback_does_not_steal_exact_name_match():
    entry = make_alias_entry({'AI品名': 'KACU', 'AI摘要': '旧'}, {'実績品名': 'ﾎｽｸﾘｰﾝ', '実績摘要': '旧'},
                             '外部建具工事', 'A邸', '2026-11-12 10:00')
    ai = [{'name': 'KACU', 'summary': '新', 'amount': 1000, 'quantity': 1, 'unit': '式'}]
    ac = [{'name': 'KACU', 'summary': '別の摘要', 'amount': 1000, 'quantity': 1, 'unit': '式'},
          {'name': 'ﾎｽｸﾘｰﾝ', 'summary': '新', 'amount': 900, 'quantity': 1, 'unit': '式'}]
    out = excel_diff._match_items(ai, ac, alias_pairs_from_library([entry]))
    pairs = [(m['ai_name'], m['actual_name']) for m in out['matched']]
    assert pairs == [('KACU', 'KACU')], pairs


# ---------------------------------------------------------------- [16] FBの照合
def test_find_feedback_index_uses_category_and_id():
    fbs = [{'category': '操作性', 'body': '同じ', 'timestamp': 't'}, {'category': 'その他', 'body': '同じ', 'timestamp': 't'}]
    assert find_feedback_index(fbs, {'category': 'その他', 'body': '同じ', 'timestamp': 't'}) == 1
    fbs2 = [{'feedback_id': 'a', 'body': 'x', 'timestamp': 't'}, {'feedback_id': 'b', 'body': 'x', 'timestamp': 't'}]
    assert find_feedback_index(fbs2, {'feedback_id': 'b', 'body': 'x', 'timestamp': 't'}) == 1
    assert find_feedback_index(fbs2, {'feedback_id': 'c', 'body': 'x', 'timestamp': 't'}) is None


def test_delete_hits_the_right_feedback(app_dir):
    fb_path = os.path.join(app_dir, "projects", "sample_furui_feedbacks.json")
    with open(fb_path, "w", encoding="utf-8") as f:
        json.dump([{"category": "操作性", "body": "同じ", "timestamp": "2026-07-01 10:00", "status": "未対応"},
                   {"category": "その他", "body": "同じ", "timestamp": "2026-07-01 10:00", "status": "未対応"}], f,
                  ensure_ascii=False)
    at = _boot(app_dir)
    # 新しい順に並ぶので先頭のカードは「その他」
    oks = [c for c in at.checkbox if c.key and c.key.endswith("_delok")]
    oks[0].check().run()
    [b for b in at.button if b.key and b.key.endswith("_del")][0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    with open(fb_path, encoding="utf-8") as f:
        left = json.load(f)
    assert [x['category'] for x in left] == ['操作性'], left
    os.remove(fb_path)


# ---------------------------------------------------------------- [17] 案件削除でキャッシュも消す
def test_delete_project_clears_session_caches(tmp_path):
    import app
    from unittest import mock
    state = {'_diffnote_cache': {'p1': {'x|y|': {'reason': '旧'}}, 'p2': {}},
             '_actual_file_cache': {'p1': b'x'},
             '_store_proj:p1:hidden_feedback': {'x': 1}, '_store_proj:p2:hidden_feedback': {'y': 1}}
    with mock.patch.object(app.st, 'session_state', state), mock.patch.object(app, 'PROJECTS_DIR', str(tmp_path)), \
            mock.patch.object(app, '_cloud_ok', lambda: False):
        app.delete_project_storage('p1')
    assert 'p1' not in state['_diffnote_cache'] and 'p2' in state['_diffnote_cache']
    assert 'p1' not in state['_actual_file_cache']
    assert '_store_proj:p1:hidden_feedback' not in state and '_store_proj:p2:hidden_feedback' in state


# ---------------------------------------------------------------- [18][20] PDFの表と階
def test_clean_editor_records_restores_types():
    rows = [{'kind': 'FIX', 'name': 'W1', 'width_mm': float('nan'), 'height_mm': 1690.0, 'count': 2.0, 'location': None}]
    out = clean_editor_records(rows, int_cols=('width_mm', 'height_mm', 'count'))
    assert out == [{'kind': 'FIX', 'name': 'W1', 'width_mm': None, 'height_mm': 1690, 'count': 2, 'location': None}]


def test_pdf_floor_areas_skip_totals():
    fa, skipped = pdf_floor_areas({'1階': 60, '2階': 40, '1・2階合計': 100, '中2階': 8, '地下1階': 10, '延床': 110})
    assert fa == {'1階': 60.0, '2階': 40.0, 'B1階': 10.0}
    assert set(skipped) == {'1・2階合計', '中2階', '延床'}


def test_pdf_floor_warning_stays(app_dir):
    raw = {'total_floor_area_m2': 300.0, 'floors': 6, 'rooms': [], 'fittings': []}
    at = _boot(app_dir)
    at.session_state["proj_sample_furui_pdf_extracted"] = {"file": "f.pdf", "raw": raw, "data": json.loads(json.dumps(raw))}
    at.run()
    at.run()
    assert any('6' in w.value and '4' in w.value for w in at.warning), [w.value for w in at.warning]


# ---------------------------------------------------------------- [21][23][24][25] 設計部マスタ
def test_search_results_carry_group():
    rows = [{'code': '4', 'category': '引込み戸', 'group': '4.内部建具工事', 'name': '標準引込み戸 上吊 ﾊﾟﾈﾙﾀｲﾌﾟ',
             'summary': '', 'unit': 'ヶ所', 'estimate_price': 80000, 'order_price': 50000}]
    hit = search_design_master(rows, '引込み戸')[0]
    assert hit.get('group') == '4.内部建具工事'
    est = _estimate()
    pairs = [(c.no, c.name) for c in est.categories]
    assert dict(pairs)[design_category_no(hit, pairs)].find('内部建具') >= 0


@pytest.mark.parametrize("name,category,expect", [
    ("【標準】SGL鋼板 たてひら葺", "屋根", "屋根"), ("畳 縁なし", "畳", "内装"), ("【標準】ﾃﾚﾋﾞﾄﾞｱﾎﾝ", "ドアホン", "電気"),
])
def test_design_category_hints(name, category, expect):
    est = _estimate()
    pairs = [(c.no, c.name) for c in est.categories]
    no = design_category_no({'name': name, 'category': category, 'group': ''}, pairs)
    assert expect in dict(pairs)[no], dict(pairs)[no]


def test_labor_rows_are_labor_in_andpad():
    sg = {'name': '造作洗面化粧台：取付費', 'summary': '', 'unit': 'ヶ所', 'estimate_price': 25000, 'order_price': 16750,
          'code': '7', 'category': '洗面', 'cost_kind': '工事費'}
    import master_lookup
    with open(os.path.join(ROOT, "tankamaster_updated.csv"), "rb") as f:
        tanka, _ = master_lookup.parse_price_master(f.read(), "tankamaster_updated.csv")
    ni = design_item_to_new_item(sg, 14, 1, tanka_rows=tanka)
    assert 'andpad_code' not in ni and ni.get('andpad_mapping', {}).get('kind') == '工', ni


# ---------------------------------------------------------------- 画面: 打合せ変更の反映（小数の数量・増減額）
def test_apply_quantity_decimal_and_delta_in_app(app_dir):
    at = _boot(app_dir)
    est = _estimate()
    cat, idx, item = next((c, i, it) for c in est.categories for i, it in enumerate(c.items) if it.estimate_price > 100000)
    gutter = next((c, i, it) for c in est.categories for i, it in enumerate(c.items) if '軒樋' in it.name)
    at.session_state["proj_sample_furui_ai_proposals"] = [
        {'category': gutter[0].name, 'item_name': gutter[2].name, 'change_type': 'quantity',
         'proposed': '軒樋 30.5m', 'confidence': 'high', 'apply': True, 'reason': ''},
        {'category': cat.name, 'item_name': item.name, 'change_type': 'price',
         'proposed': '△50,000円', 'confidence': 'high', 'apply': True, 'reason': ''},
    ]
    at.session_state["proj_sample_furui_ai_proposals_src"] = "音声変更"
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    [c for c in at.checkbox if c.key == "props_to_new_plan"][0].uncheck().run()
    [b for b in at.button if b.key == "btn_apply_props"][0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    ov = at.session_state["proj_sample_furui_mods"]['overrides']
    assert ov[f"{gutter[0].no}_{gutter[1]}"]['quantity'] == 30.5
    assert ov[f"{cat.no}_{idx}"]['estimate_price'] == item.estimate_price - 50000
