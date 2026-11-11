"""2026-11-12 反証3回目で確認された問題の再現テスト（第7群）"""
import os

import pytest

from tests.regression.test_feedback_20260804_app import ROOT, app_dir, _no_cloud  # noqa: F401
from ui_helpers import (apply_note_table, design_category_no, design_item_to_new_item, merge_proposal_edits,
                        proposal_delta, proposal_from_row, proposal_table_values, proposal_value, special_note_key,
                        split_legacy_note)

LABEL_TO_CTYPE = {'単価変更': 'price', '数量変更': 'quantity', '追加': 'add', '削除': 'remove', '仕様変更': 'spec'}
DESIGN_XLSX = os.path.expanduser("~/Downloads/単価マスタ.xlsx")


def _tanka():
    import master_lookup
    with open(os.path.join(ROOT, "tankamaster_updated.csv"), "rb") as f:
        return master_lookup.parse_price_master(f.read(), "tankamaster_updated.csv")[0]


def _pairs():
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    est = calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                             load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")))
    return [(c.no, c.name) for c in est.categories if c.no != 22]


# ---------------------------------------------------------------- [0] 全て選択で直した数字を捨てない
def _row(p, **kw):
    init = proposal_table_values(p)
    r = {'反映': False, '工事種別': p['category'], '品名': p['item_name'],
         '変更の種類': {v: k for k, v in LABEL_TO_CTYPE.items()}[p['change_type']],
         '変更後の内容': p['proposed'], '金額(円)': init['金額(円)'], '数量': init['数量'], '根拠': ''}
    r.update(kw)
    return r


@pytest.mark.parametrize("p,edit,expected", [
    ({'change_type': 'price', 'category': '住宅設備機器工事', 'item_name': 'UB 1616', 'proposed': 'UB 371,500円'},
     {'金額(円)': 398000}, 398000),
    ({'change_type': 'quantity', 'category': '電気設備工事', 'item_name': '一般ｺﾝｾﾝﾄ', 'proposed': '3ヶ所'},
     {'数量': 4.0}, 4),
])
def test_select_all_keeps_user_number(p, edit, expected):
    merged = merge_proposal_edits([p], [_row(p, **edit)], LABEL_TO_CTYPE, apply_all=True)
    merged2 = merge_proposal_edits(merged, [_row(merged[0])], LABEL_TO_CTYPE, apply_all=True)   # 2回続けて押す
    for m in (merged, merged2):
        row = _row(m[0], 反映=True)
        out = proposal_from_row(m[0], row, LABEL_TO_CTYPE)
        assert proposal_value(out['change_type'], out['proposed'], out['amount']) == expected


# ---------------------------------------------------------------- [1] 絶対額の書き方を増減額にしない
@pytest.mark.parametrize("ctype,text,value", [
    ("price", "値引き後 350,000円", 350000), ("spec", "同等品（差額なし） 371,500円", 371500),
    ("price", "値下げして35万円に", 350000), ("spec", "プラスチック製に変更 12,000円", 12000),
    ("price", "値上げ後 400,000円", 400000), ("price", "差額込み 400,000円", 400000), ("price", "増額なし 371,500円", 371500),
    ("price", "キッチン 514,300円を値引きして480,000円に", 480000), ("price", "出精値引き後 350,000円", 350000),
    ("spec", "398,000円 アップグレード品", 398000),
])
def test_absolute_amounts_are_not_deltas(ctype, text, value):
    assert proposal_delta(ctype, text) is None
    assert proposal_value(ctype, text, None) == value


@pytest.mark.parametrize("ctype,text,delta", [
    ("price", "値引き 30,000円", -30000), ("spec", "差額 +85,000円", 85000), ("price", "2万円アップ", 20000),
    ("spec", "プラス 5,000円", 5000), ("price", "減額：12,000円", -12000),
])
def test_real_deltas_still_read(ctype, text, delta):
    assert proposal_delta(ctype, text) == delta


# ---------------------------------------------------------------- [2] 後ろ側に金額が無ければ文全体で読む
@pytest.mark.parametrize("ctype,text,value", [
    ("price", "UB 371,500円をサザナ1620に", 371500), ("spec", "UB 371,500円をサザナ1620に変更", 371500),
    ("spec", "グレードアップ 398,000円 → 2台", 398000), ("add", "ルームエアコン 98,000円を2台追加", 98000),
    ("price", "398,000円 → 1620サイズ", 398000), ("price", "キッチン 1,250,000円からI型2550に変更", 1250000),
    ("price", "360000から380000に", 380000), ("price", "UB 26万→371,500円", 371500),
])
def test_after_change_needs_usable_value(ctype, text, value):
    assert proposal_value(ctype, text, None) == value


# ---------------------------------------------------------------- [3] 桁区切りの千円
@pytest.mark.parametrize("text,value", [("1,980千円", 1980000), ("12,500千円", 12500000), ("外構工事 1,250千円", 1250000)])
def test_thousand_yen_with_commas(text, value):
    assert proposal_value("add", text, None) == value


# ---------------------------------------------------------------- [4][5][6] 旧形式キーの差異理由
LEGACY = '木工事|標準外工事'


def _note_rows(*specs):
    return [{'_key': key, '工事種別': '木工事', '品名': '標準外工事', '摘要': sm, '分類': cat, '差異の理由': reason,
             '適用範囲': '全案件共通', '差額': 1000} for key, sm, cat, reason in specs]


def test_legacy_key_rows_are_split_and_edits_kept():
    notes = {LEGACY: {'category': 'その他', 'reason': '旧', 'scope': 'global'}}
    rows = _note_rows((LEGACY, '横引き', 'その他', '旧'), (LEGACY, '曲がり', 'その他', '新'))
    new, entries = apply_note_table(notes, rows, [LEGACY, LEGACY], 'now')
    assert LEGACY not in new
    assert new[special_note_key('木工事', '標準外工事', '横引き')]['reason'] == '旧'
    assert new[special_note_key('木工事', '標準外工事', '曲がり')]['reason'] == '新'


def test_legacy_key_row_cleared_only_for_that_row():
    notes = {LEGACY: {'category': 'その他', 'reason': '旧', 'scope': 'global'}}
    rows = _note_rows((LEGACY, '横引き', 'その他', '旧'), (LEGACY, '曲がり', '', ''))
    new, _ = apply_note_table(notes, rows, [LEGACY, LEGACY], 'now')
    assert LEGACY not in new and special_note_key('木工事', '標準外工事', '曲がり') not in new
    assert new[special_note_key('木工事', '標準外工事', '横引き')]['reason'] == '旧'


def test_legacy_key_kept_when_some_readers_are_hidden():
    notes = {LEGACY: {'category': 'その他', 'reason': '旧', 'scope': 'global'}}
    rows = _note_rows((LEGACY, '曲がり', 'その他', '新'))
    new, _ = apply_note_table(notes, rows, [LEGACY, LEGACY], 'now')
    assert new[LEGACY]['reason'] == '旧'
    assert new[special_note_key('木工事', '標準外工事', '曲がり')]['reason'] == '新'


def test_same_new_key_prefers_changed_row():
    k = special_note_key('木工事', '標準外工事', 'x')
    notes = {k: {'category': 'その他', 'reason': '旧', 'scope': 'global'}}
    rows = _note_rows((k, 'x', 'その他', '旧'), (k, 'x', 'その他', '新'))
    new, _ = apply_note_table(notes, rows, [k, k], 'now')
    assert new[k]['reason'] == '新'


def test_split_legacy_note_moves_reason_to_each_row():
    notes = {'基礎工事|人通口補強': {'category': 'その他', 'reason': 'R旧', 'scope': 'global'}}
    out = split_legacy_note(notes, '基礎工事', '人通口補強', ['60㎡超過分', '別仕様'])
    assert '基礎工事|人通口補強' not in out
    assert out[special_note_key('基礎工事', '人通口補強', '60㎡超過分')]['reason'] == 'R旧'
    assert out[special_note_key('基礎工事', '人通口補強', '別仕様')]['reason'] == 'R旧'
    assert notes == {'基礎工事|人通口補強': {'category': 'その他', 'reason': 'R旧', 'scope': 'global'}}  # 元は変えない


# ---------------------------------------------------------------- [7] 本番どおりの分類で追加先を決める
@pytest.mark.parametrize("name,category,expect", [
    ("可動棚5枚　W3.0尺・D450まで", "可動棚", "内部建具"), ("ｶｳﾝﾀｰ W2.0m以下", "カウンター", "造作家具"),
    ("ニッチ", "造作", "木工"), ("【標準】1階床 無垢フローリング", "床材", "木工"), ("室内物干し金物", "金物", "雑工事"),
    ("クロス 量産品", "クロス", "内装"), ("間接照明 折上げ天井", "造作", "木工"),
])
def test_design_category_with_real_group(name, category, expect):
    pairs = _pairs()
    sg = {'name': name, 'category': category, 'group': '6.内装オプション工事'}
    no = design_category_no(sg, pairs, tanka_rows=_tanka())
    assert expect in dict(pairs)[no], dict(pairs)[no]


# ---------------------------------------------------------------- [8][9] ANDPADの振り分け
def test_exact_name_match_keeps_vendor_even_with_breakdown():
    design_rows = [{'code': '6-KD-001', 'name': '可動棚5枚 W3.0尺･D450まで', 'order_price': 30000, 'estimate_price': 45000},
                   {'code': '6-KD-001-1', 'name': '内訳①: [材] 棚板', 'order_price': 20000, 'breakdown_of': '6-KD-001'},
                   {'code': '6-KD-001-2', 'name': '内訳②: [工] 取付費', 'order_price': 10000, 'breakdown_of': '6-KD-001'}]
    sg = {'code': '6-KD-001', 'name': '可動棚5枚 W3.0尺･D450まで', 'unit': 'ヶ所', 'estimate_price': 45000, 'order_price': 30000}
    ni = design_item_to_new_item(sg, 8, 1, tanka_rows=_tanka(), design_rows=design_rows)
    assert ni.get('andpad_code') == 80001, ni


def test_close_candidates_with_same_vendor_still_match():
    sg = {'code': '7', 'name': 'UB 1618', 'unit': '', 'estimate_price': 380000, 'order_price': 260000}
    ni = design_item_to_new_item(sg, 14, 1, tanka_rows=_tanka())
    assert ni.get('andpad_code') in (140006, 140007), ni


def test_negative_breakdown_rows_keep_labor_amount():
    design_rows = [{'code': 'P', 'name': 'スリット階段フル＜標準差額＞', 'order_price': 119100},
                   {'code': 'P-1', 'name': '内訳①: [材] 階段', 'order_price': 205100, 'breakdown_of': 'P'},
                   {'code': 'P-2', 'name': '内訳②: [工] 取付費', 'order_price': 20000, 'breakdown_of': 'P'},
                   {'code': 'P-3', 'name': '内訳③: 【標準】階段 ＜マイナス＞', 'order_price': -106000, 'breakdown_of': 'P'}]
    sg = {'code': 'P', 'name': 'スリット階段フル＜標準差額＞', 'unit': '式', 'estimate_price': 178000, 'order_price': 119100}
    ni = design_item_to_new_item(sg, 3, 1, design_rows=design_rows)
    parts = ni['andpad_mapping']['parts']
    assert abs(parts[0]['ratio'] * 119100 - 20000) < 1, parts


# ---------------------------------------------------------------- [10] 親の無い内訳行
@pytest.mark.skipif(not os.path.exists(DESIGN_XLSX), reason="設計部マスタのファイルが無い")
def test_orphan_breakdown_rows_not_attached_to_previous_parent():
    import master_lookup
    from ui_helpers import design_breakdown_rows
    with open(DESIGN_XLSX, "rb") as f:
        rows, _ = master_lookup.parse_price_master(f.read(), "単価マスタ.xlsx")
    assert len(design_breakdown_rows(rows, '6-MK-018')) <= 3


def test_drilldown_split_keeps_legacy_reason_for_each_row():
    from ui_helpers import detail_row_note_target, split_legacy_for_rows
    rows = [{'AI品名': '深基礎', 'AI摘要': '高低差450㎜以下', '実績品名': '深基礎', '実績摘要': '450以下', '区分': '金額が違う'},
            {'AI品名': '深基礎', 'AI摘要': '高低差600㎜', '実績品名': '', '実績摘要': '', '区分': 'AIにしかない'}]
    notes = {'基礎工事|深基礎': {'category': '現場条件・特殊要因', 'reason': '敷地の高低差（施主要望）', 'scope': 'global'}}
    out = split_legacy_for_rows(notes, '基礎工事', rows[0], rows)
    assert '基礎工事|深基礎' not in out
    nm, sm = detail_row_note_target(rows[0])
    assert out[special_note_key('基礎工事', nm, sm)]['reason'] == '敷地の高低差（施主要望）'
    assert out[special_note_key('基礎工事', '深基礎', '高低差600㎜')]['reason'] == '敷地の高低差（施主要望）'


def test_apply_after_select_all_uses_user_number_in_app(app_dir):
    """全て選択で保存された利用者の数字（amount_auto=False）が、反映時に使われる"""
    from tests.regression.test_feedback_20260804_app import _boot
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    est = calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                             load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")))
    cat, idx, item = next((c, i, it) for c in est.categories for i, it in enumerate(c.items) if it.estimate_price > 100000)
    at = _boot(app_dir)
    at.session_state["proj_sample_furui_ai_proposals"] = [
        {'category': cat.name, 'item_name': item.name, 'change_type': 'price', 'proposed': f"{item.name} 371,500円",
         'amount': 398000, 'amount_auto': False, 'confidence': 'high', 'apply': True, 'reason': ''}]
    at.session_state["proj_sample_furui_ai_proposals_src"] = "音声変更"
    at.run()
    [c for c in at.checkbox if c.key == "props_to_new_plan"][0].uncheck().run()
    [b for b in at.button if b.key == "btn_apply_props"][0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    assert at.session_state["proj_sample_furui_mods"]['overrides'][f"{cat.no}_{idx}"]['estimate_price'] == 398000
