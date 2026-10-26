"""2026-08-04 フィードバック2件の再現テスト（純粋ロジック）

① 打合せ変更タブ: 反映チェックボックスの「全て選択／全て解除」
   → 表で修正した内容（品名・金額など）を失わずに反映フラグだけを揃えること
② 変更見積（別プラン）の表示名: 元見積名の後ろにプラン名が連結されて切れる
   → プラン名を入力した案件は、選択欄ではプラン名のみを表示すること
"""
import math

import pytest

from ui_helpers import merge_proposal_edits, project_display_label, project_full_label

LABEL_TO_CTYPE = {'単価変更': 'price', '数量変更': 'quantity', '追加': 'add',
                  '削除': 'remove', '仕様変更': 'spec'}


def _proposals():
    return [
        {'category': '木工事', 'item_name': 'フローリング', 'change_type': 'spec',
         'proposed': '無垢材に変更', 'amount': 120000, 'confidence': 'high', 'reason': '議事録p2'},
        {'category': '住宅設備機器工事', 'item_name': 'キッチン', 'change_type': 'price',
         'proposed': '上位グレード', 'amount': None, 'confidence': 'low', 'reason': ''},
        {'category': '外構工事', 'item_name': 'カーポート', 'change_type': 'add',
         'proposed': '2台用を追加', 'amount': 0, 'confidence': 'medium', 'reason': '口頭'},
    ]


def _rows(apply=(False, False, False)):
    return [
        {'反映': apply[0], '工事種別': '木工事', '品名': 'フローリング', '変更の種類': '仕様変更',
         '変更後の内容': '無垢材に変更', '金額(円)': 120000, '読み取り': '🟢 高', '根拠': '議事録p2'},
        {'反映': apply[1], '工事種別': '住宅設備機器工事', '品名': 'キッチン', '変更の種類': '単価変更',
         '変更後の内容': '上位グレード', '金額(円)': 0, '読み取り': '🔴 低', '根拠': ''},
        {'反映': apply[2], '工事種別': '外構工事', '品名': 'カーポート', '変更の種類': '追加',
         '変更後の内容': '2台用を追加', '金額(円)': 0, '読み取り': '🟡 中', '根拠': '口頭'},
    ]


# ---------------------------------------------------------------- ① 全て選択／全て解除
class TestSelectAllClearAll:
    def test_select_all_sets_every_apply_flag(self):
        out = merge_proposal_edits(_proposals(), _rows(), LABEL_TO_CTYPE, apply_all=True)
        assert [p['apply'] for p in out] == [True, True, True]

    def test_clear_all_unsets_every_apply_flag(self):
        out = merge_proposal_edits(_proposals(), _rows((True, True, True)), LABEL_TO_CTYPE,
                                   apply_all=False)
        assert [p['apply'] for p in out] == [False, False, False]

    def test_edits_in_table_survive_select_all(self):
        """お客様ご指摘①の肝: 表で直した内容を失わずに全選択できること"""
        rows = _rows()
        rows[1]['品名'] = 'システムキッチン（修正）'
        rows[1]['金額(円)'] = 350000
        rows[1]['変更の種類'] = '数量変更'
        rows[2]['変更後の内容'] = '1台用に変更'
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=True)
        assert out[1]['item_name'] == 'システムキッチン（修正）'
        assert out[1]['amount'] == 350000
        assert out[1]['change_type'] == 'quantity'
        assert out[2]['proposed'] == '1台用に変更'
        assert all(p['apply'] for p in out)

    def test_apply_all_none_keeps_each_rows_checkbox(self):
        out = merge_proposal_edits(_proposals(), _rows((True, False, True)), LABEL_TO_CTYPE,
                                   apply_all=None)
        assert [p['apply'] for p in out] == [True, False, True]

    def test_untouched_keys_are_preserved(self):
        out = merge_proposal_edits(_proposals(), _rows(), LABEL_TO_CTYPE, apply_all=True)
        assert out[0]['confidence'] == 'high'
        assert out[2]['confidence'] == 'medium'

    def test_input_is_not_mutated(self):
        src = _proposals()
        snapshot = [dict(p) for p in src]
        merge_proposal_edits(src, _rows(), LABEL_TO_CTYPE, apply_all=True)
        assert src == snapshot

    def test_nan_amount_becomes_zero(self):
        rows = _rows()
        rows[0]['金額(円)'] = float('nan')
        rows[1]['金額(円)'] = None
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=True)
        assert out[0]['amount'] == 0 and out[1]['amount'] == 0

    def test_float_and_string_amounts_are_normalised_to_int(self):
        rows = _rows()
        rows[0]['金額(円)'] = 120000.0
        rows[1]['金額(円)'] = '98,000'
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=False)
        assert out[0]['amount'] == 120000 and isinstance(out[0]['amount'], int)
        assert out[1]['amount'] == 98000

    def test_negative_amount_is_clamped_to_zero(self):
        rows = _rows()
        rows[0]['金額(円)'] = -500
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=True)
        assert out[0]['amount'] == 0

    def test_unknown_change_type_label_keeps_original(self):
        rows = _rows()
        rows[0]['変更の種類'] = '不明なラベル'
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=True)
        assert out[0]['change_type'] == 'spec'

    def test_fewer_rows_than_proposals_keeps_remaining_proposals(self):
        out = merge_proposal_edits(_proposals(), _rows()[:2], LABEL_TO_CTYPE, apply_all=True)
        assert len(out) == 3
        assert out[2]['item_name'] == 'カーポート'
        assert out[2]['apply'] is True

    def test_more_rows_than_proposals_ignores_extra_rows(self):
        rows = _rows() + [{'反映': True, '工事種別': 'x', '品名': '余分', '変更の種類': '追加',
                           '変更後の内容': '', '金額(円)': 1, '読み取り': '', '根拠': ''}]
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=True)
        assert len(out) == 3

    def test_empty_inputs(self):
        assert merge_proposal_edits([], [], LABEL_TO_CTYPE, apply_all=True) == []
        assert merge_proposal_edits(None, None, LABEL_TO_CTYPE, apply_all=True) == []

    def test_nan_checkbox_is_treated_as_unchecked(self):
        rows = _rows()
        rows[0]['反映'] = float('nan')
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=None)
        assert out[0]['apply'] is False

    def test_missing_columns_fall_back_to_proposal_values(self):
        rows = [{'反映': True} for _ in range(3)]
        out = merge_proposal_edits(_proposals(), rows, LABEL_TO_CTYPE, apply_all=None)
        assert out[0]['item_name'] == 'フローリング'
        assert out[0]['amount'] == 120000
        assert out[1]['amount'] is None  # 元が未設定なら未設定のまま（表に列が無い）


# ---------------------------------------------------------------- ② プラン名の表示
class TestPlanDisplayLabel:
    def test_plan_name_only_when_plan_is_set(self):
        """お客様ご指摘②: 選択欄にはアップ時に入力したタイトルだけを出す"""
        pinfo = {'name': '山田太郎様邸新築工事320260601', 'plan': '山田様　打合せ変更 08/04'}
        assert project_display_label(pinfo) == '山田様　打合せ変更 08/04'

    def test_name_when_no_plan(self):
        assert project_display_label({'name': '山田太郎様邸新築工事', 'plan': None}) == '山田太郎様邸新築工事'
        assert project_display_label({'name': '山田太郎様邸新築工事'}) == '山田太郎様邸新築工事'

    def test_whitespace_only_plan_counts_as_no_plan(self):
        assert project_display_label({'name': 'A邸', 'plan': '   '}) == 'A邸'

    def test_label_is_stripped(self):
        assert project_display_label({'name': 'A邸', 'plan': '  B案 '}) == 'B案'
        assert project_display_label({'name': ' A邸 '}) == 'A邸'

    def test_empty_pinfo(self):
        assert project_display_label({}) == ''
        assert project_display_label({'name': None, 'plan': None}) == ''

    def test_full_label_keeps_both_for_history_messages(self):
        pinfo = {'name': 'A邸新築工事', 'plan': 'B案'}
        assert project_full_label(pinfo) == 'A邸新築工事（B案）'
        assert project_full_label({'name': 'A邸新築工事'}) == 'A邸新築工事'
