"""2026-08-07 MTG 宿題の再現テスト（純粋ロジック）

#1 販価変更時に原価を粗利率維持で連動  #2 寸法表記を数量・金額と誤認識しない  #3 FBステータス3段階
"""
import pytest

from ui_helpers import (FB_STATUSES, apply_fb_status_change, fb_status_color, linked_order_price,
                        normalize_fb_status, parse_amount_jp, strip_dimension_tokens,
                        validate_fb_status_change)


class TestLinkedOrderPrice:
    def test_keeps_items_own_cost_ratio(self):
        # 元: 見積260,000 / 原価174,200（原価率67%）→ 販価371,500 なら原価も67%
        assert linked_order_price(260000, 174200, 371500) == round(371500 * 174200 / 260000)

    def test_default_ratio_when_no_cost(self):
        assert linked_order_price(360000, 0, 514300) == round(514300 * 0.67)
        assert linked_order_price(0, 0, 100000) == 67000

    def test_zero_or_negative_new_price(self):
        assert linked_order_price(100, 67, 0) == 0
        assert linked_order_price(100, 67, -5) == 0

    def test_ratio_is_clamped_to_1(self):
        # 原価が販価を上回っていた明細でも、新原価が新販価を超えない
        assert linked_order_price(100000, 150000, 200000) == 200000

    def test_bad_inputs(self):
        assert linked_order_price(None, None, None) == 0
        assert linked_order_price('abc', 'x', 1000) == 670
        assert linked_order_price(1000, 700, 'abc') == 0

    def test_result_is_int(self):
        assert isinstance(linked_order_price(1000.0, 700.0, 1234.5), int)


class TestDimensionTokens:
    @pytest.mark.parametrize('text,expected', [
        ('パナソニック Sクラス カップボード W2700', None),
        ('W2700×D650×H850', None),
        ('カップボード W2700 1台', 1.0),
        ('φ100 配管 3本', 3.0),
        ('2700mm 幅のカウンター 2台', 2.0),
        ('150万', 1500000.0),
        ('1,500,000円', 1500000.0),
        ('２台', 2.0),
        ('W2700 カップボード 350,000円', 350000.0),
        ('見積 371,500円（W2700）', 371500.0),
        ('', None),
        (None, None),
        (2700, 2700.0),
        (float('nan'), None),
        (True, None),
    ])
    def test_parse(self, text, expected):
        assert parse_amount_jp(text) == expected

    def test_strip_keeps_plain_numbers(self):
        assert strip_dimension_tokens('2台 150万') == '2台 150万'
        assert '2700' not in strip_dimension_tokens('W2700')
        assert '850' not in strip_dimension_tokens('H 850')


class TestFeedbackStatusFlow:
    def test_statuses(self):
        assert FB_STATUSES == ['未対応', '着手中', '高井対応済み', '対応完了', '差し戻し']

    def test_legacy_mapping(self):
        assert normalize_fb_status('対応済み') == '対応完了'
        assert normalize_fb_status(None) == '未対応'
        assert normalize_fb_status('') == '未対応'
        assert normalize_fb_status('なにか') == '未対応'
        for s in FB_STATUSES:
            assert normalize_fb_status(s) == s

    def test_colors_are_distinct(self):
        assert len({fb_status_color(s) for s in FB_STATUSES}) == 5

    def test_returned_requires_note(self):
        assert validate_fb_status_change('差し戻し', '') is not None
        assert validate_fb_status_change('差し戻し', '   ') is not None
        assert validate_fb_status_change('差し戻し', '金額が違う') is None
        assert validate_fb_status_change('高井対応済み', '') is None
        assert validate_fb_status_change('不正', 'x') is not None

    def test_apply_keeps_history_and_does_not_mutate(self):
        fb = {'body': 'x', 'status': '未対応'}
        out = apply_fb_status_change(fb, '高井対応済み', '全選択ボタンを追加', '高井', '2026-10-26 15:00')
        assert fb['status'] == '未対応'
        assert out['status'] == '高井対応済み'
        assert out['status_note'] == '全選択ボタンを追加'
        assert out['status_history'][-1] == {'status': '高井対応済み', 'note': '全選択ボタンを追加',
                                             'by': '高井', 'at': '2026-10-26 15:00'}
        out2 = apply_fb_status_change(out, '差し戻し', 'まだ直っていない', '中村', '2026-10-27 09:00')
        assert [h['status'] for h in out2['status_history']] == ['高井対応済み', '差し戻し']
        assert len(out['status_history']) == 1
