"""実績比較: 工事種別ごとの明細ドリルダウン（純粋ロジック）"""
from ui_helpers import (DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY, DETAIL_KIND_DIFF, DETAIL_KIND_SAME,
                        category_detail_rows, detail_header, detail_kind_counts)

ITEM_DIFF = {
    'matched': [
        {'ai_name': '分電盤', 'actual_name': '分電盤', 'ai_amount': 50000, 'actual_amount': 50000, 'diff': 0,
         'ai_quantity': 1, 'actual_quantity': 1, 'unit': '面'},
        {'ai_name': 'コンセント', 'actual_name': 'コンセント工事', 'ai_amount': 120000, 'actual_amount': 135000,
         'diff': -15000, 'ai_quantity': 40, 'actual_quantity': 45, 'unit': '箇所'},
        {'ai_name': 'スイッチ', 'actual_name': 'スイッチ', 'ai_amount': 30000, 'actual_amount': 28000, 'diff': 2000},
    ],
    'ai_only': [{'name': '太陽光配線', 'amount': 80000, 'quantity': 1, 'unit': '式'}],
    'actual_only': [{'name': 'EV用コンセント', 'amount': 25000}, {'name': '仮設電気', 'amount': 15000}],
}


class TestDetailRows:
    def test_kinds_and_order(self):
        rows = category_detail_rows(ITEM_DIFF)
        assert [r['区分'] for r in rows] == [DETAIL_KIND_DIFF, DETAIL_KIND_DIFF, DETAIL_KIND_AI_ONLY,
                                             DETAIL_KIND_ACTUAL_ONLY, DETAIL_KIND_ACTUAL_ONLY, DETAIL_KIND_SAME]
        # 区分内は差額の大きい順
        assert rows[0]['AI品名'] == 'コンセント' and rows[0]['差額(AI−実績)'] == -15000
        assert rows[1]['AI品名'] == 'スイッチ' and rows[1]['差額(AI−実績)'] == 2000
        assert rows[3]['実績品名'] == 'EV用コンセント' and rows[3]['差額(AI−実績)'] == -25000

    def test_quantity_text_and_blank_sides(self):
        rows = category_detail_rows(ITEM_DIFF)
        assert rows[0]['AI数量'] == '40箇所' and rows[0]['実績数量'] == '45箇所'
        ai_only = [r for r in rows if r['区分'] == DETAIL_KIND_AI_ONLY][0]
        assert ai_only['実績品名'] == '' and ai_only['実績金額'] == 0 and ai_only['AI数量'] == '1式'
        actual_only = [r for r in rows if r['区分'] == DETAIL_KIND_ACTUAL_ONLY][0]
        assert actual_only['AI品名'] == '' and actual_only['AI金額'] == 0 and actual_only['実績数量'] == ''

    def test_totals_add_up(self):
        rows = category_detail_rows(ITEM_DIFF)
        assert sum(r['AI金額'] for r in rows) == 50000 + 120000 + 30000 + 80000
        assert sum(r['実績金額'] for r in rows) == 50000 + 135000 + 28000 + 25000 + 15000
        assert sum(r['差額(AI−実績)'] for r in rows) == sum(r['AI金額'] for r in rows) - sum(r['実績金額'] for r in rows)

    def test_empty_and_bad_input(self):
        assert category_detail_rows(None) == []
        assert category_detail_rows({}) == []
        assert category_detail_rows({'matched': ['x', None], 'ai_only': [None], 'actual_only': 'junk'}) == []

    def test_counts_and_header(self):
        rows = category_detail_rows(ITEM_DIFF)
        counts = detail_kind_counts(rows)
        assert {k: counts[k] for k in (DETAIL_KIND_DIFF, DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY, DETAIL_KIND_SAME)} == {
            DETAIL_KIND_DIFF: 2, DETAIL_KIND_AI_ONLY: 1, DETAIL_KIND_ACTUAL_ONLY: 2, DETAIL_KIND_SAME: 1}
        assert counts['案件特有'] == 0   # 案件特有の登録が無ければ 0
        h = detail_header({'name': '電気設備工事', 'ai_total': 1243400, 'actual_total': 1246550}, counts)
        assert h.startswith('電気設備工事　AI 1,243,400円 / 実績 1,246,550円 / 差額 -3,150円')
        assert '金額が違う 2' in h and 'AIのみ 1' in h and '実績のみ 2' in h and '一致 1' in h
        assert all(v == 0 for v in detail_kind_counts([]).values())
