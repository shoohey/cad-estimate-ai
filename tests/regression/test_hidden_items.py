"""図面に現れない項目: 統計・フィードバック学習・パターン登録・見積反映（純粋ロジック）"""
import copy

import pytest

from hidden_items import (SOURCE_STATS, apply_feedback, build_candidates, derive_hidden_item_stats,
                          make_key, new_pattern, pattern_from_candidates, pattern_id, record_judgments,
                          resolve_category_no, summarize_feedback, to_new_items)

# 3物件: エコキュートは 2/3（表記ゆれ: 半角カナ／全角）、仮設電気は 3/3、浄化槽は 1/3
SAMPLES = [
    {'property': '物件A', 'actual_only': [
        {'category': '給排水設備工事', 'name': 'ｴｺｷｭｰﾄ', 'summary': '460L', 'quantity': 1, 'unit': '台', 'amount': 450000},
        {'category': '電気設備工事', 'name': '仮設電気', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 30000},
    ]},
    {'property': '物件B', 'actual_only': [
        {'category': '給排水設備工事', 'name': 'エコキュート', 'summary': '370L', 'quantity': 1, 'unit': '台', 'amount': 410000},
        {'category': '電気設備工事', 'name': '仮設電気', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 36000},
        {'category': '給排水設備工事', 'name': '浄化槽', 'summary': '7人槽', 'quantity': 1, 'unit': '基', 'amount': 800000},
    ]},
    {'property': '物件C', 'actual_only': [
        {'category': '電気設備工事', 'name': '仮設電気', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 32000},
        # 同一物件内の重複（2行に分かれている）は出現1回・金額合算
        {'category': '雑工事', 'name': '産廃処分', 'summary': '', 'quantity': 2, 'unit': '回', 'amount': 40000},
        {'category': '雑工事', 'name': '産廃処分', 'summary': '', 'quantity': 1, 'unit': '回', 'amount': 20000},
    ]},
]

CATEGORIES = [(1, '仮設工事'), (14, '住宅設備機器工事'), (15, '給排水設備工事'), (16, '電気設備工事'),
              (21, '雑工事'), (23, '屋外付帯工事')]


def _by_key(stats):
    return {s['key']: s for s in stats}


class TestDeriveStats:
    def test_freq_prob_median(self):
        stats = derive_hidden_item_stats(SAMPLES)
        s = _by_key(stats)
        eco = s['給排水設備工事|エコキュート']
        assert eco['freq'] == 2 and eco['n_properties'] == 3
        assert eco['prob'] == pytest.approx(2 / 3)
        assert eco['median_amount'] == 430000           # (450000+410000)/2
        assert eco['median_quantity'] == 1.0
        assert eco['properties'] == ['物件A', '物件B']
        kari = s['電気設備工事|仮設電気']
        assert kari['freq'] == 3 and kari['prob'] == pytest.approx(1.0)
        assert kari['median_amount'] == 32000            # 30000, 32000, 36000 の中央値
        joka = s['給排水設備工事|浄化槽']
        assert joka['freq'] == 1 and joka['prob'] == pytest.approx(1 / 3) and joka['median_amount'] == 800000

    def test_name_variants_merge_into_one_key(self):
        stats = derive_hidden_item_stats(SAMPLES)
        eco_keys = [s['key'] for s in stats if 'エコキュート' in s['key']]
        assert eco_keys == ['給排水設備工事|エコキュート']
        eco = _by_key(stats)[eco_keys[0]]
        # 同数なら最初に現れた表記（半角カナ）がそのまま代表になる
        assert eco['name'] in ('ｴｺｷｭｰﾄ', 'エコキュート')
        assert eco['unit'] == '台'
        assert eco['summary'] in ('460L', '370L')

    def test_duplicate_within_property_counted_once_and_summed(self):
        s = _by_key(derive_hidden_item_stats(SAMPLES))
        dup = s['雑工事|産廃処分']
        assert dup['freq'] == 1 and dup['properties'] == ['物件C']
        assert dup['median_amount'] == 60000 and dup['median_quantity'] == 3.0

    def test_sort_order(self):
        stats = derive_hidden_item_stats(SAMPLES)
        probs = [s['prob'] for s in stats]
        assert probs == sorted(probs, reverse=True)
        # 同じ prob（1/3）の中は金額降順: 浄化槽(800000) → 産廃処分(60000)
        third = [s['key'] for s in stats if abs(s['prob'] - 1 / 3) < 1e-9]
        assert third == ['給排水設備工事|浄化槽', '雑工事|産廃処分']

    def test_empty_and_garbage(self):
        assert derive_hidden_item_stats([]) == []
        stats = derive_hidden_item_stats([{'property': 'X', 'actual_only': []},
                                          {'property': 'Y'},
                                          {'property': 'Z', 'actual_only': [{'name': '', 'amount': 1}, None, 'x']}])
        assert stats == []

    def test_does_not_mutate_input(self):
        before = copy.deepcopy(SAMPLES)
        derive_hidden_item_stats(SAMPLES)
        assert SAMPLES == before

    def test_category_differs_means_different_key(self):
        samples = [{'property': 'P', 'actual_only': [
            {'category': '電気設備工事', 'name': '配線', 'amount': 1000},
            {'category': '太陽光発電システム工事', 'name': '配線', 'amount': 2000},
        ]}]
        keys = [s['key'] for s in derive_hidden_item_stats(samples)]
        assert sorted(keys) == ['太陽光発電システム工事|配線', '電気設備工事|配線']


class TestApplyFeedback:
    def test_prob_goes_up_and_down(self):
        stats = derive_hidden_item_stats(SAMPLES)
        eco_key = '給排水設備工事|エコキュート'
        up = _by_key(apply_feedback(stats, {eco_key: {'need': 3, 'unneed': 0}}))[eco_key]
        assert up['prob_adj'] == pytest.approx(5 / 6) and up['need'] == 3 and up['unneed'] == 0
        down = _by_key(apply_feedback(stats, {eco_key: {'need': 0, 'unneed': 3}}))[eco_key]
        assert down['prob_adj'] == pytest.approx(2 / 6)
        none = _by_key(apply_feedback(stats, {}))[eco_key]
        assert none['prob_adj'] == pytest.approx(2 / 3) and none['need'] == 0 and none['unneed'] == 0

    def test_original_untouched_and_new_objects(self):
        stats = derive_hidden_item_stats(SAMPLES)
        before = copy.deepcopy(stats)
        adj = apply_feedback(stats, {'電気設備工事|仮設電気': {'need': 1, 'unneed': 1}})
        assert stats == before
        assert all('prob_adj' not in s for s in stats)
        assert adj is not stats and all(a is not s for a, s in zip(adj, stats))
        adj[0]['properties'].append('改変')
        assert stats[0]['properties'] == before[0]['properties']

    def test_boundaries(self):
        assert apply_feedback([], {'x': {'need': 1}}) == []
        zero = apply_feedback([{'key': 'k', 'freq': 0, 'n_properties': 0}], {'k': {'need': 2, 'unneed': 1}})[0]
        assert zero['prob_adj'] == pytest.approx(2 / 3)
        zero2 = apply_feedback([{'key': 'k', 'freq': 0, 'n_properties': 0}], {})[0]
        assert zero2['prob_adj'] == 0.0
        # 全物件で出現 + 必要判定でも 1.0 を超えない
        full = apply_feedback([{'key': 'k', 'freq': 3, 'n_properties': 3}], {'k': {'need': 5, 'unneed': 0}})[0]
        assert full['prob_adj'] == 1.0
        # 負数・文字列は 0 扱い
        bad = apply_feedback([{'key': 'k', 'freq': 1, 'n_properties': 2}], {'k': {'need': -4, 'unneed': 'x'}})[0]
        assert bad['need'] == 0 and bad['unneed'] == 0 and bad['prob_adj'] == pytest.approx(0.5)


class TestBuildCandidates:
    def _adj(self, feedback=None):
        return apply_feedback(derive_hidden_item_stats(SAMPLES), feedback or {})

    def test_threshold_filters(self):
        adj = self._adj()
        keys = [c['key'] for c in build_candidates(adj, 0.5, [])]
        assert keys == ['電気設備工事|仮設電気', '給排水設備工事|エコキュート']
        assert [c['source'] for c in build_candidates(adj, 0.5, [])] == [SOURCE_STATS, SOURCE_STATS]
        assert len(build_candidates(adj, 0.0, [])) == 4
        assert build_candidates(adj, 1.01, []) == []
        # しきい値ちょうどは採用
        assert '給排水設備工事|浄化槽' in [c['key'] for c in build_candidates(adj, 1 / 3, [])]

    def test_feedback_changes_adoption(self):
        adj = self._adj({'給排水設備工事|浄化槽': {'need': 4, 'unneed': 0}})
        keys = [c['key'] for c in build_candidates(adj, 0.6, [])]
        assert '給排水設備工事|浄化槽' in keys     # (1+4)/(3+4)=0.71
        adj2 = self._adj({'電気設備工事|仮設電気': {'need': 0, 'unneed': 6}})
        keys2 = [c['key'] for c in build_candidates(adj2, 0.6, [])]
        assert '電気設備工事|仮設電気' not in keys2  # 3/9=0.33

    def test_unit_price_and_cost(self):
        adj = self._adj()
        cands = {c['key']: c for c in build_candidates(adj, 0.0, [])}
        eco = cands['給排水設備工事|エコキュート']
        assert eco['quantity'] == 1.0 and eco['estimate_price'] == 430000
        assert eco['order_price'] == 288100               # 430000 × 0.67
        assert eco['prob'] == pytest.approx(2 / 3)
        dup = cands['雑工事|産廃処分']
        assert dup['quantity'] == 3.0 and dup['estimate_price'] == 20000 and dup['order_price'] == 13400
        # 数量 0 の統計は数量 1 扱い
        zero_q = build_candidates([{'key': 'c|x', 'category': 'c', 'name': 'x', 'median_amount': 1000,
                                    'median_quantity': 0, 'prob_adj': 1.0}], 0.5, [])[0]
        assert zero_q['quantity'] == 1.0 and zero_q['estimate_price'] == 1000 and zero_q['order_price'] == 670

    def test_pattern_priority_and_dedupe(self):
        adj = self._adj()
        pattern = new_pattern('オール電化', '', [
            {'category': '給排水設備工事', 'name': 'エコキュート', 'quantity': 1, 'unit': '台',
             'estimate_price': 500000, 'order_price': 300000},
            {'category': '電気設備工事', 'name': 'IHクッキングヒーター', 'quantity': 1, 'unit': '台',
             'estimate_price': 150000},
        ])
        cands = build_candidates(adj, 0.5, [pattern])
        keys = [c['key'] for c in cands]
        assert keys.count('給排水設備工事|エコキュート') == 1
        eco = cands[keys.index('給排水設備工事|エコキュート')]
        assert eco['source'] == 'パターン: オール電化' and eco['prob'] == 1.0
        assert eco['estimate_price'] == 500000 and eco['order_price'] == 300000   # 統計値ではなくパターン値
        ih = cands[keys.index('電気設備工事|ihクッキングヒーター')]
        assert ih['order_price'] == 100500 and ih['source'] == 'パターン: オール電化'
        assert '電気設備工事|仮設電気' in keys                                     # 統計分も残る
        # パターン由来が先頭
        assert keys[:2] == ['給排水設備工事|エコキュート', '電気設備工事|ihクッキングヒーター']

    def test_two_patterns_same_key_first_wins(self):
        p1 = new_pattern('A', '', [{'category': 'c', 'name': 'x', 'estimate_price': 100}])
        p2 = new_pattern('B', '', [{'category': 'c', 'name': 'x', 'estimate_price': 200}])
        cands = build_candidates([], 0.5, [p1, p2])
        assert len(cands) == 1 and cands[0]['estimate_price'] == 100 and cands[0]['source'] == 'パターン: A'

    def test_garbage_inputs(self):
        assert build_candidates([], 0.5, []) == []
        assert build_candidates([None, 'x'], 0.5, [None, {'items': None}]) == []


class TestNewPattern:
    def test_empty_name_raises(self):
        with pytest.raises(ValueError):
            new_pattern('', 'd', [])
        with pytest.raises(ValueError):
            new_pattern('   ', 'd', [])

    def test_numeric_normalization_and_missing_keys(self):
        p = new_pattern('テスト', '説明', [
            {'name': 'A', 'quantity': '2', 'estimate_price': '1500.4', 'order_price': '900'},
            {'category': '雑工事', 'name': 'B', 'estimate_price': 1000},         # order_price なし → 67%
            {'name': 'C', 'quantity': 0, 'estimate_price': None, 'order_price': 0},
            {'name': '', 'estimate_price': 5},                                  # 名称なしは捨てる
        ], at='2026-10-26 10:00')
        assert p['name'] == 'テスト' and p['description'] == '説明' and p['at'] == '2026-10-26 10:00'
        assert [i['name'] for i in p['items']] == ['A', 'B', 'C']
        a, b, c = p['items']
        assert a['quantity'] == 2.0 and isinstance(a['quantity'], float)
        assert a['estimate_price'] == 1500 and isinstance(a['estimate_price'], int) and a['order_price'] == 900
        assert a['key'] == '|a' and a['category'] == '' and a['summary'] == '' and a['unit'] == ''
        assert b['key'] == '雑工事|b' and b['order_price'] == 670
        assert c['quantity'] == 1.0 and c['estimate_price'] == 0 and c['order_price'] == 0

    def test_id_is_stable_and_differs_by_name(self):
        assert new_pattern('X', '', [])['id'] == new_pattern('X', '', [])['id']
        assert new_pattern('X', '', [])['id'] == pattern_id('X')
        assert new_pattern('X', '', [])['id'] != new_pattern('Y', '', [])['id']
        assert pattern_id('ｵｰﾙ電化') == pattern_id('オール電化')
        assert pattern_id('X').startswith('pat_') and len(pattern_id('X')) == 14

    def test_pattern_from_candidates(self):
        adj = apply_feedback(derive_hidden_item_stats(SAMPLES), {})
        cands = build_candidates(adj, 0.5, [])
        p = pattern_from_candidates('標準セット', '仮設＋給湯', cands)
        assert [i['key'] for i in p['items']] == [c['key'] for c in cands]
        assert set(p['items'][0].keys()) == {'key', 'category', 'name', 'summary', 'quantity', 'unit',
                                             'estimate_price', 'order_price'}
        assert p['items'][1]['estimate_price'] == 430000 and p['items'][1]['order_price'] == 288100


class TestToNewItems:
    def test_exact_partial_and_fallback(self):
        cands = [
            {'key': '給排水設備工事|x', 'category': '給排水設備工事', 'name': 'エコキュート', 'summary': '460L',
             'quantity': 1.0, 'unit': '台', 'estimate_price': 430000, 'order_price': 288100, 'source': '統計'},
            {'key': '給排水|y', 'category': '給排水', 'name': '散水栓', 'quantity': 2, 'unit': '箇所',
             'estimate_price': 15000, 'order_price': 10050, 'source': '統計'},
            {'key': '給排水衛生設備工事|z', 'category': '給排水衛生設備工事', 'name': '止水栓', 'quantity': 1,
             'unit': '式', 'estimate_price': 5000, 'order_price': 3350, 'source': 'パターン: 標準'},
            {'key': '外構工事|w', 'category': '外構工事', 'name': 'カーポート', 'quantity': 1, 'unit': '台',
             'estimate_price': 300000, 'order_price': 201000, 'source': '統計'},
            {'key': '|v', 'category': '', 'name': '不明項目', 'quantity': 1, 'unit': '',
             'estimate_price': 100, 'order_price': 67, 'source': '統計'},
        ]
        items = to_new_items(cands, CATEGORIES)
        # 完全一致 / 部分一致（給排水 ⊂ 給排水設備工事）/ 「給排水衛生設備工事」は包含関係にないので雑工事 /
        # 外構工事は候補に無いので雑工事 / 種別空は雑工事
        assert [i['cat_no'] for i in items] == [15, 15, 21, 21, 21]
        first = items[0]
        assert first == {'cat_no': 15, 'name': 'エコキュート', 'summary': '460L', 'quantity': 1.0, 'unit': '台',
                         'estimate_price': 430000, 'order_price': 288100,
                         'hidden_key': '給排水設備工事|x', 'hidden_source': '統計'}
        assert items[2]['hidden_source'] == 'パターン: 標準'
        assert items[4]['unit'] == '式' and items[4]['hidden_key'] == '|v'

    def test_category_matching_is_normalized(self):
        assert resolve_category_no('電気設備工事', CATEGORIES) == 16
        assert resolve_category_no(' 電気 設備 工事 ', CATEGORIES) == 16
        assert resolve_category_no('ﾃﾞﾝｷ', CATEGORIES) == 21          # 読みは一致しない
        assert resolve_category_no('電気', CATEGORIES) == 16
        assert resolve_category_no('板金・樋工事', CATEGORIES) == 21
        assert resolve_category_no('', CATEGORIES) == 21
        assert resolve_category_no('電気設備工事', []) == 21
        # 完全一致が部分一致より優先される（住宅設備機器工事 vs 設備）
        assert resolve_category_no('住宅設備機器工事', [(15, '給排水設備工事'), (14, '住宅設備機器工事')]) == 14

    def test_types_and_defaults(self):
        items = to_new_items([{'category': '雑工事', 'name': 'X', 'quantity': '3', 'estimate_price': '100.6'},
                              {'name': ''}, None], CATEGORIES)
        assert len(items) == 1
        it = items[0]
        assert it['quantity'] == 3.0 and it['estimate_price'] == 101 and it['order_price'] == 68
        assert it['hidden_key'] == '雑工事|x' and it['hidden_source'] == '' and it['summary'] == ''

    def test_does_not_mutate_candidates(self):
        cands = build_candidates(apply_feedback(derive_hidden_item_stats(SAMPLES), {}), 0.0, [])
        before = copy.deepcopy(cands)
        to_new_items(cands, CATEGORIES)
        assert cands == before


class TestRecordJudgments:
    def test_increment_and_ignore_invalid(self):
        fb = {'a': {'need': 1, 'unneed': 0}}
        out = record_judgments(fb, [
            {'key': 'a', 'judgment': '必要'},
            {'key': 'a', 'judgment': '不要'},
            {'key': 'b', 'judgment': '不要'},
            {'key': 'b', 'judgment': ''},          # 未判定
            {'key': 'b', 'judgment': 'maybe'},     # 不正値
            {'key': '', 'judgment': '必要'},       # キーなし
            {'judgment': '必要'}, None, 'x',
        ])
        assert out == {'a': {'need': 2, 'unneed': 1}, 'b': {'need': 0, 'unneed': 1}}

    def test_non_destructive(self):
        fb = {'a': {'need': 1, 'unneed': 0}}
        before = copy.deepcopy(fb)
        out = record_judgments(fb, [{'key': 'a', 'judgment': '必要'}])
        assert fb == before and out is not fb and out['a'] is not fb['a']

    def test_empty_and_garbage_feedback(self):
        assert record_judgments({}, []) == {}
        assert record_judgments(None, [{'key': 'k', 'judgment': '不要'}]) == {'k': {'need': 0, 'unneed': 1}}
        assert record_judgments({'k': 'broken', 'j': {'need': -2, 'unneed': 'x'}}, []) == {
            'k': {'need': 0, 'unneed': 0}, 'j': {'need': 0, 'unneed': 0}}


class TestSummarizeFeedback:
    def test_counts(self):
        adj = apply_feedback(derive_hidden_item_stats(SAMPLES),
                             {'電気設備工事|仮設電気': {'need': 3, 'unneed': 1},
                              '給排水設備工事|浄化槽': {'need': 0, 'unneed': 2}})
        s = summarize_feedback(adj)
        assert s['n_items'] == 4 and s['n_need'] == 3 and s['n_unneed'] == 3
        assert s['need_rate'] == pytest.approx(0.5)

    def test_empty(self):
        assert summarize_feedback([]) == {'n_items': 0, 'n_need': 0, 'n_unneed': 0, 'need_rate': 0.0}
        assert summarize_feedback(None)['n_items'] == 0


class TestRoundTrip:
    def test_feedback_loop_end_to_end(self):
        """統計 → 候補 → 見積反映 → 判定 → 再計算 の一巡で、不要判定が候補から外れること。"""
        stats = derive_hidden_item_stats(SAMPLES)
        fb = {}
        cands = build_candidates(apply_feedback(stats, fb), 0.6, [])
        assert [c['key'] for c in cands] == ['電気設備工事|仮設電気', '給排水設備工事|エコキュート']
        items = to_new_items(cands, CATEGORIES)
        assert [i['cat_no'] for i in items] == [16, 15]
        # 施主が「エコキュート不要」を3回続けて判定
        for _ in range(3):
            fb = record_judgments(fb, [{'key': i['hidden_key'], 'judgment': '不要'} for i in items if '給排水' in i['hidden_key']])
        cands2 = build_candidates(apply_feedback(stats, fb), 0.6, [])
        assert [c['key'] for c in cands2] == ['電気設備工事|仮設電気']
        assert make_key('給排水設備工事', 'ｴｺｷｭｰﾄ') == '給排水設備工事|エコキュート'
