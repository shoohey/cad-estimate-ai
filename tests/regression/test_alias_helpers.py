"""「同じ商品」登録（名寄せ辞書）と「この案件特有」登録の純粋ロジック"""
import excel_diff
from ui_helpers import (ALIAS_SCOPE_GLOBAL, ALIAS_SCOPE_PROJECT, DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY,
                        DETAIL_KIND_DIFF, DETAIL_KIND_SAME, PROJECT_SPECIFIC_LABEL, alias_applies,
                        alias_entry_exists, alias_pairs_from_library, category_detail_rows, detail_header,
                        detail_kind_counts, make_alias_entry, special_note_name, validate_alias_selection,
                        validate_special_selection)

AI_ROW = {'区分': DETAIL_KIND_AI_ONLY, 'AI品名': 'SGL鋼板 たてひら葺', 'AI摘要': 'ﾆｽｸｶﾗｰPro ※唐草･棟含む',
          'AI金額': 681671, '実績品名': '', '実績摘要': '', '実績金額': 0}
AC_ROW = {'区分': DETAIL_KIND_ACTUAL_ONLY, 'AI品名': '', 'AI摘要': '', 'AI金額': 0,
          '実績品名': '平板瓦葺（切妻/寄棟）', '実績摘要': 'ｾﾗﾑFF程度　標準勾配', '実績金額': 691272}


class TestAliasEntry:
    def test_entry_keys_and_scope(self):
        e = make_alias_entry(AI_ROW, AC_ROW, '屋根工事', 'サンプル邸', '2026-10-26 18:00',
                             scope=ALIAS_SCOPE_PROJECT, project_id='p1')
        assert e['ai_name'] == 'SGL鋼板 たてひら葺' and e['actual_name'] == '平板瓦葺（切妻/寄棟）'
        assert 'sgl鋼板たてひら葺' in e['ai_keys'] and len(e['ai_keys']) == 2   # 名称のみ／名称＋摘要
        assert e['scope'] == 'project' and e['project_id'] == 'p1'
        assert make_alias_entry(AI_ROW, AC_ROW, 'c', 'p', 'd', scope='junk')['scope'] == ALIAS_SCOPE_GLOBAL

    def test_pairs_respect_scope(self):
        g = make_alias_entry(AI_ROW, AC_ROW, '屋根工事', 'A邸', 'd', scope=ALIAS_SCOPE_GLOBAL)
        p_other = make_alias_entry({'AI品名': 'X', 'AI摘要': ''}, {'実績品名': 'Y', '実績摘要': ''}, 'c', 'B邸', 'd',
                                   scope=ALIAS_SCOPE_PROJECT, project_id='other')
        p_mine = make_alias_entry({'AI品名': 'P', 'AI摘要': ''}, {'実績品名': 'Q', '実績摘要': ''}, 'c', 'A邸', 'd',
                                  scope=ALIAS_SCOPE_PROJECT, project_id='mine')
        lib = [g, p_other, p_mine, 'junk', None]
        pairs = alias_pairs_from_library(lib, 'mine')
        assert frozenset(('x', 'y')) not in pairs and frozenset(('p', 'q')) in pairs
        assert frozenset(('sgl鋼板たてひら葺', '平板瓦葺切妻寄棟')) in pairs
        assert alias_applies(p_other, 'mine') is False and alias_applies(p_other, 'other') is True
        assert alias_applies(g, None) is True and alias_applies(p_mine, None) is False
        assert alias_pairs_from_library(None) == set()

    def test_exists(self):
        e = make_alias_entry(AI_ROW, AC_ROW, '屋根工事', 'A邸', 'd')
        assert alias_entry_exists([e], make_alias_entry(AI_ROW, AC_ROW, '屋根工事', 'B邸', 'x'))
        assert not alias_entry_exists([e], make_alias_entry({'AI品名': 'Z'}, AC_ROW, '屋根工事', 'B邸', 'x'))
        assert not alias_entry_exists(None, e)

    def test_validate_selection(self):
        err, a, b = validate_alias_selection([AI_ROW, AC_ROW])
        assert err is None and a is AI_ROW and b is AC_ROW
        assert validate_alias_selection([AI_ROW])[0]
        assert validate_alias_selection([AI_ROW, AI_ROW, AC_ROW])[0]
        assert validate_alias_selection([AI_ROW, AC_ROW, {'区分': DETAIL_KIND_DIFF}])[0]
        assert validate_alias_selection([])[0]


class TestAliasMatching:
    def test_alias_forces_match_even_when_names_differ(self):
        ai = [{'name': 'SGL鋼板 たてひら葺', 'summary': 'ﾆｽｸｶﾗｰPro', 'quantity': 1, 'unit': '式', 'amount': 681671}]
        ac = [{'name': '平板瓦葺（切妻/寄棟）', 'summary': 'ｾﾗﾑFF程度', 'quantity': 1, 'unit': '式', 'amount': 691272}]
        base = excel_diff._match_items(ai, ac)
        assert not base['matched'] and base['ai_only'] and base['actual_only']
        pairs = alias_pairs_from_library([make_alias_entry(AI_ROW, AC_ROW, '屋根工事', 'A邸', 'd')])
        out = excel_diff._match_items(ai, ac, pairs)
        assert len(out['matched']) == 1 and out['matched'][0]['via_alias'] is True
        assert out['matched'][0]['diff'] == 681671 - 691272 and not out['ai_only'] and not out['actual_only']

    def test_alias_beats_fuzzy_candidate(self):
        ai = [{'name': '鋼板屋根', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 100}]
        ac = [{'name': '鋼板屋根 荷揚げ費', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 90},
              {'name': 'ガルバ屋根本体', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 200}]
        pairs = {frozenset(('鋼板屋根', 'ガルバ屋根本体'))}
        out = excel_diff._match_items(ai, ac, pairs)
        assert out['matched'][0]['actual_name'] == 'ガルバ屋根本体' and out['matched'][0]['via_alias']

    def test_no_alias_keeps_behaviour(self):
        ai = [{'name': 'システムバス', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 1}]
        ac = [{'name': 'システムバス', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 1}]
        out = excel_diff._match_items(ai, ac, set())
        assert out['matched'] and out['matched'][0]['via_alias'] is False


class TestProjectSpecific:
    def test_special_names_mark_rows_and_counts(self):
        item_diff = {'matched': [], 'ai_only': [{'name': '人通口補強', 'amount': 8188}],
                     'actual_only': [{'name': '深基礎', 'summary': '高低差450㎜以下', 'amount': 188485},
                                     {'name': '布基礎', 'amount': 10396}]}
        rows = category_detail_rows(item_diff, {'深基礎'})
        by_name = {r['実績品名'] or r['AI品名']: r for r in rows}
        assert by_name['深基礎']['扱い'] == PROJECT_SPECIFIC_LABEL
        assert by_name['布基礎']['扱い'] == '' and by_name['人通口補強']['扱い'] == ''
        counts = detail_kind_counts(rows)
        assert counts[PROJECT_SPECIFIC_LABEL] == 1 and counts[DETAIL_KIND_ACTUAL_ONLY] == 2
        assert '案件特有 1' in detail_header({'name': '基礎工事', 'ai_total': 1, 'actual_total': 2}, counts)
        assert '案件特有' not in detail_header({'name': 'x'}, detail_kind_counts(category_detail_rows(item_diff)))

    def test_validate_special_selection(self):
        assert validate_special_selection([])[0]
        assert validate_special_selection([{'区分': DETAIL_KIND_SAME}])[0]
        err, rows = validate_special_selection([AC_ROW, AI_ROW])
        assert err is None and len(rows) == 2
        assert special_note_name(AC_ROW) == '平板瓦葺（切妻/寄棟）' and special_note_name(AI_ROW) == 'SGL鋼板 たてひら葺'
