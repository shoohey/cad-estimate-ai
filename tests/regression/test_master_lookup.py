"""設計部の単価マスタ取込・類似候補・対応表取込（master_lookup）の回帰テスト

顧客実名は書かない。アプリ同梱の tankamaster_updated.csv 以外の入力は一時ファイルで生成する。
"""
import csv
import json
import os

import openpyxl
import pytest

import master_lookup as ml
from ui_helpers import ALIAS_SCOPE_GLOBAL, ALIAS_SCOPE_PROJECT, alias_entry_exists, make_alias_entry

ROOT = os.path.dirname(os.path.abspath(ml.__file__))
MASTER_CSV = os.path.join(ROOT, 'tankamaster_updated.csv')


def _xlsx_bytes(tmp_path, sheets: dict, name='master.xlsx') -> bytes:
    """{シート名: [[行], ...]} を一時 xlsx に保存してバイト列で返す（シートは辞書の順）"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    path = tmp_path / name
    wb.save(path)
    return path.read_bytes()


def _csv_bytes(tmp_path, text: str, encoding='utf-8', name='master.csv') -> bytes:
    path = tmp_path / name
    path.write_bytes(text.encode(encoding))
    return path.read_bytes()


# ================================================================ 1. 単価マスタの取込
class TestParsePriceMasterBundled:
    @pytest.fixture(scope='class')
    def parsed(self):
        with open(MASTER_CSV, 'rb') as f:
            return ml.parse_price_master(f.read(), os.path.basename(MASTER_CSV))

    def test_row_count_matches_item_lines(self, parsed):
        rows, errors = parsed
        with open(MASTER_CSV, encoding='utf-8') as f:
            raw = list(csv.DictReader(f))
        expected = sum(1 for r in raw if int(r['工事コード']) >= ml.CATEGORY_CODE_MAX)
        assert errors == []
        assert len(rows) == expected and 250 < len(rows) < 320      # 種別見出し行（1〜24）は含めない
        assert '仮設工事' not in {r['name'] for r in rows}

    def test_every_row_has_category_and_int_prices(self, parsed):
        rows, _ = parsed
        assert all(r['category'] for r in rows)
        assert all(isinstance(r['estimate_price'], int) and isinstance(r['order_price'], int) for r in rows)
        assert all(isinstance(r['row_no'], int) for r in rows)
        assert [r['row_no'] for r in rows] == sorted(r['row_no'] for r in rows)
        assert set(rows[0].keys()) == {'code', 'category', 'name', 'summary', 'unit', 'order_price',
                                       'estimate_price', 'row_no'}
        json.dumps(rows, ensure_ascii=False)                          # クラウド保存できる形

    def test_categories_follow_code_prefix(self, parsed):
        rows, _ = parsed
        by_code = {r['code']: r for r in rows}
        assert by_code['10001']['category'] == '仮設工事' and by_code['10001']['estimate_price'] == 900
        assert by_code['10001']['unit'] == '㎡' and by_code['10001']['order_price'] == 600
        assert by_code['20002']['category'] == '基礎工事'
        assert by_code['30006']['estimate_price'] == -85800           # マイナス単価もそのまま
        # ファイル末尾に追記された行も接頭辞から正しい種別になる（直前の見出し「諸費用」に引きずられない）
        assert by_code['80035']['category'] == '内部建具工事'
        assert by_code['30099']['category'] == '木工事'
        assert by_code['160030']['category'] == '電気設備工事'
        # 22 管理費は率計算で明細を持たない → 220001〜は屋外付帯工事、230001〜は諸費用（estimate_calculator と同じ）
        assert by_code['220001']['category'] == '屋外付帯工事'
        assert by_code['230001']['category'] == '諸費用'
        assert len(ml.master_categories(rows)) == 23


class TestParsePriceMasterFiles:
    def test_xlsx_header_on_third_row_with_aliases_and_sheet_choice(self, tmp_path):
        data = _xlsx_bytes(tmp_path, {
            'メモ': [['このブックは設計部の単価一覧です'], ['更新日', '2026-10-01']],
            'マスタ': [
                ['設計部 単価マスタ'],
                [],
                ['工種', '品名', '規格・仕様', '単位', '売単価(円)', '仕入単価'],
                ['1', '仮設工事', None, None, None, None],
                ['10001', '外部足場組', '養生ﾈｯﾄ等含む', '㎡', '1,200円', 600],
                [None, '仮設ﾄｲﾚ', None, '式', 35800.0, 25000.0],
                [2, '基礎工事', None, '式', 0, 0],
                [20002, 'ﾍﾞﾀ基礎', '', '㎡', 28200, 19700],
                [20099, '要打合せ項目', '', '式', '要相談', None],
                [None, None, None, None, None, None],
            ],
        })
        rows, errors = ml.parse_price_master(data, 'design_master.xlsx')
        assert [r['name'] for r in rows] == ['外部足場組', '仮設ﾄｲﾚ', 'ﾍﾞﾀ基礎', '要打合せ項目']
        assert [r['category'] for r in rows] == ['仮設工事', '仮設工事', '基礎工事', '基礎工事']
        assert rows[0] == {'code': '10001', 'category': '仮設工事', 'name': '外部足場組', 'summary': '養生ﾈｯﾄ等含む',
                           'unit': '㎡', 'order_price': 600, 'estimate_price': 1200, 'row_no': 5}
        assert rows[1]['estimate_price'] == 35800 and rows[1]['order_price'] == 25000 and rows[1]['code'] == ''
        assert rows[2]['code'] == '20002' and rows[2]['row_no'] == 8
        assert rows[3]['estimate_price'] == 0
        assert len(errors) == 1 and errors[0].startswith('9行目') and '要相談' in errors[0]

    def test_cp932_csv_with_halfwidth_kana(self, tmp_path):
        text = ('工事コード,名称,摘要,単位,発注単価,見積単価\n'
                '14,住宅設備機器工事,,式,0,0\n'
                '140007,ｼｽﾃﾑﾊﾞｽ 1616,TOTO『ｻｻﾞﾅ』,ｾｯﾄ,260000,371500\n'
                '140001,ｷｯﾁﾝｾｯﾄ　壁付型,標準,ｾｯﾄ,360000,514300\n')
        rows, errors = ml.parse_price_master(_csv_bytes(tmp_path, text, 'cp932'), 'master_sjis.csv')
        assert errors == []
        assert [r['name'] for r in rows] == ['ｼｽﾃﾑﾊﾞｽ 1616', 'ｷｯﾁﾝｾｯﾄ　壁付型']
        assert rows[0]['summary'] == 'TOTO『ｻｻﾞﾅ』' and rows[0]['category'] == '住宅設備機器工事'
        assert rows[0]['estimate_price'] == 371500 and rows[0]['unit'] == 'ｾｯﾄ'

    def test_utf8_sig_csv_first_column_is_code(self, tmp_path):
        text = '工事コード,名称,摘要,単位,発注単価,見積単価\n1,仮設工事,,式,0,0\n10001,外部足場組,,㎡,600,900\n'
        rows, errors = ml.parse_price_master(_csv_bytes(tmp_path, text, 'utf-8-sig'), 'bom.csv')
        assert errors == [] and len(rows) == 1
        assert rows[0]['code'] == '10001' and rows[0]['category'] == '仮設工事'

    def test_explicit_category_column_fills_down(self, tmp_path):
        text = ('工事種別,項目,仕様,単価\n'
                '仮設工事,外部足場,養生ﾈｯﾄ含む,900\n'
                ',仮設トイレ,,35800\n'
                '基礎工事,ベタ基礎,,28200\n'
                '基礎工事,基礎工事一式,,0\n')
        rows, errors = ml.parse_price_master(_csv_bytes(tmp_path, text), 'm.csv')
        assert errors == []
        assert [(r['name'], r['category']) for r in rows] == [
            ('外部足場', '仮設工事'), ('仮設トイレ', '仮設工事'), ('ベタ基礎', '基礎工事'), ('基礎工事一式', '基礎工事')]
        assert rows[0]['estimate_price'] == 900 and rows[0]['order_price'] == 0 and rows[0]['summary'] == '養生ﾈｯﾄ含む'

    def test_blank_name_row_reported_but_blank_line_silent(self, tmp_path):
        text = '名称,単価\n外部足場,900\n,\n,500\n仮設トイレ,35800\n'
        rows, errors = ml.parse_price_master(_csv_bytes(tmp_path, text), 'm.csv')
        assert [r['name'] for r in rows] == ['外部足場', '仮設トイレ']
        assert len(errors) == 1 and errors[0].startswith('4行目')

    def test_empty_file(self):
        rows, errors = ml.parse_price_master(b'', 'master.csv')
        assert rows == [] and errors and '空' in errors[0]

    def test_unsupported_extension(self):
        rows, errors = ml.parse_price_master(b'abc', 'master.pdf')
        assert rows == [] and errors and 'ファイル形式' in errors[0]

    def test_broken_xlsx(self):
        rows, errors = ml.parse_price_master(b'this is not a zip', 'master.xlsx')
        assert rows == [] and errors and 'Excel' in errors[0]

    def test_header_not_found(self, tmp_path):
        rows, errors = ml.parse_price_master(_csv_bytes(tmp_path, 'a,b,c\n1,2,3\n'), 'm.csv')
        assert rows == [] and errors and '見出し' in errors[0]

    def test_master_categories_order_and_junk(self):
        rows = [{'category': '木工事'}, {'category': ''}, 'junk', None, {'category': '仮設工事'}, {'category': '木工事'}]
        assert ml.master_categories(rows) == ['木工事', '仮設工事']
        assert ml.master_categories(None) == []


# ================================================================ 2. 類似候補
def _m(code, category, name, summary='', unit='式', est=1000):
    return {'code': code, 'category': category, 'name': name, 'summary': summary, 'unit': unit,
            'order_price': int(est * 0.7), 'estimate_price': est, 'row_no': int(code)}


MASTER = [
    _m('140007', '住宅設備機器工事', 'ｼｽﾃﾑﾊﾞｽ', '1616 LIXIL ﾘﾃﾞｱ', 'ｾｯﾄ', 371500),
    _m('140001', '住宅設備機器工事', 'ｼｽﾃﾑｷｯﾁﾝ', '壁付型 I-255', 'ｾｯﾄ', 514300),
    _m('50003', '屋根工事', '鋼板屋根', '荷揚げ費', '式', 35800),
    _m('50001', '屋根工事', '小屋裏換気(鋼板屋根用)', '', '㎡', 1200),
    _m('10001', '仮設工事', '外部足場組', '養生ﾈｯﾄ等含む', '㎡', 900),
    _m('10011', '仮設工事', '養生費', '', '式', 20000),
    _m('30044', '木工事', '養生費', '', '式', 15000),
    _m('30001', '木工事', '1階床 無垢ﾌﾛｰﾘﾝｸﾞ（標準）', '北恵 ※施工費込', '坪', 31200),
    _m('130001', '内装工事', 'ｸﾛｽ工事(一般部)', '壁･天井 SP級(無地)', '㎡', 1300),
]
RESULT_KEYS = {'name', 'summary', 'category', 'unit', 'estimate_price', 'score', 'reason'}


class TestSuggestSimilar:
    def test_halfwidth_and_fullwidth_kana_match(self):
        res = ml.suggest_similar('システムバス', '', MASTER)
        assert res and res[0]['name'] == 'ｼｽﾃﾑﾊﾞｽ' and res[0]['score'] >= 0.8
        assert res[0]['reason'].startswith('名称が一致')
        assert res[0]['estimate_price'] == 371500 and res[0]['unit'] == 'ｾｯﾄ'
        assert all(r['name'] != 'ｼｽﾃﾑｷｯﾁﾝ' or r['score'] < res[0]['score'] for r in res)
        # 逆方向（実績が半角カナ、マスタが全角）も同じ
        rev = ml.suggest_similar('ｼｽﾃﾑﾊﾞｽ', '', [_m('1', 'x', 'システムバス')])
        assert rev and rev[0]['score'] == 1.0

    def test_name_plus_summary_match(self):
        res = ml.suggest_similar('鋼板屋根 荷揚げ費', '', MASTER)
        assert res[0]['name'] == '鋼板屋根' and res[0]['summary'] == '荷揚げ費'
        assert res[0]['score'] == 1.0 and res[0]['reason'].startswith('名称＋摘要が一致')
        others = [r for r in res if r['name'] == '小屋裏換気(鋼板屋根用)']
        assert all(r['score'] < 0.6 for r in others)
        # 実績側が名称と摘要に分けて書いても同じ
        res2 = ml.suggest_similar('鋼板屋根', '荷揚げ費', MASTER, category='屋根工事')
        assert res2[0]['summary'] == '荷揚げ費' and res2[0]['score'] == 1.0 and '同じ工事種別' in res2[0]['reason']

    def test_summary_holding_the_product_name(self):
        res = ml.suggest_similar('1階床材変更', '1階床 無垢ﾌﾛｰﾘﾝｸﾞ', MASTER)
        assert res and res[0]['name'].startswith('1階床 無垢')

    def test_unrelated_word_gives_nothing(self):
        assert ml.suggest_similar('ペンキ塗装一式', '', MASTER) == []
        assert ml.suggest_similar('ガス工事', 'プロパン', MASTER) == []   # 「工事」が共通なだけ
        assert ml.suggest_similar('左官工事', '', MASTER) == []

    def test_category_preference_keeps_other_categories(self):
        res = ml.suggest_similar('養生費', '', MASTER, category='木工事')
        assert [r['category'] for r in res[:2]] == ['木工事', '仮設工事']
        assert res[0]['score'] == 1.0 and res[1]['score'] == 1.0
        assert '同じ工事種別' in res[0]['reason'] and '同じ工事種別' not in res[1]['reason']
        plain = ml.suggest_similar('養生費', '', MASTER)
        assert [r['category'] for r in plain[:2]] == ['仮設工事', '木工事']    # 指定なしはマスタの並び順
        # 工事種別名の表記ゆれ（半角の中黒・末尾の「工事」抜け）も同じ種別とみなす
        fuzzy = ml.suggest_similar('養生費', '', MASTER, category='木工事 ')
        assert fuzzy[0]['category'] == '木工事'

    def test_category_bonus_is_capped_and_bounded(self):
        for cat in (None, '仮設工事', '存在しない種別'):
            res = ml.suggest_similar('養生', '', MASTER, category=cat)
            assert all(0.0 <= r['score'] <= 1.0 for r in res)
            assert [r['score'] for r in res] == sorted((r['score'] for r in res), reverse=True)

    def test_top_n(self):
        assert len(ml.suggest_similar('養生費', '', MASTER, top_n=1)) == 1
        assert len(ml.suggest_similar('養生費', '', MASTER, top_n=5)) >= 2
        assert ml.suggest_similar('養生費', '', MASTER, top_n=0) == []

    def test_empty_inputs_and_junk_rows(self):
        assert ml.suggest_similar('', '', MASTER) == []
        assert ml.suggest_similar(None, None, MASTER) == []
        assert ml.suggest_similar('養生費', '', []) == []
        assert ml.suggest_similar('養生費', '', None) == []
        junk = MASTER + ['junk', None, {'name': ''}, {'summary': 'x'}, {'name': '養生費', 'estimate_price': '1,500円'}]
        res = ml.suggest_similar('養生費', '', junk)
        assert res and all(set(r.keys()) == RESULT_KEYS for r in res)
        assert any(r['estimate_price'] == 1500 and r['category'] == '' for r in res)

    def test_duplicate_master_rows_collapse(self):
        res = ml.suggest_similar('養生費', '', MASTER + [MASTER[5], MASTER[5]])
        assert [r['category'] for r in res] == ['仮設工事', '木工事']

    def test_containment(self):
        res = ml.suggest_similar('外部足場', '', MASTER)
        assert res[0]['name'] == '外部足場組' and res[0]['reason'].startswith('名称を含む')


# ================================================================ 3. 対応表の取込・辞書マージ
TEMPLATE_HEADER = '図面表記,AI摘要,実績品名,実績摘要,工事種別,適用範囲\n'


class TestParseAliasTable:
    def test_template_csv(self, tmp_path):
        text = (TEMPLATE_HEADER +
                'カウンター,,造作カウンター,,木工事,全案件共通\n'
                '可動棚,,可動棚 ﾌｨｯﾄﾗｯｸ,材工,内部建具工事,この案件のみ\n')
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, text, 'utf-8-sig', 'taiou.csv'), 'taiou.csv',
                                               'サンプル邸', 'p1', '2026-10-26 10:00')
        assert errors == [] and len(entries) == 2
        e0, e1 = entries
        assert e0['ai_name'] == 'カウンター' and e0['actual_name'] == '造作カウンター' and e0['category'] == '木工事'
        assert e0['scope'] == ALIAS_SCOPE_GLOBAL and e0['property'] == 'サンプル邸' and e0['at'] == '2026-10-26 10:00'
        assert e1['scope'] == ALIAS_SCOPE_PROJECT and e1['project_id'] == 'p1'
        assert e1['actual_summary'] == '材工' and e1['ai_summary'] == ''
        assert e1 == make_alias_entry({'AI品名': '可動棚', 'AI摘要': ''}, {'実績品名': '可動棚 ﾌｨｯﾄﾗｯｸ', '実績摘要': '材工'},
                                      '内部建具工事', 'サンプル邸', '2026-10-26 10:00',
                                      scope=ALIAS_SCOPE_PROJECT, project_id='p1')
        json.dumps(entries, ensure_ascii=False)

    def test_column_aliases_without_scope(self, tmp_path):
        text = 'AI品名,マスタ名,工種\nUB,ｼｽﾃﾑﾊﾞｽ,住宅設備機器工事\n'
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, text), 'x.csv', 'A邸', '', 'd')
        assert errors == [] and len(entries) == 1
        assert entries[0]['ai_name'] == 'UB' and entries[0]['actual_name'] == 'ｼｽﾃﾑﾊﾞｽ'
        assert entries[0]['category'] == '住宅設備機器工事' and entries[0]['scope'] == ALIAS_SCOPE_GLOBAL

    def test_xlsx_header_on_second_row(self, tmp_path):
        data = _xlsx_bytes(tmp_path, {'対応表': [
            ['図面表記と実績名称の対応表'],
            ['図面の表記', 'AI 摘要', '実績見積の名称', '実績 摘要', '種別', '適用'],
            ['UB', '1616', 'ｼｽﾃﾑﾊﾞｽ', 'LIXIL', '住宅設備機器工事', 'この案件のみ'],
            [None, None, None, None, None, None],
        ]}, 'taiou.xlsx')
        entries, errors = ml.parse_alias_table(data, 'taiou.xlsx', 'A邸', 'p9', 'd')
        assert errors == [] and len(entries) == 1
        e = entries[0]
        assert (e['ai_name'], e['ai_summary'], e['actual_name'], e['actual_summary']) == ('UB', '1616', 'ｼｽﾃﾑﾊﾞｽ', 'LIXIL')
        assert e['scope'] == ALIAS_SCOPE_PROJECT and e['project_id'] == 'p9' and e['category'] == '住宅設備機器工事'

    def test_scope_values(self, tmp_path):
        text = (TEMPLATE_HEADER +
                'A1,,B1,,,\n'
                'A2,,B2,,,全案件共通\n'
                'A3,,B3,,,この案件のみ\n'
                'A4,,B4,,,案件特有\n'
                'A5,,B5,,,ばらばら\n')
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, text), 'x.csv', 'A邸', 'p1', 'd')
        assert [e['scope'] for e in entries] == [ALIAS_SCOPE_GLOBAL, ALIAS_SCOPE_GLOBAL,
                                                 ALIAS_SCOPE_PROJECT, ALIAS_SCOPE_PROJECT]
        assert len(errors) == 1 and errors[0].startswith('6行目') and '適用範囲' in errors[0] and 'ばらばら' in errors[0]

    def test_project_scope_needs_project_id(self, tmp_path):
        text = TEMPLATE_HEADER + 'A,,B,,,この案件のみ\nC,,D,,,\n'
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, text), 'x.csv', '', '', 'd')
        assert [e['ai_name'] for e in entries] == ['C']
        assert len(errors) == 1 and errors[0].startswith('2行目') and '案件' in errors[0]

    def test_missing_side_skipped_with_reason_blank_rows_silent(self, tmp_path):
        text = (TEMPLATE_HEADER +
                '片側だけ,,,,,\n'
                ',,,,,\n'
                ',,実績だけ,,,\n'
                'おなじ,,おなじ,,,\n'
                'OK,,ok側,,,\n')
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, text), 'x.csv', 'A邸', 'p1', 'd')
        assert [e['ai_name'] for e in entries] == ['OK']
        assert len(errors) == 3
        assert errors[0].startswith('2行目') and '実績品名' in errors[0]
        assert errors[1].startswith('4行目') and '図面表記' in errors[1]
        assert errors[2].startswith('5行目') and '同じ表記' in errors[2]

    def test_in_file_duplicates_collapse(self, tmp_path):
        text = TEMPLATE_HEADER + 'UB,,ｼｽﾃﾑﾊﾞｽ,,,\nUB,,ｼｽﾃﾑﾊﾞｽ,,住宅設備機器工事,共通\nUB,,システムバス,,,\n'
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, text), 'x.csv', 'A邸', 'p1', 'd')
        assert len(entries) == 1                     # 2行目は同一、3行目も半角/全角が正規化後に同じキー → 重複
        assert entries[0]['actual_name'] == 'ｼｽﾃﾑﾊﾞｽ' and entries[0]['category'] == ''   # 先に出た行を採用
        assert len(errors) == 2 and all('先にある' in e for e in errors)
        assert errors[0].startswith('3行目') and errors[1].startswith('4行目')

    def test_file_errors(self, tmp_path):
        assert ml.parse_alias_table(b'', 'x.csv', 'A', 'p', 'd') == ([], ['ファイルが空です'])
        entries, errors = ml.parse_alias_table(b'x', 'x.docx', 'A', 'p', 'd')
        assert entries == [] and 'ファイル形式' in errors[0]
        entries, errors = ml.parse_alias_table(_csv_bytes(tmp_path, '名称,単価\nA,1\n'), 'x.csv', 'A', 'p', 'd')
        assert entries == [] and '見出し' in errors[0]

    def test_parse_scope_label(self):
        assert ml.parse_scope_label('') == ALIAS_SCOPE_GLOBAL
        assert ml.parse_scope_label(None) == ALIAS_SCOPE_GLOBAL
        assert ml.parse_scope_label('全案件共通') == ALIAS_SCOPE_GLOBAL
        assert ml.parse_scope_label('共通') == ALIAS_SCOPE_GLOBAL
        assert ml.parse_scope_label('この案件のみ') == ALIAS_SCOPE_PROJECT
        assert ml.parse_scope_label('この案件だけ') == ALIAS_SCOPE_PROJECT
        assert ml.parse_scope_label('project') == ALIAS_SCOPE_PROJECT
        assert ml.parse_scope_label('不明') is None


class TestMergeAliasEntries:
    def _entries(self):
        return [
            make_alias_entry({'AI品名': 'UB', 'AI摘要': ''}, {'実績品名': 'ｼｽﾃﾑﾊﾞｽ', '実績摘要': ''}, 'c', 'A邸', 'd'),
            make_alias_entry({'AI品名': 'カウンター', 'AI摘要': ''}, {'実績品名': '造作カウンター', '実績摘要': ''}, 'c', 'A邸', 'd'),
            make_alias_entry({'AI品名': '可動棚', 'AI摘要': ''}, {'実績品名': '可動棚 ﾌｨｯﾄﾗｯｸ', '実績摘要': '材工'}, 'c', 'A邸', 'd'),
        ]

    def test_adds_only_new_entries_without_mutating(self):
        entries = self._entries()
        existing = make_alias_entry({'AI品名': 'UB', 'AI摘要': ''}, {'実績品名': 'システムバス', '実績摘要': ''},
                                    'other', 'B邸', 'x', scope=ALIAS_SCOPE_PROJECT, project_id='p2')
        lib = [existing, 'junk']
        new_lib, added = ml.merge_alias_entries(lib, entries)
        assert added == 2 and len(new_lib) == 4
        assert lib == [existing, 'junk']                          # 元の辞書は変えない
        assert new_lib[0] is existing                             # 既存の scope は上書きしない
        assert all(alias_entry_exists(new_lib, e) for e in entries)

    def test_duplicates_inside_entries_and_empty_inputs(self):
        entries = self._entries()
        new_lib, added = ml.merge_alias_entries(None, entries + [dict(entries[0]), 'junk', None])
        assert added == 3 and len(new_lib) == 3
        assert ml.merge_alias_entries([entries[0]], None) == ([entries[0]], 0)
        assert ml.merge_alias_entries(None, None) == ([], 0)
