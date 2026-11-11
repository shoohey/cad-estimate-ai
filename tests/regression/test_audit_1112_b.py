"""2026-11-12 監査で確認された不具合の再現テスト（第2群: 学習データ・履歴・PDFの上限）

[2] 複数人で使うと「同じ商品」の登録が古い一覧で上書きされて消える
[6] 除外・案件特有の表示が、同じ品名の「一致」行にも付く
[8] 差異理由を消しても、他案件の「過去の理由（参考）」に残る
[9][10] 履歴に着手中や1回目の修正・差し戻しが出ない／順番が逆になる
[21] 24〜32MBのPDFはアプリの確認を通るのにAPIで失敗する
"""
import pytest

import ai_client
from ui_helpers import (DETAIL_KIND_AI_ONLY, DETAIL_KIND_SAME, alias_library_add, alias_library_remove,
                        alias_library_set_scope, apply_fb_status_change, category_detail_rows, diff_library_remove,
                        fb_timeline_text, make_alias_entry, release_log_rows)


def _entry(ai, ac, scope='global', pid=''):
    return make_alias_entry({'AI品名': ai, 'AI摘要': ''}, {'実績品名': ac, '実績摘要': ''}, '基礎工事', 'A邸', 'd',
                            scope=scope, project_id=pid)


class TestAliasLibraryConcurrency:
    def test_add_to_latest_keeps_others_registration(self):
        latest = [_entry('人通口補強', '深基礎')]          # 別の担当者が先に登録した
        lib, added = alias_library_add(latest, _entry('ﾀﾀﾐ', '畳'))
        assert added is True and len(lib) == 2
        assert {e['ai_name'] for e in lib} == {'人通口補強', 'ﾀﾀﾐ'}
        assert len(latest) == 1                            # 元の一覧は変えない

    def test_add_duplicate_is_noop(self):
        lib, added = alias_library_add([_entry('a', 'b')], _entry('a', 'b'))
        assert added is False and len(lib) == 1

    def test_remove_by_identity_not_position(self):
        x, y, z = _entry('x', 'X'), _entry('y', 'Y'), _entry('z', 'Z')
        # 画面に出ていた y を消す。その間に先頭へ別の登録が入っていても y だけが消える
        latest = [_entry('new', 'NEW'), x, y, z]
        out = alias_library_remove(latest, y)
        assert [e['ai_name'] for e in out] == ['new', 'x', 'z']
        assert alias_library_remove(out, y) == out          # もう無ければ何もしない

    def test_set_scope_by_identity(self):
        x = _entry('x', 'X')
        out = alias_library_set_scope([_entry('a', 'A'), x], x, 'project', 'pid-1')
        assert out[1]['scope'] == 'project' and out[1]['project_id'] == 'pid-1'
        assert out[0]['scope'] == 'global'


def test_labels_only_on_registered_kinds():
    item_diff = {'matched': [{'ai_name': '人通口補強', 'actual_name': '人通口補強', 'ai_summary': '60㎡以下',
                              'actual_summary': '60㎡以下', 'ai_amount': 137200, 'actual_amount': 137200, 'diff': 0}],
                 'ai_only': [{'name': '人通口補強', 'summary': '60㎡超過分', 'amount': 8188}],
                 'actual_only': []}
    rows = category_detail_rows(item_diff, {'人通口補強': '除外: 地域・条例の違い'})
    by_kind = {r['区分']: r for r in rows}
    assert by_kind[DETAIL_KIND_AI_ONLY]['扱い'].startswith('除外')
    assert by_kind[DETAIL_KIND_SAME]['扱い'] == '', "一致している行に除外の表示が付いた"


def test_removed_reason_disappears_from_library():
    lib = {'基礎工事|深基礎': {'item': '深基礎', 'entries': [
        {'property': 'A邸', 'reason': '高低差あり', 'category': '', 'diff': -1},
        {'property': 'B邸', 'reason': '敷地が斜面', 'category': '', 'diff': -2}]}}
    out = diff_library_remove(lib, ['基礎工事|深基礎'], 'A邸')
    assert [e['property'] for e in out['基礎工事|深基礎']['entries']] == ['B邸']
    assert [e['property'] for e in lib['基礎工事|深基礎']['entries']] == ['A邸', 'B邸']   # 元は変えない
    out2 = diff_library_remove(out, ['基礎工事|深基礎'], 'B邸')
    assert '基礎工事|深基礎' not in out2                  # 空になった項目は消える


def test_history_shows_every_step_in_order():
    fb = {'timestamp': '2026-08-04 17:25', 'body': 'x', 'status': '未対応'}
    steps = [('着手中', '試行中', '高井', '2026-10-26 10:00'), ('高井対応済み', '1回目の修正', '高井', '2026-10-26 17:00'),
             ('差し戻し', 'まだ別扱い', 'お客様', '2026-10-27 09:00'), ('着手中', '再調査', '高井', '2026-10-27 10:00'),
             ('高井対応済み', '2回目の修正', '高井', '2026-10-28 12:00'), ('対応完了', 'OK', 'お客様', '2026-10-29 09:00')]
    for st, note, by, at in steps:
        fb = apply_fb_status_change(fb, st, note, by, at)
    txt = fb_timeline_text(fb)
    expected_order = ['受付 2026-08-04 17:25', '着手中 2026-10-26 10:00（試行中）', '高井対応済み 2026-10-26 17:00（1回目の修正）',
                      '差し戻し 2026-10-27 09:00（まだ別扱い）', '着手中 2026-10-27 10:00（再調査）',
                      '高井対応済み 2026-10-28 12:00（2回目の修正）', '対応完了 2026-10-29 09:00（OK）']
    assert txt == ' → '.join(expected_order)
    rows = release_log_rows([fb])
    assert rows[0]['経緯'] == txt


class TestPdfSizeLimit:
    def test_limit_counts_base64_size(self):
        limit = ai_client.PDF_MAX_RAW_BYTES
        assert ai_client._pdf_request_bytes(limit) <= ai_client.PDF_MAX_REQUEST_BYTES
        assert ai_client._pdf_request_bytes(limit + 3) > ai_client.PDF_MAX_REQUEST_BYTES
        assert 22 * 1024 * 1024 < limit < 24 * 1024 * 1024

    def test_25mb_rejected_before_api(self, monkeypatch):
        called = []
        monkeypatch.setattr(ai_client.requests, "post", lambda *a, **k: called.append(1))
        monkeypatch.setattr(ai_client, "_get_secret", lambda k: "sk-test" if k == "ANTHROPIC_API_KEY" else "")
        big = b"%PDF-1.4\n" + b"\0" * (25 * 1024 * 1024)
        with pytest.raises(ai_client.AIClientError) as exc:
            ai_client.extract_quantities_from_pdf(big, "big.pdf")
        assert "big.pdf" in str(exc.value) and "上限" in str(exc.value)
        assert called == []
