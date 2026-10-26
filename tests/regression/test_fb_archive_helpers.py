"""問い合わせの保持・ステータス管理・両者確認後のアーカイブ（純粋ロジック）"""
from ui_helpers import (fb_is_archived, fb_timeline, fb_timeline_text, partition_feedbacks,
                        release_log_rows)


def _fb(status, history=None, **kw):
    d = {'timestamp': '2026-08-04 17:25', 'body': '全選択が欲しい', 'category': '操作性', 'status': status}
    if history is not None:
        d['status_history'] = history
    d.update(kw)
    return d


HIST = [
    {'status': '高井対応済み', 'note': '全て選択ボタンを追加', 'by': '高井', 'at': '2026-10-26 17:00'},
    {'status': '対応完了', 'note': '確認しました', 'by': 'お客様', 'at': '2026-10-27 09:30'},
]


class TestArchiveRule:
    def test_done_is_archived_others_active(self):
        assert fb_is_archived(_fb('対応完了'))
        assert fb_is_archived(_fb('対応済み'))          # 旧2値データも完了扱い
        for s in ('未対応', '高井対応済み', '差し戻し', None, ''):
            assert not fb_is_archived(_fb(s))

    def test_partition_keeps_order_and_everything(self):
        fbs = [_fb('未対応', body='a'), _fb('対応完了', body='b'), _fb('差し戻し', body='c'), _fb('対応済み', body='d')]
        active, archived = partition_feedbacks(fbs)
        assert [f['body'] for f in active] == ['a', 'c']
        assert [f['body'] for f in archived] == ['b', 'd']
        assert len(active) + len(archived) == len(fbs)

    def test_partition_tolerates_bad_input(self):
        assert partition_feedbacks(None) == ([], [])
        assert partition_feedbacks([None, 'x', {'status': '未対応'}]) == ([{'status': '未対応'}], [])


class TestTimeline:
    def test_full_history(self):
        t = fb_timeline(_fb('対応完了', HIST))
        assert t['received'] == '2026-08-04 17:25'
        assert t['fixed_at'] == '2026-10-26 17:00' and t['fixed_note'] == '全て選択ボタンを追加'
        assert t['done_at'] == '2026-10-27 09:30' and t['done_note'] == '確認しました'
        assert t['legacy'] is False

    def test_legacy_done_without_history(self):
        t = fb_timeline(_fb('対応済み'))
        assert t['legacy'] is True and t['fixed_at'] == '' and t['done_at'] == ''
        assert '日時記録なし' in fb_timeline_text(_fb('対応済み'))

    def test_text_order_and_notes(self):
        txt = fb_timeline_text(_fb('対応完了', HIST))
        assert txt == ('受付 2026-08-04 17:25 → 高井対応済み 2026-10-26 17:00（全て選択ボタンを追加）'
                       ' → 対応完了 2026-10-27 09:30（確認しました）')

    def test_returned_after_fix_is_shown(self):
        hist = HIST[:1] + [{'status': '差し戻し', 'note': 'まだ直っていない', 'by': 'お客様', 'at': '2026-10-27 10:00'}]
        txt = fb_timeline_text(_fb('差し戻し', hist))
        assert '差し戻し 2026-10-27 10:00（まだ直っていない）' in txt
        assert '対応完了' not in txt

    def test_latest_event_wins_when_repeated(self):
        hist = HIST + [{'status': '差し戻し', 'note': 'x', 'by': 'お客様', 'at': '2026-10-28 09:00'},
                       {'status': '高井対応済み', 'note': '再修正', 'by': '高井', 'at': '2026-10-29 09:00'},
                       {'status': '対応完了', 'note': 'OK', 'by': 'お客様', 'at': '2026-10-30 09:00'}]
        t = fb_timeline(_fb('対応完了', hist))
        assert t['fixed_note'] == '再修正' and t['done_note'] == 'OK'

    def test_bad_history_entries_ignored(self):
        t = fb_timeline(_fb('対応完了', ['junk', None, {'status': '対応完了', 'at': '2026-10-27 09:30'}]))
        assert t['done_at'] == '2026-10-27 09:30'


class TestReleaseLog:
    def test_rows_only_archived_newest_first(self):
        fbs = [_fb('未対応', body='open'),
               _fb('対応完了', HIST, body='older', timestamp='2026-08-04 17:25'),
               _fb('対応完了', [{'status': '高井対応済み', 'note': 'B', 'at': '2026-11-01 10:00'},
                              {'status': '対応完了', 'note': '', 'at': '2026-11-02 10:00'}],
                   body='newer', timestamp='2026-09-01 10:00'),
               _fb('対応済み', body='legacy', timestamp='2026-07-21 05:22')]
        rows = release_log_rows(fbs)
        assert [r['ご指摘'] for r in rows] == ['newer', 'older', 'legacy']
        assert rows[0]['修正反映'] == '2026-11-01 10:00' and rows[0]['変更箇所'] == 'B'
        assert rows[2]['対応完了'] == '旧データ'

    def test_project_name_resolver(self):
        rows = release_log_rows([_fb('対応完了', HIST, _proj_name='A邸（プラン2）')],
                                project_name_of=lambda fb: fb.get('_proj_name'))
        assert rows[0]['案件'] == 'A邸（プラン2）'
