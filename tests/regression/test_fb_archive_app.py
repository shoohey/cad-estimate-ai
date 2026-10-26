"""問い合わせの保持・アーカイブ表示を app.py 本体（AppTest）で確認"""
import json
import os

from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud  # noqa: F401


def _write_fbs(app_dir, fbs):
    with open(os.path.join(app_dir, "projects", "sample_furui_feedbacks.json"), "w", encoding="utf-8") as f:
        json.dump(fbs, f, ensure_ascii=False)


def test_archive_view_and_retention(app_dir):
    _write_fbs(app_dir, [
        {"category": "操作性", "body": "まだ直っていない件", "timestamp": "2026-08-04 17:25", "name": "テスト",
         "status": "未対応", "feedback_id": "a", "attachments": []},
        {"category": "操作性", "body": "直って確認済みの件", "timestamp": "2026-07-21 05:22", "name": "テスト",
         "status": "対応完了", "feedback_id": "b", "attachments": [],
         "status_history": [{"status": "高井対応済み", "note": "ボタン追加", "by": "高井", "at": "2026-10-26 17:00"},
                            {"status": "対応完了", "note": "OK", "by": "お客様", "at": "2026-10-27 09:30"}]},
        {"category": "その他", "body": "旧データの完了件", "timestamp": "2026-07-25 01:45", "name": "テスト",
         "status": "対応済み", "feedback_id": "c", "attachments": []},
    ])
    at = _boot(app_dir)
    view = [r for r in at.radio if r.key == "fb_archive_view"][0]
    assert list(view.options) == ["対応中（1）", "アーカイブ・対応完了（2）", "全て"]
    # 既定は対応中: 未対応の1件だけがカードに出る
    md = " ".join(m.value for m in at.markdown)
    assert "まだ直っていない件" in md and "直って確認済みの件" not in md
    # 一括クリアは無い／削除は確認チェックが無いと押せない
    assert not [b for b in at.button if b.label == "この案件のFBをクリア"]
    del_btn = [b for b in at.button if b.key == "del_fb_sample_furui_0"][0]
    assert del_btn.disabled is True

    # アーカイブ表示: 完了2件と履歴・修正履歴テーブル
    view.set_value("アーカイブ・対応完了（2）").run()
    assert not at.exception, [str(e.value) for e in at.exception]
    md = " ".join(m.value for m in at.markdown)
    assert "直って確認済みの件" in md and "旧データの完了件" in md and "まだ直っていない件" not in md
    assert "高井対応済み 2026-10-26 17:00（ボタン追加） → 対応完了 2026-10-27 09:30（OK）" in md
    assert "日時記録なし" in md
    logs = [d.value for d in at.dataframe if "変更箇所" in list(d.value.columns)]
    assert logs, "修正履歴テーブルが無い"
    log = logs[0]
    assert list(log["変更箇所"])[0] == "ボタン追加"
    assert set(log["対応完了"]) == {"2026-10-27 09:30", "旧データ"}

    # 全て表示: 3件（ラジオは最新の画面ツリーから取り直す）
    view = [r for r in at.radio if r.key == "fb_archive_view"][0]
    view.set_value("全て").run()
    assert not at.exception, [str(e.value) for e in at.exception]
    md = " ".join(m.value for m in at.markdown)
    assert all(s in md for s in ("まだ直っていない件", "直って確認済みの件", "旧データの完了件"))
    # ファイルは3件のまま（何も消えていない）
    with open(os.path.join(app_dir, "projects", "sample_furui_feedbacks.json"), encoding="utf-8") as f:
        assert len(json.load(f)) == 3
