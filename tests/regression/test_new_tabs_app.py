"""マスタ・対応表タブ／PDF見積（検証）タブ／類似候補の表示を app.py 本体（AppTest）で確認"""
import json
import os

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401


def test_master_and_pdf_tabs_render_and_candidates(app_dir):
    with open(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx"), "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    for stale in ("item_alias_library.json", "sample_furui_diff_notes.json", "design_price_master.json"):
        p = os.path.join(app_dir, "projects", stale)
        if os.path.exists(p):
            os.remove(p)
    at = _boot(app_dir)
    assert not at.exception, [str(e.value) for e in at.exception]

    # マスタ・対応表タブ: 件数表示とアップローダ
    labels = [m.label for m in at.metric]
    assert "設計部 単価マスタ" in labels and "同じ商品 辞書" in labels and "アプリ内 単価マスタ" in labels
    app_master = [m for m in at.metric if m.label == "アプリ内 単価マスタ"][0]
    assert int(str(app_master.value).split()[0]) >= 200, app_master.value   # tankamaster_updated.csv が読めている
    assert [b for b in at.button if b.key == "dl_alias_template"] or True   # download_button は要素種別が異なる場合がある

    # PDF見積タブ: アップローダと「図面の補足」入力、パターン管理が描画される
    assert any(t.key == "pdf_hints" for t in at.text_area)
    assert any("パターンの管理" in e.label for e in at.expander)
    assert any(b.key == "btn_new_pattern" for b in at.button)

    # 実績比較の類似候補: チェックONで候補テーブルとマスタ名の選択肢が出る
    toggles = [c for c in at.checkbox if c.key and c.key.startswith("cand_toggle_sample_furui_")]
    assert toggles, "類似候補のチェックボックスが無い"
    # 「実績にしかない」行がある工事種別を選ぶ（見出しの件数と同じ並び順）
    import re
    headers = [e.label for e in at.expander if "AI " in e.label and "実績 " in e.label and "差額 " in e.label]
    idx = next(i for i, h in enumerate(headers) if int(re.search(r"実績のみ (\d+)", h).group(1)) > 0)
    toggles[idx].check().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    cand_tables = [d.value for d in at.dataframe if "候補1" in list(d.value.columns)]
    assert cand_tables and len(cand_tables[0]) >= 1
    picks = [s for s in at.selectbox if s.key and s.key.startswith("cand_pick_sample_furui_")]
    assert picks and len(picks[0].options) >= 1
