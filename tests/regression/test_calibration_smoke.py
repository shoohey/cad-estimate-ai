"""実データ（校正8物件）で全タブが最後まで描画されることを確かめる煙テスト

2026-11-12 監査の教訓: サンプル1物件では、実績にしかない工事種別の番号重複などが出ず、
本番のお客様の物件でだけ画面が止まる不具合を見逃した。calibration/ が無い環境ではスキップ。
"""
import glob
import json
import os
import re
import shutil

import pytest

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT, SAMPLE_TXT  # noqa: F401

CALIB = os.path.join(ROOT, "calibration")
PAIRS = []
for txt in sorted(glob.glob(os.path.join(CALIB, "*【数量データ】*.TXT"))):
    m = re.match(r"(\d+)-", os.path.basename(txt))
    xl = glob.glob(os.path.join(CALIB, f"{m.group(1)}-*【提出見積】*.xlsx*")) if m else []
    if xl:
        PAIRS.append((m.group(1), txt, xl[0]))


@pytest.mark.skipif(not PAIRS, reason="calibration/ が無い環境")
@pytest.mark.parametrize("pid,txt,xlsx", PAIRS, ids=[p[0] for p in PAIRS])
def test_all_tabs_render_with_real_property(app_dir, pid, txt, xlsx):
    # サンプルボタンが読む TXT を、その物件の数量データに差し替える
    shutil.copy(txt, os.path.join(app_dir, SAMPLE_TXT))
    with open(xlsx, "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    try:
        at = _boot(app_dir)
        assert not at.exception, [str(e.value) for e in at.exception]
        assert [b for b in at.button if b.key == "save_diff_notes"], "実績比較が最後まで描画されていない"
        assert [t for t in at.text_area if t.key == "pdf_hints"], "PDF見積タブが描画されていない"
        assert [r for r in at.radio if r.key == "fb_view_mode"], "フィードバックタブが描画されていない"
    finally:
        shutil.copy(os.path.join(ROOT, SAMPLE_TXT), os.path.join(app_dir, SAMPLE_TXT))
        os.remove(os.path.join(app_dir, "projects", "sample_furui_actual.json"))
