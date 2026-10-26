"""2026-08-07 MTG 宿題 #1（原価連動）の通し確認: 打合せ変更 → 全て選択 → 反映 で原価も更新される。"""
import json
import os
import re

from streamlit.testing.v1 import AppTest

from tests.regression.test_feedback_20260804_app import _boot, _apply_button, app_dir, _no_cloud  # noqa: F401


def test_price_change_links_order_price(app_dir):
    at = _boot(app_dir)
    at.session_state["proj_sample_furui_ai_proposals"] = [
        {'category': '住宅設備機器工事', 'item_name': 'キッチン', 'change_type': 'price',
         'proposed': '514,300円', 'amount': 514300, 'confidence': 'high', 'reason': '打合せ'},
    ]
    at.session_state["proj_sample_furui_ai_proposals_src"] = "提案書反映（test.pdf）"
    at.run()
    [b for b in at.button if b.key == "btn_props_select_all"][0].click().run()
    assert _apply_button(at).label == "選択した1件を反映"
    _apply_button(at).click().run()
    assert not at.exception, [str(e.value) for e in at.exception]

    # 変更見積（別プラン）が作られ、保存ファイルに反映内容が入っていること
    plans = [pid for pid in at.session_state["projects"] if pid.startswith("sample_furui__plan")]
    assert plans, "変更見積（別プラン）が作られていない"
    target = plans[-1]
    # 反映後は作成した変更見積が開かれたままになる（元見積に戻されない）
    assert at.session_state["current_project"] == target
    assert at.session_state["project_selector"] == target
    with open(os.path.join(app_dir, "projects", f"{target}_data.json"), encoding="utf-8") as f:
        mods = json.load(f)["mods"]
    changed = [ov for ov in mods["overrides"].values() if ov.get("estimate_price") == 514300]
    assert changed, mods["overrides"]
    ov = changed[0]
    # 変更履歴の「見積単価 A円 → B円（原価 C円 → D円）」から元の原価率を取り出して検算
    with open(os.path.join(app_dir, "projects", f"{target}_changelog.json"), encoding="utf-8") as f:
        logs = json.load(f)
    msg = next(l for l in logs if "原価" in json.dumps(l, ensure_ascii=False))
    m = re.search(r"見積単価 ([\d,]+)円 → ([\d,]+)円（原価 ([\d,]+)円 → ([\d,]+)円）", json.dumps(msg, ensure_ascii=False))
    assert m, msg
    old_est, new_est, old_ord, new_ord = (int(x.replace(",", "")) for x in m.groups())
    assert new_est == 514300 and ov["order_price"] == new_ord
    expected = round(new_est * old_ord / old_est) if old_est > 0 and old_ord > 0 else round(new_est * 0.67)
    assert new_ord == expected
    assert 0 < new_ord < new_est, "原価が販価を超えている／0になっている"
