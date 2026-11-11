"""2026-11-12 監査で確認された不足の再現テスト（第4群: 図面の造作・パターンを案件の見積へ）

[13] PDF図面の可動棚・カウンターなどの造作を、単価マスタの項目として見積に入れる
[24] パターン（図面に現れない項目のセット）を、PDFを読み取らない案件でも見積に入れる
"""
import json
import os

import ai_client
from tests.regression.test_audit_1112_c import DESIGN_ROWS
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud  # noqa: F401


def test_normalize_built_ins():
    out = ai_client.normalize_extraction({
        'built_ins': [{'name': '可動棚', 'spec': 'W4.5尺 5枚', 'count': '2', 'location': 'WIC'},
                      {'name': '', 'spec': 'x'}, 'junk', {'name': 'カウンター', 'count': 0}]})
    assert out['built_ins'] == [
        {'name': '可動棚', 'spec': 'W4.5尺 5枚', 'count': 2, 'location': 'WIC'},
        {'name': 'カウンター', 'spec': None, 'count': 1, 'location': None},   # 無い値は None（他の項目と同じ）
    ]
    assert tuple(out)[-1] == 'built_ins'


def test_pdf_built_ins_to_case_estimate(app_dir):
    with open(os.path.join(app_dir, "projects", "design_price_master.json"), "w", encoding="utf-8") as f:
        json.dump({'rows': DESIGN_ROWS, 'file': 'design.xlsx', 'at': '2026-11-12'}, f, ensure_ascii=False)
    at = _boot(app_dir)
    ext = {'total_floor_area_m2': 100.0, 'floors': 2, 'rooms': [], 'fittings': [],
           'built_ins': [{'name': '可動棚', 'spec': 'W4.5尺 5枚 ﾌｨｯﾄﾗｯｸ', 'count': 2, 'location': 'WIC'}]}
    at.session_state["proj_sample_furui_pdf_extracted"] = {"file": "t.pdf", "raw": ext, "data": ext}
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    btn = [b for b in at.button if b.key == "btn_builtins_to_case"]
    assert btn, "造作を案件の見積に入れるボタンが無い"
    btn[0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    items = [i for i in at.session_state["proj_sample_furui_mods"]["new_items"] if "可動棚" in i["name"]]
    assert items and items[0]["quantity"] == 2.0 and items[0]["estimate_price"] == 77500
    # 2回押しても二重に入らない
    [b for b in at.button if b.key == "btn_builtins_to_case"][0].click().run()
    assert len([i for i in at.session_state["proj_sample_furui_mods"]["new_items"] if "可動棚" in i["name"]]) == 1
    os.remove(os.path.join(app_dir, "projects", "design_price_master.json"))


def test_patterns_from_edit_tab_without_pdf(app_dir):
    with open(os.path.join(app_dir, "projects", "hidden_item_patterns.json"), "w", encoding="utf-8") as f:
        json.dump([{'id': 'pat_x', 'name': '全館空調仕様', 'description': '', 'at': '',
                    'items': [{'key': '空調工事|全館空調', 'category': '空調工事', 'name': '全館空調', 'summary': '',
                               'quantity': 1.0, 'unit': '式', 'estimate_price': 800000, 'order_price': 536000}]}],
                  f, ensure_ascii=False)
    at = _boot(app_dir)
    ms = [m for m in at.multiselect if m.key == "mod_patterns"]
    assert ms, "見積修正タブにパターンの選択欄が無い"
    ms[0].select("全館空調仕様").run()
    [b for b in at.button if b.key == "btn_mod_patterns"][0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    items = [i for i in at.session_state["proj_sample_furui_mods"]["new_items"] if i["name"] == "全館空調"]
    assert len(items) == 1 and items[0]["estimate_price"] == 800000
    [b for b in at.button if b.key == "btn_mod_patterns"][0].click().run()
    assert len([i for i in at.session_state["proj_sample_furui_mods"]["new_items"] if i["name"] == "全館空調"]) == 1
    os.remove(os.path.join(app_dir, "projects", "hidden_item_patterns.json"))
