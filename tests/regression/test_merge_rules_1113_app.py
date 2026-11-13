"""2026-11-13 打合せ: 実績比較の「まとめて計上」を app.py 本体（AppTest）で確認

本番の既存案件と同じ「安全対策費 387,300 の1行だけ」の実績で、
金額が違う1行（安全対策費）＋AIにしかない3行（駐車場代・基礎廻り養生ｼｰﾄ・整地費用）をチェックして登録すると、
次からは1行にまとめて「一致」になる。
"""
import json
import os

from streamlit.proto.WidgetStates_pb2 import WidgetState

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401
from tests.regression.test_merge_rules_1113 import _old_style_actual

SAFETY_NAMES = {"安全対策費", "駐車場代", "基礎廻り養生ｼｰﾄ", "整地費用"}


def _setup(app_dir):
    with open(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx"), "rb") as f:
        actual = _old_style_actual(excel_diff.parse_actual_estimate(f.read()))
    proj = os.path.join(app_dir, "projects")
    with open(os.path.join(proj, "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    for name in ("item_merge_library.json", "item_alias_library.json", "sample_furui_diff_notes.json"):
        p = os.path.join(proj, name)
        if os.path.exists(p):
            os.remove(p)


def _find_id(at, key_suffix):
    ids = []

    def walk(node):
        proto = getattr(node, 'proto', None)
        pid = getattr(proto, 'id', None) if proto is not None else None
        if isinstance(pid, str) and pid.endswith(key_suffix):
            ids.append(pid)
        ch = getattr(node, 'children', None)
        for c in (ch.values() if isinstance(ch, dict) else []):
            walk(c)
    walk(at._tree)
    assert len(ids) == 1, (key_suffix, ids)
    return ids[0]


def _kasetsu_ci(at):
    heads = [e.label for e in at.expander if "AI " in e.label and "実績 " in e.label and "差額 " in e.label]
    return [i for i, h in enumerate(heads) if h.startswith("仮設工事")][0], heads


def _reg_table(at):
    tables = [d.value for d in at.dataframe if "選択" in list(d.value.columns)
              and "駐車場代" in set(d.value["AI品名"])]
    assert len(tables) == 1
    return tables[0]


def _run_with(at, overrides: dict):
    """今の画面の入力に、表のチェック・ボタン押下などを重ねて1回だけ実行する"""
    ws = at._tree.get_widget_states()
    keep = [w for w in ws.widgets if w.id not in overrides]
    del ws.widgets[:]
    ws.widgets.extend(keep)
    ws.widgets.extend(overrides.values())
    at._run(ws)
    assert not at.exception, [str(e.value) for e in at.exception]


def _check_rows(eid, idx):
    return WidgetState(id=eid, string_value=json.dumps(
        {"edited_rows": {str(i): {"選択": True} for i in idx}, "added_rows": [], "deleted_rows": []}))


def _merge_lib(app_dir):
    p = os.path.join(app_dir, "projects", "item_merge_library.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else []


def _detail_rows(at):
    tables = [d.value for d in at.dataframe if "区分" in list(d.value.columns)
              and "差額(AI−実績)" in list(d.value.columns) and "選択" not in list(d.value.columns)]
    return [r for t in tables for _, r in t.iterrows()]


def test_register_merge_from_checked_rows(app_dir):
    _setup(app_dir)
    at = _boot(app_dir)
    assert not at.exception, [str(e.value) for e in at.exception]
    ci, _ = _kasetsu_ci(at)
    # 金額が違う行（安全対策費）も登録の表に出る（以前は AIにしかない／実績にしかない だけだった）
    reg = _reg_table(at)
    assert "金額が違う" in set(reg["区分"])
    idx = [i for i, r in reg.iterrows() if r["AI品名"] in SAFETY_NAMES]
    assert len(idx) == 4
    # 登録の種類は4つから選べる（二択だけではない）
    kind = at.radio(key=f"reg_kind_sample_furui_{ci}")
    assert len(kind.options) == 4 and kind.value.startswith("まとめて計上")

    eid = _find_id(at, f"alias_editor_sample_furui_{ci}_0")
    _run_with(at, {eid: _check_rows(eid, idx)})
    assert any("合計が一致" in s.value and "387,300" in s.value for s in at.success), [s.value for s in at.success]

    note_id = at.text_input(key=f"reg_note_sample_furui_{ci}_0").proto.id
    btn_id = at.button(key=f"reg_btn_sample_furui_{ci}").proto.id
    _run_with(at, {eid: _check_rows(eid, idx),
                   note_id: WidgetState(id=note_id, string_value="お客様に見せないため安全対策費に含める"),
                   btn_id: WidgetState(id=btn_id, trigger_value=True)})
    lib = _merge_lib(app_dir)
    assert len(lib) == 1
    e = lib[0]
    assert {m["name"] for m in e["ai_members"]} == SAFETY_NAMES
    assert [m["name"] for m in e["actual_members"]] == ["安全対策費"]
    assert e["note"] == "お客様に見せないため安全対策費に含める" and e["scope"] == "global"
    assert e["category"] == "仮設工事"

    # 登録後: 仮設工事は1行にまとまって「一致」、見出しに「まとめ 1」、一覧に1件
    rows = _detail_rows(at)
    merged = [r for r in rows if str(r["扱い"]).startswith("まとめ")]
    assert len(merged) == 1 and merged[0]["区分"] == "一致" and int(merged[0]["AI金額"]) == 387300
    assert not [r for r in rows if r["区分"] == "AIにしかない" and r["AI品名"] in SAFETY_NAMES]
    _, heads = _kasetsu_ci(at)
    assert any(h.startswith("仮設工事") and "まとめ 1" in h for h in heads)
    assert any("登録済みの「まとめて計上」（この案件に適用 1件）" in e.label for e in at.expander)
    assert any("補足: お客様に見せないため" in c.value for c in at.caption)


def test_register_merge_from_suggestion_and_undo(app_dir):
    _setup(app_dir)
    at = _boot(app_dir)
    ci, _ = _kasetsu_ci(at)
    btn = at.button(key=f"merge_sugg_sample_furui_{ci}_0")
    btn.click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    lib = _merge_lib(app_dir)
    assert len(lib) == 1 and {m["name"] for m in lib[0]["ai_members"]} == SAFETY_NAMES
    # 解除すると元の分かれた表示に戻る
    at.button(key="merge_del_0").click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    assert _merge_lib(app_dir) == []
    rows = _detail_rows(at)
    assert {r["AI品名"] for r in rows if r["区分"] == "AIにしかない"} >= {"駐車場代", "基礎廻り養生ｼｰﾄ", "整地費用"}


def test_merge_validation_error_is_shown(app_dir):
    """1行だけチェックして「まとめて計上」を押すと、理由を画面に出して何も登録しない"""
    _setup(app_dir)
    at = _boot(app_dir)
    ci, _ = _kasetsu_ci(at)
    reg = _reg_table(at)
    idx = [i for i, r in reg.iterrows() if r["AI品名"] == "駐車場代"]
    eid = _find_id(at, f"alias_editor_sample_furui_{ci}_0")
    btn_id = at.button(key=f"reg_btn_sample_furui_{ci}").proto.id
    _run_with(at, {eid: _check_rows(eid, idx), btn_id: WidgetState(id=btn_id, trigger_value=True)})
    assert any("実績側" in e.value for e in at.error), [e.value for e in at.error]
    assert _merge_lib(app_dir) == []


def test_project_specific_with_note(app_dir):
    """「この案件特有」を選んで補足を書くと、その補足が理由として残る"""
    _setup(app_dir)
    at = _boot(app_dir)
    ci, _ = _kasetsu_ci(at)
    at.radio(key=f"reg_kind_sample_furui_{ci}").set_value("この案件特有（図面に表れない項目）")
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    reg = _reg_table(at)
    idx = [i for i, r in reg.iterrows() if r["AI品名"] == "駐車場代"]
    eid = _find_id(at, f"alias_editor_sample_furui_{ci}_0")
    note_id = at.text_input(key=f"reg_note_sample_furui_{ci}_0").proto.id
    btn_id = at.button(key=f"reg_btn_sample_furui_{ci}").proto.id
    _run_with(at, {eid: _check_rows(eid, idx),
                   note_id: WidgetState(id=note_id, string_value="敷地内に駐車できるため"),
                   btn_id: WidgetState(id=btn_id, trigger_value=True)})
    notes = json.load(open(os.path.join(app_dir, "projects", "sample_furui_diff_notes.json"), encoding="utf-8"))
    hit = [v for k, v in notes.items() if k.startswith("仮設工事|駐車場代|")]
    assert hit and hit[0]["reason"] == "敷地内に駐車できるため" and hit[0]["scope"] == "project"
    assert _merge_lib(app_dir) == []


def test_note_in_other_category_survives_registration(app_dir):
    """仮設工事で登録しても、基礎工事に書きかけの補足は消えない（リセットは登録した工事種別だけ）"""
    _setup(app_dir)
    at = _boot(app_dir)
    ci, heads = _kasetsu_ci(at)
    ci2 = [i for i, h in enumerate(heads) if h.startswith("基礎工事")][0]
    at.text_input(key=f"reg_note_sample_furui_{ci2}_0").input("書きかけ")
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    kiso_editor_before = _find_id(at, f"alias_editor_sample_furui_{ci2}_0")
    at.button(key=f"merge_sugg_sample_furui_{ci}_0").click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    assert len(_merge_lib(app_dir)) == 1
    # 登録した仮設工事の表・補足だけが作り直され（キーの世代が進む）、基礎工事のキーは変わらない。
    # ブラウザは同じキーの入力値を保持して次の操作で送り直すので、キーが変わらなければ書きかけは消えない
    # （実ブラウザで「書きかけ」が残ることを 2026-11-13 に確認。AppTest にはフロントエンドが無く、
    #  st.rerun() で描画されなかった入力の値は持ち越せないため、ここではキーの安定性を確かめる）
    note_keys = {t.key for t in at.text_input if str(t.key).startswith("reg_note_sample_furui_")}
    assert f"reg_note_sample_furui_{ci2}_0" in note_keys, sorted(note_keys)
    assert f"reg_note_sample_furui_{ci}_1" in note_keys and f"reg_note_sample_furui_{ci}_0" not in note_keys
    assert _find_id(at, f"alias_editor_sample_furui_{ci2}_0") == kiso_editor_before
    assert not [t for t in at.text_input if str(t.key).startswith(f"reg_note_sample_furui_{ci2}_1")]


def test_scope_button_says_origin_when_viewed_from_other_project(app_dir):
    """登録元と違う案件から見たときは「登録元の案件のみに変更」と出す（押した案件だけに効くと誤解しない）"""
    _setup(app_dir)
    from ui_helpers import make_merge_entry
    e = make_merge_entry([{"name": "安全対策費", "summary": ""}, {"name": "駐車場代", "summary": ""}],
                         [{"name": "安全対策費", "summary": ""}], "仮設工事", "別の邸", "2026-11-13 10:00",
                         project_id="other_pid")
    with open(os.path.join(app_dir, "projects", "item_merge_library.json"), "w", encoding="utf-8") as f:
        json.dump([e], f, ensure_ascii=False)
    at = _boot(app_dir)
    labels = [b.label for b in at.button if b.key and str(b.key).startswith("merge_scope_toggle_")]
    assert labels == ["登録元の案件のみに変更"], labels
