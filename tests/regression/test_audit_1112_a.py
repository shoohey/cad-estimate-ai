"""2026-11-12 監査で確認された不具合の再現テスト（第1群: 画面停止・金額の誤り・入力の消失）

[0] 実績にしかない工事種別の番号がAI側と重なると、実績比較タブが重複キーで止まる
[1] 同じ品名で摘要違いの行があると、「同じ商品」登録がチェックしていない行に効く
[3] 「統計を更新」ボタンが NameError で止まる
[17] PDF読み取り結果の部屋表を続けて直すと、1回おきに修正が消える
[25] 打合せ変更の「追加」の原価が 0 円になる
[27] 仕様変更の文のサイズ数字（1616 等）が見積単価になる
[28] 「2700mmの」「幅2700」「W:2700」が数量 2700 と読まれる
[29] フィードバックを1件更新すると、次のカードに前の選択とコメントが残る
"""
import json
import os

import pytest
from streamlit.proto.WidgetStates_pb2 import WidgetState

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401
from ui_helpers import (alias_pairs_from_library, make_alias_entry, parse_money_jp, parse_quantity_jp,
                        proposal_number, strip_dimension_tokens)

SAMPLE_XLSX = os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx")


# ---------------------------------------------------------------- [27][28] 数字の読み方
class TestNumbers:
    @pytest.mark.parametrize("text,expected", [
        ("TOTO サザナ 1616サイズ", None),
        ("I型2550 に変更", None),
        ("514,300円", 514300.0),
        ("150万", 1500000.0),
        ("¥371,500", 371500.0),
        ("見積 371,500円（W2700）", 371500.0),
        ("", None), (None, None),
    ])
    def test_money_needs_currency_mark(self, text, expected):
        assert parse_money_jp(text) == expected

    @pytest.mark.parametrize("text,expected", [
        ("2700mmのカップボードを2台", 2.0),
        ("幅2700のカップボード 1台", 1.0),
        ("W:2700 1台", 1.0),
        ("W＝2700 1台", 1.0),
        ("間口 1820 の収納 3ヶ所", 3.0),
        ("フローリング 30㎡", 30.0),
        ("カップボード W2700", None),
    ])
    def test_quantity_needs_unit(self, text, expected):
        assert parse_quantity_jp(text) == expected

    def test_quantity_bare_number_allowed_only_when_asked(self):
        assert parse_quantity_jp("数量を3に") is None
        assert parse_quantity_jp("数量を3に", allow_bare=True) == 3.0
        assert parse_quantity_jp("W2700", allow_bare=True) is None

    def test_strip_dimension_variants(self):
        for t in ("2700mmの", "2700mmタイプ", "幅2700", "W:2700", "奥行き 650", "高さ850"):
            assert not any(ch.isdigit() for ch in strip_dimension_tokens(t)), t

    @pytest.mark.parametrize("ctype,proposed,amount,expected", [
        ("spec", "TOTO サザナ 1616サイズ", None, 0),
        ("spec", "TOTO サザナ 1616サイズ 371,500円", None, 371500),
        ("price", "514,300", None, 514300),        # 単価変更は単位なしの数字も金額
        ("add", "エコキュート 450,000円 追加", None, 450000),
        ("add", "カップボード W2700", None, 0),
        ("quantity", "2700mmのカップボードを2台", None, 2),
        ("quantity", "3", None, 3),                  # 数量変更は単位なしの数字も数量
        ("spec", "何でも", 250000, 250000),          # 表で入れた金額は最優先
    ])
    def test_proposal_number(self, ctype, proposed, amount, expected):
        assert proposal_number(ctype, proposed, amount) == expected


# ---------------------------------------------------------------- [1] 同じ商品の登録は摘要まで区別
def test_alias_with_summary_only_hits_that_row():
    ai = [{'name': '打合せ造作', 'summary': '', 'quantity': 1, 'unit': '式', 'amount': 50000}]
    ac = [{'name': '床ｶﾞﾗﾘ変更', 'summary': '塗装費のみ', 'quantity': 1, 'unit': '式', 'amount': 49000},
          {'name': '床ｶﾞﾗﾘ変更', 'summary': '樹脂→オーク無垢', 'quantity': 1, 'unit': '式', 'amount': 20000}]
    entry = make_alias_entry({'AI品名': '打合せ造作', 'AI摘要': ''},
                             {'実績品名': '床ｶﾞﾗﾘ変更', '実績摘要': '樹脂→オーク無垢'}, '木工事', 'サンプル邸', 'd')
    out = excel_diff._match_items(ai, ac, alias_pairs_from_library([entry]))
    assert len(out['matched']) == 1
    assert out['matched'][0]['actual_summary'] == '樹脂→オーク無垢'
    assert [x['summary'] for x in out['actual_only']] == ['塗装費のみ']


# ---------------------------------------------------------------- [0] 種別番号の重複で止まらない
def test_duplicate_category_numbers_do_not_crash(app_dir):
    with open(SAMPLE_XLSX, "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    wood = next(c for c in actual["categories"] if c["name"] == "木工事")
    actual["categories"].append({"no": wood["no"], "name": "金属工事", "total": 235800, "groups": [],
                                 "items": [{"name": "笠木", "summary": "", "quantity": 1.0, "unit": "式",
                                            "price": 235800, "amount": 235800}]})
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    at = _boot(app_dir)
    assert not at.exception, [str(e.value) for e in at.exception]
    assert [b for b in at.button if b.key == "save_diff_notes"], "差異理由の保存ボタンまで描画されていない"
    assert [t for t in at.text_area if t.key == "pdf_hints"], "後ろのタブが描画されていない"
    os.remove(os.path.join(app_dir, "projects", "sample_furui_actual.json"))


# ---------------------------------------------------------------- [25] 追加の原価
def test_added_item_gets_linked_cost():
    import app
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    est = calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                             load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")))
    mods = {}
    app._apply_ai_proposal(est, mods, {'change_type': 'add', 'category': '住宅設備機器工事',
                                       'item_name': 'エコキュート', 'proposed': '450,000円', 'amount': 450000})
    assert mods['new_items'][0]['estimate_price'] == 450000
    assert mods['new_items'][0]['order_price'] == 301500


def test_spec_change_with_size_number_keeps_price():
    import app
    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices
    est = calculate_estimate(load_cad_file(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT")),
                             load_unit_prices(os.path.join(ROOT, "tankamaster_updated.csv")))
    ub = next((c, i, it) for c in est.categories for i, it in enumerate(c.items) if 'UB' in it.name or 'ﾕﾆｯﾄﾊﾞｽ' in it.name or 'ユニットバス' in it.name)
    cat, idx, item = ub
    mods = {}
    app._apply_ai_proposal(est, mods, {'change_type': 'spec', 'category': cat.name, 'item_name': item.name,
                                       'proposed': 'TOTO サザナ 1616サイズ', 'amount': 0})
    ov = mods['overrides'][f"{cat.no}_{idx}"]
    assert ov['estimate_price'] == item.estimate_price, "サイズの数字が単価になった"
    assert ov['name'] == 'TOTO サザナ 1616サイズ'


# ---------------------------------------------------------------- [3] 統計を更新
def test_recompute_hidden_stats_button(app_dir):
    with open(SAMPLE_XLSX, "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    at = _boot(app_dir)
    at.session_state["proj_sample_furui_pdf_extracted"] = {"file": "t.pdf", "raw": {"total_floor_area_m2": 100.0, "floors": 2},
                                                          "data": {"total_floor_area_m2": 100.0, "floors": 2}}
    at.run()
    btn = [b for b in at.button if b.key == "btn_recompute_hidden"]
    assert btn, "統計を更新ボタンが無い"
    btn[0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    with open(os.path.join(app_dir, "projects", "hidden_item_stats.json"), encoding="utf-8") as f:
        stats = json.load(f)
    assert stats["n_properties"] >= 1 and stats["items"]
    os.remove(os.path.join(app_dir, "projects", "sample_furui_actual.json"))


# ---------------------------------------------------------------- [17] PDF表の連続修正
EXTRACT = {
    'confidence': 'medium', 'notes': [], 'total_floor_area_m2': 100.0, 'floors': 2, 'building_area_m2': 60.0,
    'rooms': [{'name': 'LDK', 'floor': '1階', 'area_m2': 30.0}, {'name': '洗面', 'floor': '1階', 'area_m2': 3.0},
              {'name': 'トイレ', 'floor': '1階', 'area_m2': 1.5}, {'name': '収納', 'floor': '2階', 'area_m2': 3.3}],
    'fittings': [{'kind': '内部建具', 'name': 'WD1', 'width_mm': 780, 'height_mm': 2000, 'count': 1, 'location': '洗面'}],
}


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
    assert len(ids) == 1, ids
    return ids[0]


def test_pdf_room_edits_all_survive(app_dir):
    at = _boot(app_dir)
    at.session_state["proj_sample_furui_pdf_extracted"] = {"file": "test.pdf", "raw": json.loads(json.dumps(EXTRACT)),
                                                          "data": json.loads(json.dumps(EXTRACT))}
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    cur_id, edits = None, {}
    for i, v in enumerate([31.0, 4.0, 3.5, 3.6]):
        eid = _find_id(at, "pdf_rooms_test_pdf")
        if eid != cur_id:          # ブラウザはIDが変わると編集履歴を捨てる
            cur_id, edits = eid, {}
        edits.setdefault(str(i), {})["area_m2"] = v
        ws = at._tree.get_widget_states()
        keep = [w for w in ws.widgets if w.id != eid]
        del ws.widgets[:]
        ws.widgets.extend(keep)
        ws.widgets.append(WidgetState(id=eid, string_value=json.dumps(
            {"edited_rows": edits, "added_rows": [], "deleted_rows": []})))
        at._run(ws)
        assert not at.exception, [str(e.value) for e in at.exception]
    rooms = at.session_state["proj_sample_furui_pdf_extracted"]["data"]["rooms"]
    assert [r["area_m2"] for r in rooms] == [31.0, 4.0, 3.5, 3.6]


# ---------------------------------------------------------------- [29] FBカードの入力が次のカードに残らない
def test_feedback_status_inputs_do_not_carry_over(app_dir):
    fb_path = os.path.join(app_dir, "projects", "sample_furui_feedbacks.json")
    with open(fb_path, "w", encoding="utf-8") as f:
        json.dump([
            {"category": "操作性", "body": "全選択が欲しい", "timestamp": "2026-08-04 17:25", "name": "テスト",
             "status": "高井対応済み", "feedback_id": "fb-a", "attachments": []},
            {"category": "操作性", "body": "タイトルだけ表示して", "timestamp": "2026-08-04 17:31", "name": "テスト",
             "status": "高井対応済み", "feedback_id": "fb-b", "attachments": []},
        ], f, ensure_ascii=False)
    at = _boot(app_dir)
    sels = [s for s in at.selectbox if s.key and s.key.startswith("status_fb_") and s.key.endswith("_sel")]
    notes = [t for t in at.text_input if t.key and t.key.startswith("status_fb_") and t.key.endswith("_note")]
    assert len(sels) == 2 and len(notes) == 2
    first_sel, first_note = sels[0], notes[0]     # 新しい順なので「タイトルだけ…」が先
    first_sel.select("対応完了")
    first_note.input("確認しました")
    btn_key = first_sel.key[:-len("_sel")]
    [b for b in at.button if b.key == btn_key][0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    sels = [s for s in at.selectbox if s.key and s.key.startswith("status_fb_") and s.key.endswith("_sel")]
    notes = [t for t in at.text_input if t.key and t.key.startswith("status_fb_") and t.key.endswith("_note")]
    assert len(sels) == 1
    assert sels[0].value != "対応完了" and notes[0].value == ""
    os.remove(fb_path)
