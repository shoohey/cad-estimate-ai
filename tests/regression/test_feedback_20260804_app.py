"""2026-08-04 フィードバック2件の再現テスト（Streamlit AppTest で app.py 本体を実行）

本番の Supabase に触れないよう、secrets を含まない作業用コピーでアプリを動かす。
"""
import os
import shutil
import sys

import pytest
from streamlit.testing.v1 import AppTest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SAMPLE_TXT = "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT"

PROPOSALS = [
    {'category': '木工事', 'item_name': 'フローリング', 'change_type': 'spec',
     'proposed': '無垢材に変更', 'amount': 120000, 'confidence': 'high', 'reason': 'p2'},
    {'category': '住宅設備機器工事', 'item_name': 'キッチン', 'change_type': 'price',
     'proposed': '上位グレード', 'amount': 350000, 'confidence': 'low', 'reason': ''},
    {'category': '外構工事', 'item_name': 'カーポート', 'change_type': 'add',
     'proposed': '2台用を追加', 'amount': 0, 'confidence': 'medium', 'reason': '口頭'},
]


@pytest.fixture(scope="module")
def app_dir(tmp_path_factory):
    """secrets 無しのアプリコピー（クラウド永続化は無効＝ローカル保存のみ）

    注意: Streamlit は st.secrets を cwd 基準（<cwd>/.streamlit/secrets.toml）で解決するため、
    コピー先に secrets が無いだけでは不十分。_boot() 側で at.secrets を差し替え、
    cloud_store の設定取得も遮断して二重に本番接続を防ぐ。
    """
    d = tmp_path_factory.mktemp("cad_app")
    for name in os.listdir(ROOT):
        if name.endswith(".py") or name.endswith(".csv"):
            shutil.copy(os.path.join(ROOT, name), d / name)
    # macOS の listdir は NFD で返すので、固定名はパス結合で直接コピーする
    shutil.copy(os.path.join(ROOT, SAMPLE_TXT), d / SAMPLE_TXT)
    assert os.path.exists(d / SAMPLE_TXT), "サンプルTXTがコピーされていない"
    os.makedirs(d / "projects")
    for name in ("index.json", "sample_furui_data.json", "sample_furui.txt", "equipment_master.json"):
        src = os.path.join(ROOT, "projects", name)
        if os.path.exists(src):
            shutil.copy(src, d / "projects" / name)
    os.makedirs(d / ".streamlit")
    shutil.copy(os.path.join(ROOT, ".streamlit", "config.toml"), d / ".streamlit" / "config.toml")
    assert not os.path.exists(d / ".streamlit" / "secrets.toml")
    return str(d)


@pytest.fixture(autouse=True)
def _no_cloud(monkeypatch):
    """本番 Supabase への接続経路を物理的に遮断する（環境変数・st.secrets の両方）"""
    import cloud_store
    for k in ("SUPABASE_URL", "SUPABASE_SERVICE_KEY"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setattr(cloud_store, "_config", lambda: ("", ""))

    def _blocked(*_a, **_k):
        raise cloud_store.CloudStoreError("テスト中のクラウドアクセスは禁止")
    monkeypatch.setattr(cloud_store, "get", _blocked)
    monkeypatch.setattr(cloud_store, "put", _blocked)
    monkeypatch.setattr(cloud_store, "delete", _blocked)
    yield


def _captions(at: AppTest) -> list[str]:
    return [c.value for c in at.sidebar.caption]


def _boot(app_dir: str) -> AppTest:
    at = AppTest.from_file(os.path.join(app_dir, "app.py"), default_timeout=180)
    at.secrets = {"_apptest_isolated": "1"}   # 非空 dict で st.secrets を差し替える（cwd の secrets を読ませない）
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]
    caps = " ".join(_captions(at))
    assert "一時保存" in caps and "クラウド永続化 有効" not in caps, caps
    assert len(at.session_state["projects"]) == 1, "ローカル index.json 以外の案件が読み込まれている"
    # サンプル案件を読み込む
    sample_btn = [b for b in at.sidebar.button if b.label == "サンプルデータで試す"]
    assert sample_btn, "サンプルボタンが見つからない"
    sample_btn[0].click().run()
    assert not at.exception, [str(e.value) for e in at.exception]
    # サンプル読込が実際に走ったこと（ファイル未検出の警告が出ていない）
    assert not [w for w in at.warning if "サンプルデータファイル" in w.value], [w.value for w in at.warning]
    assert at.session_state["current_project"] == "sample_furui"
    assert at.session_state["proj_sample_furui_file_name"] == "古居様邸サンプルデータ"
    return at


def _apply_button(at: AppTest):
    btns = [b for b in at.button if b.key == "btn_apply_props"]
    assert btns, "「選択したN件を反映」ボタンが見つからない"
    return btns[0]


class TestSelectAllClearAllInApp:
    def test_select_all_then_clear_all(self, app_dir):
        at = _boot(app_dir)
        at.session_state["proj_sample_furui_ai_proposals"] = [dict(p) for p in PROPOSALS]
        at.session_state["proj_sample_furui_ai_proposals_src"] = "提案書反映（test.pdf）"
        at.run()
        assert not at.exception, [str(e.value) for e in at.exception]

        # 提案書由来なので初期は全て未チェック
        assert _apply_button(at).label == "選択した0件を反映"

        select_all = [b for b in at.button if b.key == "btn_props_select_all"]
        clear_all = [b for b in at.button if b.key == "btn_props_clear_all"]
        assert select_all and clear_all, "全て選択／全て解除ボタンが無い"
        assert select_all[0].label == "全て選択"
        assert clear_all[0].label == "全て解除"

        select_all[0].click().run()
        assert not at.exception, [str(e.value) for e in at.exception]
        props = at.session_state["proj_sample_furui_ai_proposals"]
        assert [p.get("apply") for p in props] == [True, True, True]
        assert at.session_state["proj_sample_furui_ai_proposals_epoch"] == 1
        assert _apply_button(at).label == "選択した3件を反映"
        # 表の内容（品名・金額）は失われていない
        assert [p["item_name"] for p in props] == ["フローリング", "キッチン", "カーポート"]
        assert [p["amount"] for p in props] == [120000, 350000, 0]

        clear_all = [b for b in at.button if b.key == "btn_props_clear_all"]
        clear_all[0].click().run()
        assert not at.exception, [str(e.value) for e in at.exception]
        props = at.session_state["proj_sample_furui_ai_proposals"]
        assert [p.get("apply") for p in props] == [False, False, False]
        assert at.session_state["proj_sample_furui_ai_proposals_epoch"] == 2
        assert _apply_button(at).label == "選択した0件を反映"
        assert _apply_button(at).disabled is True

    def test_voice_proposals_default_check_still_works(self, app_dir):
        """音声由来は高確度のみ初期チェック（既存仕様の維持）"""
        at = _boot(app_dir)
        at.session_state["proj_sample_furui_ai_proposals"] = [dict(p) for p in PROPOSALS]
        at.session_state["proj_sample_furui_ai_proposals_src"] = "音声変更"
        at.run()
        assert not at.exception, [str(e.value) for e in at.exception]
        assert _apply_button(at).label == "選択した1件を反映"


class TestPlanLabelInSidebar:
    def test_plan_selector_shows_plan_name_only(self, app_dir):
        at = _boot(app_dir)
        projects = dict(at.session_state["projects"])
        projects["sample_furui__plan2"] = {
            "name": "古居様邸（サンプル）",
            "file_name": "サンプルデータ",
            "created": "2026-10-26 10:00",
            "plan": "古居様　打合せ変更 10/26",
        }
        at.session_state["projects"] = projects
        at.run()
        assert not at.exception, [str(e.value) for e in at.exception]
        sel = [s for s in at.sidebar.selectbox if s.key == "project_selector"]
        assert sel, "案件（プラン）選択欄が見つからない"
        sb = sel[0]
        # AppTest の Selectbox.options は format_func 適用後の表示文字列
        labels = list(sb.options)
        assert labels == ["古居様邸（サンプル）", "古居様　打合せ変更 10/26"], labels


class TestFeedbackStatusFlowInApp:
    def test_three_stage_status_and_history(self, app_dir):
        import json
        fb_path = os.path.join(app_dir, "projects", "sample_furui_feedbacks.json")
        with open(fb_path, "w", encoding="utf-8") as f:
            json.dump([{"category": "操作性・画面について", "target": "", "body": "全選択が欲しい",
                        "priority": "通常", "name": "テスト", "email": "", "timestamp": "2026-08-04 17:25",
                        "property": "サンプル邸", "status": "未対応", "feedback_id": "fb-test-1",
                        "attachments": []}], f, ensure_ascii=False)
        at = _boot(app_dir)
        # 既定の「対応中」表示では、絞り込みは対応中の3ステータスだけ（対応完了はアーカイブ側）
        views = [r for r in at.radio if r.key == "fb_archive_view"]
        assert views and list(views[0].options) == ["対応中（1）", "アーカイブ・対応完了（0）", "全て"]
        radios = [r for r in at.radio if r.key == "fb_filter_active"]
        assert radios, "表示フィルターが無い"
        assert list(radios[0].options) == ["全て", "未対応", "高井対応済み", "差し戻し"]

        sel = [s for s in at.selectbox if s.key == "status_fb_sample_furui_0_sel"]
        assert sel, "ステータス変更の選択欄が無い"
        assert list(sel[0].options) == ["高井対応済み", "対応完了", "差し戻し"]

        # 差し戻しはコメント必須
        sel[0].select("差し戻し")
        [b for b in at.button if b.key == "status_fb_sample_furui_0"][0].click().run()
        assert not at.exception, [str(e.value) for e in at.exception]
        assert any("コメント" in e.value for e in at.error)
        with open(fb_path, encoding="utf-8") as f:
            assert json.load(f)[0]["status"] == "未対応"

        # 高井対応済み + 変更箇所 → 保存され履歴が残る（メールは未設定なので送らない）
        sel = [s for s in at.selectbox if s.key == "status_fb_sample_furui_0_sel"][0]
        sel.select("高井対応済み")
        [t for t in at.text_input if t.key == "status_fb_sample_furui_0_note"][0].input("全て選択／全て解除ボタンを追加")
        [b for b in at.button if b.key == "status_fb_sample_furui_0"][0].click().run()
        assert not at.exception, [str(e.value) for e in at.exception]
        with open(fb_path, encoding="utf-8") as f:
            saved = json.load(f)[0]
        assert saved["status"] == "高井対応済み"
        assert saved["status_note"] == "全て選択／全て解除ボタンを追加"
        assert saved["status_history"][-1]["status"] == "高井対応済み"
        assert saved["status_history"][-1]["by"] == "高井"
