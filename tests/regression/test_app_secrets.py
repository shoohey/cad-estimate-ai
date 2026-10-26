"""AIキーの取得順: 環境変数 → Streamlit secrets → クラウドの app_secrets（許可キーのみ）"""
import ai_client
import cloud_store


def _cloud(monkeypatch, values, enabled=True, raises=False):
    monkeypatch.setattr(cloud_store, "enabled", lambda: enabled)
    calls = {"n": 0}

    def _get(key):
        calls["n"] += 1
        assert key == cloud_store.APP_SECRETS_KEY
        if raises:
            raise cloud_store.CloudStoreError("down")
        return values
    monkeypatch.setattr(cloud_store, "get", _get)
    cloud_store.clear_app_secret_cache()
    return calls


def test_env_or_streamlit_takes_precedence(monkeypatch):
    monkeypatch.setattr(ai_client, "_secret_from_env_or_streamlit", lambda k: "from-env")
    calls = _cloud(monkeypatch, {"ANTHROPIC_API_KEY": "from-cloud"})
    assert ai_client._get_secret("ANTHROPIC_API_KEY") == "from-env"
    assert calls["n"] == 0


def test_falls_back_to_cloud_and_caches(monkeypatch):
    monkeypatch.setattr(ai_client, "_secret_from_env_or_streamlit", lambda k: "")
    calls = _cloud(monkeypatch, {"ANTHROPIC_API_KEY": "from-cloud"})
    assert ai_client._get_secret("ANTHROPIC_API_KEY") == "from-cloud"
    assert ai_client._get_secret("ANTHROPIC_API_KEY") == "from-cloud"
    assert calls["n"] == 1                      # 2回目はキャッシュ
    assert ai_client.ai_capabilities()["text"] is True


def test_only_allowed_keys_are_read_from_cloud(monkeypatch):
    monkeypatch.setattr(ai_client, "_secret_from_env_or_streamlit", lambda k: "")
    calls = _cloud(monkeypatch, {"SUPABASE_SERVICE_KEY": "x", "RANDOM": "y"})
    assert ai_client._get_secret("SUPABASE_SERVICE_KEY") == ""
    assert ai_client._get_secret("RANDOM") == ""
    assert calls["n"] == 0


def test_cloud_disabled_or_failing_means_unset(monkeypatch):
    monkeypatch.setattr(ai_client, "_secret_from_env_or_streamlit", lambda k: "")
    _cloud(monkeypatch, {"ANTHROPIC_API_KEY": "k"}, enabled=False)
    assert ai_client._get_secret("ANTHROPIC_API_KEY") == ""
    _cloud(monkeypatch, None, raises=True)
    assert ai_client._get_secret("ANTHROPIC_API_KEY") == ""
    assert ai_client.ai_capabilities()["text"] is False


def test_non_dict_value_is_ignored(monkeypatch):
    monkeypatch.setattr(ai_client, "_secret_from_env_or_streamlit", lambda k: "")
    _cloud(monkeypatch, ["not", "a", "dict"])
    assert ai_client._get_secret("ANTHROPIC_API_KEY") == ""
