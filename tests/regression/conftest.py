"""回帰テスト共通設定

- プロジェクト直下のモジュールを import できるようにする
- 全テストで本番 Supabase への接続を遮断する（Streamlit は cwd の .streamlit/secrets.toml を読むため、
  案件ルートで pytest を実行すると本番に繋がってしまう。ここで物理的に止める）
"""
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


@pytest.fixture(autouse=True)
def _block_cloud_everywhere(monkeypatch):
    import cloud_store

    monkeypatch.setattr(cloud_store, "_config", lambda: ("", ""))

    def _blocked(*_a, **_k):
        raise cloud_store.CloudStoreError("テスト中のクラウドアクセスは禁止")

    monkeypatch.setattr(cloud_store, "get", _blocked)
    monkeypatch.setattr(cloud_store, "put", _blocked)
    monkeypatch.setattr(cloud_store, "delete", _blocked)
    cloud_store.clear_app_secret_cache()
    yield
    cloud_store.clear_app_secret_cache()
