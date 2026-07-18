"""
クラウド永続化モジュール (CAD→概算見積もり生成AI / 日本住建株式会社)

Streamlit Community Cloud のファイルシステムはエフェメラル（再起動・再デプロイで消える）
ため、案件データ・CADファイル・フィードバックを Supabase の app_store テーブル
(key text PK / value jsonb) に保存する。

設定（環境変数 → Streamlit secrets の順で読む）:
  SUPABASE_URL          https://<ref>.supabase.co
  SUPABASE_SERVICE_KEY  service_role キー（サーバーサイド専用。クライアントに渡さない）

どちらかが未設定の場合は enabled() が False となり、呼び出し側（app.py）は
従来どおりローカルファイルのみで動作する。

依存: requests のみ（PostgREST を直接叩く）
"""

from __future__ import annotations

import json
import os
from typing import Any

import requests

TABLE = "app_store"
REQUEST_TIMEOUT_SEC = 15


class CloudStoreError(Exception):
    """Supabase への読み書きに失敗したことを表す。"""


def _get_secret(key: str) -> str:
    """環境変数 → Streamlit secrets の順で値を取得する。どちらにも無ければ空文字。"""
    val = os.environ.get(key)
    if val:
        return val
    try:
        import streamlit as st  # type: ignore

        return str(st.secrets.get(key, "")) if hasattr(st, "secrets") else ""
    except Exception:  # noqa: BLE001 - secrets 未配置時は無効扱い
        return ""


def _config() -> tuple[str, str]:
    return _get_secret("SUPABASE_URL").rstrip("/"), _get_secret("SUPABASE_SERVICE_KEY")


def enabled() -> bool:
    """Supabase 永続化が構成済みかどうか。"""
    url, key = _config()
    return bool(url and key)


def _endpoint() -> str:
    url, _ = _config()
    return f"{url}/rest/v1/{TABLE}"


def _headers(extra: dict | None = None) -> dict:
    _, key = _config()
    h = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if extra:
        h.update(extra)
    return h


def get(key: str) -> Any | None:
    """key の value を返す。行が無ければ None。接続失敗は CloudStoreError。"""
    try:
        resp = requests.get(
            _endpoint(),
            params={"key": f"eq.{key}", "select": "value"},
            headers=_headers(),
            timeout=REQUEST_TIMEOUT_SEC,
        )
        if resp.status_code >= 300:
            raise CloudStoreError(f"GET {key}: HTTP {resp.status_code} - {resp.text[:200]}")
        rows = resp.json()
        return rows[0]["value"] if rows else None
    except CloudStoreError:
        raise
    except Exception as exc:  # noqa: BLE001 - ネットワーク系をまとめて変換
        raise CloudStoreError(f"GET {key}: {type(exc).__name__}: {exc}") from exc


def put(key: str, value: Any) -> None:
    """key に value を upsert する。失敗は CloudStoreError。"""
    try:
        resp = requests.post(
            _endpoint(),
            params={"on_conflict": "key"},
            data=json.dumps(
                [{"key": key, "value": value}], ensure_ascii=False
            ).encode("utf-8"),
            headers=_headers({"Prefer": "resolution=merge-duplicates"}),
            timeout=REQUEST_TIMEOUT_SEC,
        )
        if resp.status_code >= 300:
            raise CloudStoreError(f"PUT {key}: HTTP {resp.status_code} - {resp.text[:200]}")
    except CloudStoreError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise CloudStoreError(f"PUT {key}: {type(exc).__name__}: {exc}") from exc


def delete(key: str) -> None:
    """key の行を削除する（存在しなくてもエラーにしない）。"""
    try:
        resp = requests.delete(
            _endpoint(),
            params={"key": f"eq.{key}"},
            headers=_headers(),
            timeout=REQUEST_TIMEOUT_SEC,
        )
        if resp.status_code >= 300:
            raise CloudStoreError(f"DELETE {key}: HTTP {resp.status_code} - {resp.text[:200]}")
    except CloudStoreError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise CloudStoreError(f"DELETE {key}: {type(exc).__name__}: {exc}") from exc
