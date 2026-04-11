"""
フィードバック送信モジュール (CAD→概算見積もり生成AI / 日本住建株式会社)

Streamlitアプリからのユーザーフィードバックを、以下の3チャネルに配信する:
  1. Slack chat.postMessage API (Bot Token方式) - 弊社Slackチャンネルへ即時通知
  2. Email via Resend API (RESEND_API_KEY + FEEDBACK_TO_EMAIL が設定されている場合)
  3. ローカル JSON ログ (feedbacks.json) - 常時フォールバックとして追記

依存: 標準ライブラリ + requests のみ
"""

from __future__ import annotations

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

# --- 定数 ---------------------------------------------------------------

FEEDBACK_LOG_PATH = Path(__file__).resolve().parent / "feedbacks.json"
RESEND_ENDPOINT = "https://api.resend.com/emails"
RESEND_FROM_ADDRESS = "onboarding@resend.dev"
REQUEST_TIMEOUT_SEC = 10

# Slack Bot Token は環境変数または Streamlit secrets から読み込む（ハードコード禁止）。
SLACK_API_ENDPOINT = "https://slack.com/api/chat.postMessage"


# --- 内部ヘルパー -------------------------------------------------------

def _now_iso() -> str:
    """現在時刻を ISO8601 (UTC) 形式で返す。"""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _append_local_log(record: dict[str, Any]) -> None:
    """ローカル JSON ログにフィードバックを追記する。

    ファイルが存在しない・壊れている場合は新規配列として再作成する。
    """
    data: list[dict[str, Any]] = []
    if FEEDBACK_LOG_PATH.exists():
        try:
            with FEEDBACK_LOG_PATH.open("r", encoding="utf-8") as fp:
                loaded = json.load(fp)
                if isinstance(loaded, list):
                    data = loaded
        except (json.JSONDecodeError, OSError):
            data = []

    data.append(record)

    with FEEDBACK_LOG_PATH.open("w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False, indent=2)


def _build_slack_blocks(record: dict[str, Any]) -> list[dict[str, Any]]:
    """Slack block kit 形式のブロックを生成する。"""
    return [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": "📩 新しいフィードバック (日本住建 見積AI)",
                "emoji": True,
            },
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*ID:*\n`{record['feedback_id'][:8]}`"},
                {"type": "mrkdwn", "text": f"*カテゴリ:*\n{record['category']}"},
                {"type": "mrkdwn", "text": f"*送信者:*\n{record['name']}"},
                {"type": "mrkdwn", "text": f"*メール:*\n{record['email'] or '(未入力)'}"},
                {"type": "mrkdwn", "text": f"*物件:*\n{record.get('property', '(未指定)')}"},
                {"type": "mrkdwn", "text": f"*日時:*\n{record['timestamp']}"},
            ],
        },
        {"type": "divider"},
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f"*メッセージ:*\n>>> {record['message']}",
            },
        },
    ]


def _build_email_html(record: dict[str, Any]) -> str:
    """Resend 用のフィードバックメール HTML 本文を組み立てる。"""
    attachments = record.get("attachments") or []
    attachments_html = (
        "<ul>" + "".join(f"<li>{str(a)}</li>" for a in attachments) + "</ul>"
        if attachments
        else "<p style='color:#4a5568;'>(なし)</p>"
    )
    # メッセージ本文の簡易エスケープ + 改行保持
    safe_message = (
        record["message"]
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br>")
    )
    return f"""
<!doctype html>
<html lang="ja">
<head><meta charset="utf-8"></head>
<body style="font-family: 'Noto Sans JP', sans-serif; background:#f5f7fa; padding:24px; color:#1a1a2e;">
  <div style="max-width:640px; margin:0 auto; background:#ffffff; border:1px solid #e2e8f0; border-radius:8px; box-shadow:0 4px 16px rgba(0,0,0,0.06); padding:32px;">
    <p style="letter-spacing:.08em; color:#1e3a5f; font-size:12px; margin:0 0 8px;">NIHON JUKEN / ESTIMATE AI</p>
    <h1 style="color:#1e3a5f; font-size:24px; margin:0 0 16px;">新しいフィードバックが届きました</h1>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
      <tr><td style="padding:8px 0; color:#4a5568; width:120px;">ID</td><td style="padding:8px 0;"><code>{record['feedback_id']}</code></td></tr>
      <tr><td style="padding:8px 0; color:#4a5568;">日時</td><td style="padding:8px 0;">{record['timestamp']}</td></tr>
      <tr><td style="padding:8px 0; color:#4a5568;">カテゴリ</td><td style="padding:8px 0;">{record['category']}</td></tr>
      <tr><td style="padding:8px 0; color:#4a5568;">送信者</td><td style="padding:8px 0;">{record['name']}</td></tr>
      <tr><td style="padding:8px 0; color:#4a5568;">メール</td><td style="padding:8px 0;">{record['email'] or '(未入力)'}</td></tr>
    </table>
    <h2 style="color:#1e3a5f; font-size:16px; margin:24px 0 8px;">メッセージ</h2>
    <div style="background:#f5f7fa; border:1px solid #e2e8f0; border-radius:8px; padding:16px; line-height:1.9;">{safe_message}</div>
    <h2 style="color:#1e3a5f; font-size:16px; margin:24px 0 8px;">添付ファイル</h2>
    {attachments_html}
    <p style="color:#4a5568; font-size:12px; margin-top:32px;">このメールは自動送信されています。</p>
  </div>
</body>
</html>
""".strip()


def _get_secret(key: str) -> str:
    """環境変数 → Streamlit secrets の順で値を取得する。どちらにも無ければ空文字を返す。"""
    val = os.environ.get(key)
    if val:
        return val
    try:
        import streamlit as st  # type: ignore

        # st.secrets はファイル未配置時に StreamlitSecretNotFoundError を投げる
        return str(st.secrets.get(key, "")) if hasattr(st, "secrets") else ""
    except Exception:  # noqa: BLE001
        return ""


def _send_slack(record: dict[str, Any], errors: list[str]) -> bool:
    """Slack chat.postMessage API で通知を送信する。成功すると True を返す。"""
    token = _get_secret("SLACK_BOT_TOKEN")
    channel = _get_secret("SLACK_CHANNEL_ID")
    if not token or not channel:
        return False
    try:
        resp = requests.post(
            SLACK_API_ENDPOINT,
            json={
                "channel": channel,
                "text": f"新しいフィードバック ({record['category']})",
                "blocks": _build_slack_blocks(record),
            },
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json; charset=utf-8",
            },
            timeout=REQUEST_TIMEOUT_SEC,
        )
        if resp.status_code >= 300:
            errors.append(f"slack: HTTP {resp.status_code} - {resp.text[:200]}")
            return False
        body = resp.json()
        if not body.get("ok"):
            errors.append(f"slack: API error - {body.get('error', 'unknown')}")
            return False
        return True
    except Exception as exc:  # noqa: BLE001 - 外部送信はまとめて捕捉
        errors.append(f"slack: {type(exc).__name__}: {exc}")
        return False


def _send_email(record: dict[str, Any], errors: list[str]) -> bool:
    """Resend API 経由でフィードバックメールを送信する。"""
    api_key = _get_secret("RESEND_API_KEY")
    to_email = _get_secret("FEEDBACK_TO_EMAIL")
    if not api_key or not to_email:
        return False
    try:
        payload = {
            "from": RESEND_FROM_ADDRESS,
            "to": [to_email],
            "subject": f"【日本住建 見積AI】新しいフィードバック - {record['category']}",
            "html": _build_email_html(record),
        }
        resp = requests.post(
            RESEND_ENDPOINT,
            json=payload,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            timeout=REQUEST_TIMEOUT_SEC,
        )
        if resp.status_code >= 300:
            errors.append(f"email: HTTP {resp.status_code} - {resp.text[:200]}")
            return False
        return True
    except Exception as exc:  # noqa: BLE001
        errors.append(f"email: {type(exc).__name__}: {exc}")
        return False


# --- 公開 API -----------------------------------------------------------

def submit_feedback(
    name: str,
    email: str,
    category: str,
    message: str,
    attachments: list | None = None,
    property_name: str = "",
    priority: str = "通常",
    target: str = "",
) -> dict:
    """フィードバックを受け取り、設定済みの全チャネルへ配信する。

    Args:
        name: 送信者の表示名。
        email: 送信者の返信先メールアドレス (空文字可)。
        category: フィードバックのカテゴリ (例: "バグ報告", "機能要望")。
        message: 本文テキスト。
        attachments: 添付ファイル名またはパスのリスト。メタデータのみ扱う。

    Returns:
        dict: {
            "success": bool,           # いずれかのチャネル送信に成功すれば True
            "channels_sent": list[str],# 成功したチャネル名
            "errors": list[str],       # 発生したエラーメッセージ
            "feedback_id": str,        # 発行した UUID
        }
    """
    feedback_id = str(uuid.uuid4())
    record: dict[str, Any] = {
        "feedback_id": feedback_id,
        "timestamp": _now_iso(),
        "name": name or "(名無し)",
        "email": email or "",
        "category": category or "未分類",
        "message": message or "",
        "attachments": attachments or [],
        "property": property_name or "",
        "priority": priority or "通常",
        "target": target or "",
    }

    errors: list[str] = []
    channels_sent: list[str] = []

    # 1) ローカル JSON ログ (常時)
    try:
        _append_local_log(record)
        channels_sent.append("local_json")
    except Exception as exc:  # noqa: BLE001
        errors.append(f"local_json: {type(exc).__name__}: {exc}")

    # 2) Slack
    if _send_slack(record, errors):
        channels_sent.append("slack")

    # 3) Email (Resend)
    if _send_email(record, errors):
        channels_sent.append("email")

    return {
        "success": len(channels_sent) > 0,
        "channels_sent": channels_sent,
        "errors": errors,
        "feedback_id": feedback_id,
    }


# --- CLI エントリポイント -----------------------------------------------

if __name__ == "__main__":
    print("=== feedback_handler.py 動作確認 ===")
    result = submit_feedback(
        name="テスト太郎",
        email="test@example.com",
        category="機能要望",
        message="CADアップロード後の概算見積もり表示をもっと見やすくしてほしい。\n特に数量と単価の内訳を展開できると嬉しいです。",
        attachments=["sample.dxf", "memo.txt"],
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nローカルログ: {FEEDBACK_LOG_PATH}")
