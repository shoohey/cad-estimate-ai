"""
フィードバック送信モジュール (CAD→概算見積もり生成AI / 日本住建株式会社)

Streamlitアプリからのユーザーフィードバックを、以下の3チャネルに配信する:
  1. Slack chat.postMessage API (Bot Token方式) - 弊社Slackチャンネルへ即時通知
  2. Email via Resend API (RESEND_API_KEY + FEEDBACK_TO_EMAIL が設定されている場合)
  3. ローカル JSON ログ (feedbacks.json) - 常時フォールバックとして追記

依存: 標準ライブラリ + requests のみ
"""

from __future__ import annotations

import base64
import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

# --- 定数 ---------------------------------------------------------------

# 送信ログの保存先。旧グローバル保存ファイル(feedbacks.json)とは名前を分ける。
# 同名にすると、アプリ起動時の「旧feedbacks.jsonを案件別へ移行する処理」が
# 送信のたびに書かれたこのログを取り込み、同じ内容が二重に登録されてしまう。
FEEDBACK_LOG_PATH = Path(__file__).resolve().parent / "feedback_sendlog.json"
RESEND_ENDPOINT = "https://api.resend.com/emails"
RESEND_FROM_ADDRESS = "onboarding@resend.dev"
REQUEST_TIMEOUT_SEC = 10
ATTACHMENT_UPLOAD_TIMEOUT_SEC = 60   # 添付ありの送信は時間がかかるため別枠

# 添付ファイルの上限（Resendのメール全体上限40MBに対し、base64化で約1.37倍に
# 膨らむことを見込んで保守的に設定する）
MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024        # 1ファイルあたり
MAX_ATTACHMENT_TOTAL_BYTES = 20 * 1024 * 1024  # 合計

# Slack Bot Token は環境変数または Streamlit secrets から読み込む（ハードコード禁止）。
SLACK_API_ENDPOINT = "https://slack.com/api/chat.postMessage"
SLACK_UPLOAD_URL_ENDPOINT = "https://slack.com/api/files.getUploadURLExternal"
SLACK_UPLOAD_COMPLETE_ENDPOINT = "https://slack.com/api/files.completeUploadExternal"


# --- 内部ヘルパー -------------------------------------------------------

def _now_iso() -> str:
    """現在時刻を ISO8601 (UTC) 形式で返す。"""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _normalize_attachments(attachments: list | None) -> list[dict[str, Any]]:
    """添付指定を {'name','data','size','error'} の形に揃える。

    受け付ける形式:
      - {'name': str, 'data': bytes}      … 実体つき（メール/Slackへ添付する）
      - str（ファイル名またはパス）        … 実体が読めればbytesを読み込む
    サイズ上限を超えたものは data を落とし、error に理由を入れて記録だけ残す。
    """
    normalized: list[dict[str, Any]] = []
    total = 0
    for att in attachments or []:
        name, data = "", None
        if isinstance(att, dict):
            name = str(att.get("name") or att.get("filename") or "")
            data = att.get("data")
        else:
            name = str(att)
            path = Path(name)
            if path.is_file():
                try:
                    data = path.read_bytes()
                except OSError:
                    data = None
                name = path.name
        if not name:
            continue
        entry: dict[str, Any] = {"name": name, "size": len(data) if data else 0}
        if data is None:
            entry["error"] = "ファイルの実体を取得できませんでした（ファイル名のみ記録）"
        elif len(data) > MAX_ATTACHMENT_BYTES:
            entry["error"] = (f"1ファイル {MAX_ATTACHMENT_BYTES // (1024 * 1024)}MB の上限超過"
                              f"（{len(data) / (1024 * 1024):.1f}MB）")
        elif total + len(data) > MAX_ATTACHMENT_TOTAL_BYTES:
            entry["error"] = f"合計 {MAX_ATTACHMENT_TOTAL_BYTES // (1024 * 1024)}MB の上限超過"
        else:
            entry["data"] = data
            total += len(data)
        normalized.append(entry)
    return normalized


def _attachment_labels(record: dict[str, Any]) -> list[str]:
    """添付ファイルの一覧表示用ラベル（名前・サイズ・エラー）"""
    labels = []
    for att in record.get("attachments", []):
        if isinstance(att, dict):
            size_kb = (att.get("size") or 0) / 1024
            label = f"{att.get('name', '')}（{size_kb:,.0f}KB）"
            if att.get("error"):
                label += f" ※未送信: {att['error']}"
        else:
            label = str(att)
        labels.append(label)
    return labels


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

    # 添付の実体（bytes）はJSONに載せられないため、メタ情報だけ残す
    loggable = dict(record)
    loggable["attachments"] = [
        {k: v for k, v in att.items() if k != "data"} if isinstance(att, dict) else att
        for att in record.get("attachments", [])
    ]
    data.append(loggable)

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
    ] + ([
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": "*添付ファイル:*\n" + "\n".join(
                    f"• {label}" for label in _attachment_labels(record)),
            },
        },
    ] if record.get("attachments") else [])


def _build_email_html(record: dict[str, Any]) -> str:
    """Resend 用のフィードバックメール HTML 本文を組み立てる。"""
    labels = _attachment_labels(record)
    attachments_html = (
        "<ul>" + "".join(f"<li>{a}</li>" for a in labels) + "</ul>"
        if labels
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
        _upload_slack_files(record, token, channel, errors)
        return True
    except Exception as exc:  # noqa: BLE001 - 外部送信はまとめて捕捉
        errors.append(f"slack: {type(exc).__name__}: {exc}")
        return False


def _upload_slack_files(record: dict[str, Any], token: str, channel: str,
                        errors: list[str]) -> None:
    """添付ファイルをSlackチャンネルへアップロードする。

    files.getUploadURLExternal で発行されたURLへPUT相当のPOSTを行い、
    completeUploadExternal でチャンネルへ紐付ける（Slackの現行手順）。
    失敗しても通知本文の送信自体は成功扱いのままとし、errors に記録する
    （添付が上がらないことよりも通知が届かないことのほうが問題のため）。
    """
    files = [a for a in record.get("attachments", [])
             if isinstance(a, dict) and a.get("data")]
    if not files:
        return
    auth = {"Authorization": f"Bearer {token}"}
    uploaded = []
    for att in files:
        try:
            got = requests.get(
                SLACK_UPLOAD_URL_ENDPOINT,
                params={"filename": att["name"], "length": len(att["data"])},
                headers=auth, timeout=REQUEST_TIMEOUT_SEC,
            ).json()
            if not got.get("ok"):
                errors.append(f"slack file: {got.get('error', 'unknown')}")
                continue
            put = requests.post(
                got["upload_url"],
                files={"file": (att["name"], att["data"])},
                timeout=ATTACHMENT_UPLOAD_TIMEOUT_SEC,
            )
            if put.status_code >= 300:
                errors.append(f"slack file: upload HTTP {put.status_code}")
                continue
            uploaded.append({"id": got["file_id"], "title": att["name"]})
        except Exception as exc:  # noqa: BLE001
            errors.append(f"slack file: {type(exc).__name__}: {exc}")
    if not uploaded:
        return
    try:
        done = requests.post(
            SLACK_UPLOAD_COMPLETE_ENDPOINT,
            json={"files": uploaded, "channel_id": channel,
                  "initial_comment": f"フィードバック `{record['feedback_id'][:8]}` の添付ファイル"},
            headers={**auth, "Content-Type": "application/json; charset=utf-8"},
            timeout=REQUEST_TIMEOUT_SEC,
        ).json()
        if not done.get("ok"):
            errors.append(f"slack file: complete - {done.get('error', 'unknown')}")
    except Exception as exc:  # noqa: BLE001
        errors.append(f"slack file: {type(exc).__name__}: {exc}")


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
        # 添付ファイル（実体があるもののみ。base64でRESENDに渡す）
        mail_attachments = [
            {"filename": att["name"],
             "content": base64.b64encode(att["data"]).decode("ascii")}
            for att in record.get("attachments", [])
            if isinstance(att, dict) and att.get("data")
        ]
        if mail_attachments:
            payload["attachments"] = mail_attachments
        resp = requests.post(
            RESEND_ENDPOINT,
            json=payload,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            timeout=(ATTACHMENT_UPLOAD_TIMEOUT_SEC if mail_attachments
                     else REQUEST_TIMEOUT_SEC),
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
        attachments: 添付ファイルのリスト。{'name': str, 'data': bytes} を渡すと
            メールとSlackへ実ファイルとして添付する。文字列を渡した場合は
            パスとして読み込みを試み、読めなければファイル名のみ記録する。

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
        "attachments": _normalize_attachments(attachments),
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
