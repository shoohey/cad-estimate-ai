"""Resend のメール設定を確認する（ローカル実行用）

使い方:
  python3 resend_check.py                  → ドメインの検証状態と設定の過不足を表示
  python3 resend_check.py --send 宛先      → 本番と同じ処理で「高井対応済み」通知のテストメールを送る
.env.local に RESEND_API_KEY（必要なら RESEND_FROM）を書いておく。キーは画面に出さない。
"""
import os
import sys

import requests


def _load_env():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env.local")
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            if "=" in line and not line.startswith("#"):
                k, v = line.rstrip("\n").split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"'))


def main():
    _load_env()
    key = os.environ.get("RESEND_API_KEY", "")
    if not key:
        print("RESEND_API_KEY が .env.local にありません。Resend の API Keys で作成して追記してください。")
        sys.exit(1)
    r = requests.get("https://api.resend.com/domains", headers={"Authorization": f"Bearer {key}"}, timeout=20)
    if r.status_code >= 300:
        print(f"Resend API エラー: HTTP {r.status_code} {r.text[:120]}")
        sys.exit(1)
    domains = r.json().get("data", [])
    print("ドメイン:", ", ".join(f"{d['name']}({d['status']})" for d in domains) or "未登録")
    verified = [d["name"] for d in domains if d.get("status") == "verified"]
    sender = os.environ.get("RESEND_FROM", "")
    if not verified:
        print("→ 検証済みドメインがありません。検証前はResendの登録メールアドレス宛てにしか送れません。")
    elif not sender:
        print(f"→ RESEND_FROM を設定してください（例: noreply@{verified[0]}）")
    elif sender.split("@")[-1] not in verified:
        print(f"→ RESEND_FROM のドメインが未検証です: {sender}")
    else:
        print(f"→ 送信元 {sender} は使用可能です")
    for k in ("FEEDBACK_NOTIFY_TO", "FEEDBACK_TO_EMAIL"):
        print(f"{k}: {'設定あり' if os.environ.get(k) else '未設定（Streamlit Secrets 側で設定）'}")

    if "--send" in sys.argv:
        to = sys.argv[sys.argv.index("--send") + 1]
        os.environ["FEEDBACK_NOTIFY_TO"] = to
        import feedback_handler
        record = {"body": "（テスト）全選択ボタンが欲しい", "category": "操作性・画面について",
                  "timestamp": "テスト送信", "email": "", "property": "サンプル邸"}
        ok, msg = feedback_handler.notify_feedback_status(
            record, "高井対応済み", "テスト: 変更案の表の下に「全て選択」「全て解除」を追加", "サンプル邸")
        print("テスト送信:", "成功" if ok else "失敗", "-", msg)


if __name__ == "__main__":
    main()
