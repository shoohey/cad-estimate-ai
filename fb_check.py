"""お客様フィードバックの棚卸し（本番 Supabase を読むだけ。書き込みはしない）

使い方: python3 fb_check.py            → 未対応・差し戻しの一覧
        python3 fb_check.py --all      → 全件（ステータス別の件数つき）
.env.local（SUPABASE_URL / SUPABASE_SERVICE_KEY）を読む。
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
                os.environ.setdefault(k.strip(), v.strip())


def main():
    _load_env()
    url = os.environ.get("SUPABASE_URL", "").rstrip("/")
    key = os.environ.get("SUPABASE_SERVICE_KEY", "")
    if not url or not key:
        print("SUPABASE_URL / SUPABASE_SERVICE_KEY が見つかりません（.env.local）")
        sys.exit(1)
    show_all = "--all" in sys.argv
    r = requests.get(f"{url}/rest/v1/app_store",
                     params={"select": "key,value", "key": "ilike.*:feedbacks"},
                     headers={"apikey": key, "Authorization": f"Bearer {key}"}, timeout=20)
    r.raise_for_status()
    rows = []
    for row in r.json():
        pid = row["key"][len("proj:"):-len(":feedbacks")]
        for fb in row["value"] or []:
            status = fb.get("status") or "未対応"
            if status == "対応済み":
                status = "対応完了"
            rows.append((fb.get("timestamp", ""), status, fb.get("name", ""), fb.get("property", "") or pid,
                         (fb.get("body") or fb.get("message") or "").replace("\n", " "), fb.get("status_note", "")))
    rows.sort()
    counts = {}
    for _, st, *_ in rows:
        counts[st] = counts.get(st, 0) + 1
    print("件数:", " / ".join(f"{k}: {v}" for k, v in counts.items()) or "0")
    todo = [x for x in rows if show_all or x[1] in ("未対応", "差し戻し")]
    if not todo:
        print("対応待ちはありません。")
        return
    print()
    for ts, st, name, prop, body, note in todo:
        print(f"[{st}] {ts} {name}様 / {prop}")
        print(f"   {body[:140]}")
        if note:
            print(f"   メモ: {note[:100]}")
    if not show_all:
        print("\n→ 修正したらアプリの「フィードバック」タブで「高井対応済み」＋変更箇所を入力してください。")


if __name__ == "__main__":
    main()
