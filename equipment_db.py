"""
住宅設備マスターモジュール (CAD→概算見積もり生成AI / 日本住建株式会社)

住宅設備機器（キッチン・システムバス・トイレ等）の選択肢マスターを管理し、
案件ごとの設備選択を単価マスターに反映する。

- マスターの保存/読込は app.py 側（クラウド永続化のヘルパー）が担当し、
  本モジュールは純粋なデータ操作のみを行う。
- マスター構造（JSONシリアライズ可能なプリミティブのみ）:
  {
    "140001": {
      "label": "キッチン",
      "default_id": "std",
      "options": [
        {"id": "std", "name": "...", "summary": "...",
         "estimate_price": 1234, "order_price": 1000},
        ...
      ]
    },
    ...
  }
"""

from __future__ import annotations

import copy

# 設備マスターの対象となる単価マスターコード（estimate_calculator の cat14 で使用）
EQUIPMENT_CODES = {
    140001: "キッチン",
    140004: "カップボード",
    140005: "キッチン水栓",
    140006: "システムバス 1620",
    140007: "システムバス 1616",
    140008: "トイレ（1台目）",
    140009: "トイレ（2台目以降）",
    140011: "洗面化粧台",
    140015: "洗濯パン",
    140016: "エコキュート 460L",
    140017: "エコキュート 370L",
}


def seed_master_from_prices(prices: dict) -> dict:
    """単価マスター（tankamaster）から設備マスターの初期値を生成する"""
    master = {}
    for code, label in EQUIPMENT_CODES.items():
        p = prices.get(code)
        if not p:
            continue
        master[str(code)] = {
            "label": label,
            "default_id": "std",
            "options": [{
                "id": "std",
                "name": p["name"],
                "summary": p.get("summary", ""),
                "estimate_price": int(p.get("estimate_price", 0)),
                "order_price": int(p.get("order_price", 0)),
            }],
        }
    return master


def get_option(master: dict, code, option_id: str) -> dict | None:
    """指定コード・選択肢IDのオプションを返す（無ければNone）"""
    entry = master.get(str(code))
    if not entry:
        return None
    for opt in entry.get("options", []):
        if opt.get("id") == option_id:
            return opt
    return None


def effective_selection(master: dict, project_selection: dict | None) -> dict:
    """案件の設備選択を確定する。

    案件側で明示選択されたものはそれを、無いものはマスターの標準
    （default_id）を採用する。返り値は {code(str): option_id}。
    """
    sel = {}
    for code, entry in master.items():
        default_id = entry.get("default_id", "std")
        chosen = (project_selection or {}).get(code, default_id)
        if get_option(master, code, chosen) is None:
            chosen = default_id
        sel[code] = chosen
    return sel


def refresh_std_options(master: dict, prices: dict):
    """各コードの標準オプション('std')を単価マスターの現行値で更新する。

    'std' は tankamaster（単価マスターCSV）連動の特別なオプションで、
    今後の単価マスター更新が見積へ反映され続けるよう毎回同期する。
    また全選択肢が削除されても std は必ず復元される。
    """
    for code_str, entry in master.items():
        try:
            p = prices.get(int(code_str))
        except (TypeError, ValueError):
            p = None
        if not p:
            continue
        std_data = {
            "id": "std",
            "name": p["name"],
            "summary": p.get("summary", ""),
            "estimate_price": int(p.get("estimate_price", 0)),
            "order_price": int(p.get("order_price", 0)),
        }
        options = entry.setdefault("options", [])
        std = next((o for o in options if o.get("id") == "std"), None)
        if std:
            std.update(std_data)
        else:
            options.insert(0, std_data)
        if not entry.get("default_id"):
            entry["default_id"] = "std"


def apply_equipment(prices: dict, master: dict, project_selection: dict | None) -> dict:
    """設備選択を単価マスターに反映した複製を返す。

    選択されたオプションの名称・摘要・見積単価・発注単価で該当コードを
    上書きする。標準（'std'）選択のコードは単価マスターの現行値を
    そのまま使う（上書きしない）。元の prices は変更しない。
    """
    result = copy.deepcopy(prices)
    sel = effective_selection(master, project_selection)
    for code_str, option_id in sel.items():
        if option_id == "std":
            continue
        opt = get_option(master, code_str, option_id)
        if not opt:
            continue
        code = int(code_str)
        if code not in result:
            continue
        result[code] = dict(result[code])
        result[code]["name"] = opt.get("name", result[code]["name"])
        result[code]["summary"] = opt.get("summary", result[code].get("summary", ""))
        result[code]["estimate_price"] = int(opt.get("estimate_price", 0))
        result[code]["order_price"] = int(opt.get("order_price", 0))
    return result


def selection_changes(master: dict, before: dict | None, after: dict | None) -> list:
    """設備選択の変更点を日本語の説明リストで返す（変更履歴用）"""
    changes = []
    eff_before = effective_selection(master, before)
    eff_after = effective_selection(master, after)
    for code, after_id in eff_after.items():
        before_id = eff_before.get(code)
        if before_id == after_id:
            continue
        entry = master.get(code, {})
        opt_b = get_option(master, code, before_id) or {}
        opt_a = get_option(master, code, after_id) or {}
        changes.append(
            f"{entry.get('label', code)}: "
            f"{opt_b.get('name', before_id)} → {opt_a.get('name', after_id)}"
        )
    return changes
