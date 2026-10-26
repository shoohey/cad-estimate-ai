"""実績比較の突合率を8物件で計測する（読み取り専用・ローカル）

使い方: python3 match_rate_check.py [--dump out.json]
  各物件について、実績見積の明細のうち AI明細と対応付いた件数・金額の割合と、
  対応付かなかった明細（AIのみ／実績のみ）を出す。突合ロジック改善の前後比較に使う。
"""
import json
import os
import sys

import excel_diff
from calibration_check import iter_pairs, DEFAULT_SAFETY_RATE
from cad_parser import load_cad_file
from estimate_calculator import calculate_estimate, load_unit_prices


def measure() -> dict:
    prices = load_unit_prices("tankamaster_updated.csv")
    report = {"properties": [], "unmatched": {}}
    # ローカルに「同じ商品」辞書があれば反映して測る（本番の辞書は projects/item_alias_library.json に同期される）
    aliases = set()
    alias_path = os.path.join("projects", "item_alias_library.json")
    if os.path.exists(alias_path):
        from ui_helpers import alias_pairs_from_library
        aliases = alias_pairs_from_library(json.load(open(alias_path, encoding="utf-8")))
        report["alias_entries"] = len(aliases)
    tot_ac_n = tot_ac_matched_n = 0
    tot_ac_amt = tot_ac_matched_amt = 0
    for pair in iter_pairs():
        pid, cad_file, xlsx = pair[0], pair[2], pair[3]
        cad = load_cad_file(cad_file)
        with open(xlsx, "rb") as f:
            actual = excel_diff.parse_actual_estimate(f.read())
        solar = any("太陽光" in c["name"] for c in actual["categories"])
        est = calculate_estimate(cad, prices, include_solar=solar, safety_rate=DEFAULT_SAFETY_RATE)
        diff = excel_diff.diff_estimates(est, actual, aliases)
        ac_n = ac_matched_n = ac_amt = ac_matched_amt = exact_n = 0
        unmatched = {}
        for cat, d in diff["item_diffs"].items():
            m, ai_only, ac_only = d["matched"], d["ai_only"], d["actual_only"]
            ac_n += len(m) + len(ac_only)
            ac_matched_n += len(m)
            exact_n += sum(1 for x in m if x["diff"] == 0)
            ac_amt += sum(x["actual_amount"] for x in m) + sum(x["amount"] for x in ac_only)
            ac_matched_amt += sum(x["actual_amount"] for x in m)
            if ai_only or ac_only:
                unmatched[cat] = {
                    "ai_only": [f'{x["name"]}｜{x.get("summary","")}｜{x["amount"]:,}' for x in ai_only],
                    "actual_only": [f'{x["name"]}｜{x.get("summary","")}｜{x["amount"]:,}' for x in ac_only],
                }
        report["properties"].append({
            "pid": pid, "actual_items": ac_n, "matched": ac_matched_n, "exact": exact_n,
            "match_rate_n": round(ac_matched_n / ac_n * 100, 1) if ac_n else 0,
            "match_rate_amt": round(ac_matched_amt / ac_amt * 100, 1) if ac_amt else 0,
        })
        report["unmatched"][pid] = unmatched
        tot_ac_n += ac_n; tot_ac_matched_n += ac_matched_n
        tot_ac_amt += ac_amt; tot_ac_matched_amt += ac_matched_amt
    report["total"] = {
        "actual_items": tot_ac_n, "matched": tot_ac_matched_n,
        "match_rate_n": round(tot_ac_matched_n / tot_ac_n * 100, 1) if tot_ac_n else 0,
        "match_rate_amt": round(tot_ac_matched_amt / tot_ac_amt * 100, 1) if tot_ac_amt else 0,
    }
    return report


def main():
    rep = measure()
    for p in rep["properties"]:
        print(f'{p["pid"][:28]:28s} 実績明細 {p["actual_items"]:3d} 件 / 突合 {p["matched"]:3d}（完全一致 {p["exact"]:2d}）'
              f' → 件数 {p["match_rate_n"]:5.1f}% / 金額 {p["match_rate_amt"]:5.1f}%')
    t = rep["total"]
    print(f'合計: 実績明細 {t["actual_items"]} 件 / 突合 {t["matched"]} → 件数 {t["match_rate_n"]}% / 金額 {t["match_rate_amt"]}%')
    if "--dump" in sys.argv:
        out = sys.argv[sys.argv.index("--dump") + 1]
        json.dump(rep, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("dumped", out)


if __name__ == "__main__":
    main()
