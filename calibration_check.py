"""
較正チェックハーネス (CAD→概算見積もり生成AI / 日本住建株式会社)

calibration/ フォルダ（gitignore対象・顧客実データ）にある正解ペア
（【数量データ】TXT ＋【提出見積】xlsx）全物件に対して AI 見積を計算し、
乖離マトリクスを表示・保存する。エンジン修正のたびに実行して回帰を確認する。

使い方:
    python3 calibration_check.py            # 全物件の乖離サマリ
    python3 calibration_check.py --detail 8 # 物件番号8の区分別詳細
"""

from __future__ import annotations

import glob
import json
import os
import re
import sys

from cad_parser import load_cad_file
from estimate_calculator import calculate_estimate, load_unit_prices
import excel_diff

CALIB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "calibration")


def iter_pairs():
    """(pid, 物件名, 数量TXTパス, 提出見積xlsxパス) を列挙する"""
    for sf in sorted(glob.glob(os.path.join(CALIB_DIR, "*【数量データ】*.TXT"))):
        m = re.match(r'(\d+)-(.+?)様', os.path.basename(sf))
        if not m:
            continue
        pid, pname = m.group(1), m.group(2)
        xlsx = glob.glob(os.path.join(CALIB_DIR, f"{pid}-*【提出見積】*.xlsx*"))
        if xlsx:
            yield pid, pname, sf, xlsx[0]


def run_all(save_path: str | None = None) -> dict:
    prices = load_unit_prices(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "tankamaster_updated.csv"))
    report = {}
    print(f"{'物件':<14} {'坪数':>6} {'AI見積':>12} {'実績':>12} {'乖離':>7}")
    rates = []
    for pid, pname, sf, xlsx in iter_pairs():
        try:
            cad = load_cad_file(sf)
            actual = excel_diff.parse_actual_estimate(open(xlsx, 'rb').read())
            # 実績に太陽光が計上されている物件は、本番でトグルONにする運用に
            # 合わせて include_solar=True 側で比較する
            solar = any('太陽光' in c['name'] and c['total'] > 0
                        for c in actual['categories'])
            est = calculate_estimate(cad, prices, include_solar=solar)
            diff = excel_diff.diff_estimates(est, actual)
            rate = (diff['total_ai'] - diff['total_actual']) / diff['total_actual'] * 100
            rates.append(rate)
            print(f"{pid}-{pname[:9]:<11} {est.total_floor_area_tsubo:>5.1f} "
                  f"{diff['total_ai']:>12,} {diff['total_actual']:>12,} {rate:>+6.1f}%")
            report[f"{pid}-{pname}"] = {
                'tsubo': est.total_floor_area_tsubo,
                'total_ai': diff['total_ai'], 'total_actual': diff['total_actual'],
                'rate_pct': round(rate, 2),
                'category_diffs': diff['category_diffs'],
            }
        except Exception as exc:  # noqa: BLE001 - 1物件の失敗で全体を止めない
            print(f"{pid}-{pname}: ERROR {type(exc).__name__}: {str(exc)[:100]}")
    if rates:
        mean_abs = sum(abs(r) for r in rates) / len(rates)
        worst = max(rates, key=abs)
        print(f"\n平均絶対乖離: {mean_abs:.2f}% / 最大乖離: {worst:+.1f}% / 対象 {len(rates)}物件")
    if save_path:
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=1)
    return report


def show_detail(target_pid: str):
    prices = load_unit_prices("tankamaster_updated.csv")
    for pid, pname, sf, xlsx in iter_pairs():
        if pid != target_pid:
            continue
        cad = load_cad_file(sf)
        est = calculate_estimate(cad, prices)
        actual = excel_diff.parse_actual_estimate(open(xlsx, 'rb').read())
        diff = excel_diff.diff_estimates(est, actual)
        print(f"=== {pid}-{pname} ===")
        for c in sorted(diff['category_diffs'],
                        key=lambda c: c['ai_total'] - c['actual_total']):
            d = c['ai_total'] - c['actual_total']
            print(f"  {c['name']:<14} AI {c['ai_total']:>11,} 実績 {c['actual_total']:>11,} 差 {d:>+11,}")


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == '--detail':
        show_detail(sys.argv[2])
    else:
        run_all(os.path.join(CALIB_DIR, "latest_report.json"))
