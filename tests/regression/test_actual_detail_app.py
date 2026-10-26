"""実績比較: 工事種別ごとの「詳細を見る」展開を app.py 本体（AppTest）で確認"""
import json
import os
import shutil

import excel_diff
from tests.regression.test_feedback_20260804_app import _boot, app_dir, _no_cloud, ROOT  # noqa: F401


def test_category_drilldown_shows_item_kinds(app_dir):
    # サンプル邸の提出見積（公開リポジトリ同梱）を実績として取り込んだ状態にする
    with open(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx"), "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    at = _boot(app_dir)
    assert not at.exception, [str(e.value) for e in at.exception]

    headers = [e.label for e in at.expander if "AI " in e.label and "実績 " in e.label and "差額 " in e.label]
    assert len(headers) >= 20, headers[:3]
    # 見出しに 工事種別・AI/実績/差額・区分ごとの件数 が入っている
    h = headers[0]
    assert "金額が違う" in h and "AIのみ" in h and "実績のみ" in h and "一致" in h
    assert "円" in h

    # 展開内の表には区分列があり、4区分のいずれか
    tables = [d.value for d in at.dataframe if "区分" in list(d.value.columns) and "差額(AI−実績)" in list(d.value.columns)]
    assert tables, "明細の表が無い"
    kinds = set()
    for t in tables:
        kinds |= set(t["区分"])
        # AI金額 − 実績金額 = 差額 が全行で成り立つ
        assert all(int(r["AI金額"]) - int(r["実績金額"]) == int(r["差額(AI−実績)"]) for _, r in t.iterrows())
    assert kinds <= {"金額が違う", "AIにしかない", "実績にしかない", "一致"}
    assert "実績にしかない" in kinds or "AIにしかない" in kinds or "金額が違う" in kinds

    # 工事種別の合計表（AI見積/実績見積）と展開内の合計が一致する（集約で金額は変わらない）
    cat_table = [d.value for d in at.dataframe if "AI見積" in list(d.value.columns)][0]
    # AI側は明細の積み上げ＝種別合計なので、展開内の合計と一致する
    # （実績側は帳票上の種別合計と明細合計が一致しない行があり得るため比較しない）
    total_ai_detail = sum(int(t["AI金額"].sum()) for t in tables)
    assert total_ai_detail == int(cat_table["AI見積"].sum())


def test_alias_and_project_specific_are_reflected(app_dir):
    """登録済みの「同じ商品」は突合され、「この案件特有」は扱い列に出る（他案件特有の辞書は使われない）"""
    from ui_helpers import make_alias_entry, ALIAS_SCOPE_PROJECT, PROJECT_SPECIFIC_CATEGORY
    with open(os.path.join(ROOT, "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx"), "rb") as f:
        actual = excel_diff.parse_actual_estimate(f.read())
    with open(os.path.join(app_dir, "projects", "sample_furui_actual.json"), "w", encoding="utf-8") as f:
        json.dump(actual, f, ensure_ascii=False)
    ai_row = {'AI品名': 'SGL鋼板 たてひら葺', 'AI摘要': 'ﾆｽｸｶﾗｰPro ※唐草･棟含む'}
    ac_row = {'実績品名': '平板瓦葺（切妻/寄棟）', '実績摘要': 'ｾﾗﾑFF程度　標準勾配　ｼｽﾃﾑ瓦対応'}
    lib = [make_alias_entry(ai_row, ac_row, '屋根工事', 'サンプル邸', '2026-10-26 18:00'),
           make_alias_entry({'AI品名': '人通口補強', 'AI摘要': '60㎡超過分'}, {'実績品名': '布基礎', '実績摘要': ''},
                            '基礎工事', '別の邸', '2026-10-26 18:00', scope=ALIAS_SCOPE_PROJECT, project_id='other_pid')]
    with open(os.path.join(app_dir, "projects", "item_alias_library.json"), "w", encoding="utf-8") as f:
        json.dump(lib, f, ensure_ascii=False)
    with open(os.path.join(app_dir, "projects", "sample_furui_diff_notes.json"), "w", encoding="utf-8") as f:
        json.dump({"基礎工事|深基礎": {"category": PROJECT_SPECIFIC_CATEGORY, "reason": "図面に表れない物件固有の項目",
                                     "scope": "project", "diff": -188485}}, f, ensure_ascii=False)
    at = _boot(app_dir)
    assert not at.exception, [str(e.value) for e in at.exception]
    tables = [d.value for d in at.dataframe if "区分" in list(d.value.columns) and "差額(AI−実績)" in list(d.value.columns)]
    rows = [r for t in tables for _, r in t.iterrows()]
    roof = [r for r in rows if r["AI品名"] == "SGL鋼板 たてひら葺"]
    assert roof and bool(roof[0]["登録済"]) is True and roof[0]["区分"] == "金額が違う"
    assert roof[0]["実績品名"] == "平板瓦葺（切妻/寄棟）"
    other = [r for r in rows if r["AI品名"] == "人通口補強"]
    assert other and other[0]["区分"] == "AIにしかない", "他案件特有の辞書が使われている"
    deep = [r for r in rows if r["実績品名"] == "深基礎"]
    assert deep and deep[0]["扱い"] == "案件特有"
    headers = [e.label for e in at.expander]
    assert any("案件特有 1" in h for h in headers)
    assert any("登録済みの「同じ商品」（この案件に適用 1件・他案件特有 1件は非表示）" in h for h in headers), headers[-3:]
    # 差異理由エディタに適用範囲の列がある
    assert any("適用範囲" in list(d.value.columns) for d in at.dataframe)
    # 登録済みの組は差異理由の対象（金額差）として並ぶ
    note_tables = [d.value for d in at.dataframe if "差異の理由" in list(d.value.columns)]
    assert note_tables and (note_tables[0]["品名"] == "SGL鋼板 たてひら葺").any()
