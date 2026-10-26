"""画面表示まわりの純粋ロジック（Streamlit に依存しない。回帰テストから直接呼べる）

2026-08-04 フィードバック対応で追加:
  - project_display_label / project_full_label: 案件の表示名
  - merge_proposal_edits: 打合せ変更タブの表編集を変更案リストへ書き戻す
"""
from __future__ import annotations

import math
import re
from typing import Any


def _clean(text: Any) -> str:
    return str(text).strip() if text is not None else ''


def project_display_label(pinfo: dict) -> str:
    """選択欄・一覧カードの見出しに出す名前。

    変更見積（別プラン）は、作成時に入力したプラン名（タイトル）だけを返す。
    元見積名の後ろにプラン名を連結すると選択欄で肝心のタイトルが切れて
    どの変更か分からなくなる（2026-08-04 お客様ご指摘②）。
    選択欄はお施主様ごとに絞られているため、プラン名だけで識別できる。
    """
    if not isinstance(pinfo, dict):
        return ''
    plan = _clean(pinfo.get('plan'))
    return plan or _clean(pinfo.get('name'))


def project_full_label(pinfo: dict) -> str:
    """元見積名とプラン名の両方を併記する表示名（変更履歴など文脈が無い場所向け）"""
    if not isinstance(pinfo, dict):
        return ''
    name = _clean(pinfo.get('name'))
    plan = _clean(pinfo.get('plan'))
    return f"{name}（{plan}）" if plan else name


def _is_nan(value: Any) -> bool:
    return isinstance(value, float) and math.isnan(value)


def _to_int_amount(value: Any, fallback: Any) -> Any:
    """表の金額セルを整数円に正規化する。空欄/NaN は 0、変換不能は元の値を保つ。"""
    if value is None or _is_nan(value):
        return 0
    if isinstance(value, bool):
        return fallback
    if isinstance(value, (int, float)):
        return max(0, int(value))
    text = _clean(value).replace(',', '').replace('円', '')
    if not text:
        return 0
    try:
        return max(0, int(float(text)))
    except ValueError:
        return fallback


def _to_bool(value: Any) -> bool:
    if value is None or _is_nan(value):
        return False
    return bool(value)


def merge_proposal_edits(proposals: list[dict] | None,
                         edited_rows: list[dict] | None,
                         label_to_ctype: dict[str, str],
                         apply_all: bool | None = None) -> list[dict]:
    """data_editor で編集した行を変更案リストに書き戻し、反映フラグを揃える。

    「全て選択／全て解除」を押したときに、表で直した品名・金額・内容を
    失わずにチェックだけを一括変更するために使う（2026-08-04 お客様ご指摘①）。

    Args:
        proposals: AI が抽出した変更案（category / item_name / change_type /
            proposed / amount / reason / confidence ...）。変更せず複製して返す。
        edited_rows: data_editor の戻り値を to_dict('records') した行。
            列は '反映' '工事種別' '品名' '変更の種類' '変更後の内容' '金額(円)' '根拠'。
            列が無ければ元の値を保つ。
        label_to_ctype: 「単価変更」→'price' のような表示ラベル→種別コードの対応。
        apply_all: True なら全行チェック、False なら全行解除、None なら表の値を採用。

    Returns:
        行数は proposals と同じ。余分な表の行は無視し、不足分は元の変更案を保つ。
    """
    proposals = list(proposals or [])
    rows = list(edited_rows or [])
    merged: list[dict] = []
    for idx, prop in enumerate(proposals):
        p = dict(prop) if isinstance(prop, dict) else {}
        row = rows[idx] if idx < len(rows) and isinstance(rows[idx], dict) else {}

        if '工事種別' in row:
            p['category'] = _clean(row.get('工事種別'))
        if '品名' in row:
            p['item_name'] = _clean(row.get('品名'))
        if '変更の種類' in row:
            p['change_type'] = label_to_ctype.get(_clean(row.get('変更の種類')),
                                                  p.get('change_type'))
        if '変更後の内容' in row:
            p['proposed'] = _clean(row.get('変更後の内容'))
        if '金額(円)' in row:
            p['amount'] = _to_int_amount(row.get('金額(円)'), p.get('amount'))
        if '根拠' in row:
            p['reason'] = _clean(row.get('根拠'))

        if apply_all is None:
            p['apply'] = _to_bool(row.get('反映', p.get('apply', False)))
        else:
            p['apply'] = bool(apply_all)
        merged.append(p)
    return merged


# ---------------------------------------------------------------- 2026-08-07 MTG 宿題
DEFAULT_COST_RATIO = 0.67   # 粗利33%運用 → 原価率67%（原価未設定の明細にだけ使う）


def linked_order_price(old_estimate: Any, old_order: Any, new_estimate: Any,
                       default_cost_ratio: float = DEFAULT_COST_RATIO) -> int:
    """見積単価（販価）を変えたとき、同じ粗利率を保つように原価（発注単価）を再計算する。

    2026-08-07 MTG: 打合せ変更で販価を更新しても原価が据え置きになっていた
    （UB 26万→371,500円、キッチン 36万→514,300円）。その明細の元の原価率を維持し、
    元の原価が無い明細は粗利33%（原価率67%）で計算する。
    """
    try:
        new_est = float(new_estimate or 0)
    except (TypeError, ValueError):
        return 0
    if new_est <= 0:
        return 0
    try:
        old_est = float(old_estimate or 0)
        old_ord = float(old_order or 0)
    except (TypeError, ValueError):
        old_est, old_ord = 0.0, 0.0
    ratio = old_ord / old_est if old_est > 0 and old_ord > 0 else default_cost_ratio
    ratio = min(max(ratio, 0.0), 1.0)
    return int(round(new_est * ratio))


_DIMENSION_TOKEN = re.compile(
    r'(?<![0-9０-９])[A-Za-zＡ-Ｚａ-ｚφΦøØ]{1,2}\s*[0-9０-９]+(?:[.．][0-9０-９]+)?'
    r'|[0-9０-９]+(?:[.．][0-9０-９]+)?\s*(?:mm|㎜|cm|㎝|ｍｍ)\b'
)


def strip_dimension_tokens(text: str) -> str:
    """『W2700』『H850』『D650』『φ100』『2700mm』など寸法表記を取り除く。

    2026-08-07 MTG: カップボード「W2700」の 2700 を数量 2700 と誤認識していた。
    アルファベット（W/H/D 等）や単位 mm が付いた数値はサイズであり数量・金額ではない。
    """
    return _DIMENSION_TOKEN.sub(' ', str(text or ''))


def parse_amount_jp(text: Any) -> float | None:
    """『150万』『1,500,000円』『2台』等の文字列から数値を取り出す（寸法表記は無視）。

    優先順位: 「〜万」 → 「〜円」 → 寸法を除いた最初の数値。
    """
    if text is None:
        return None
    if isinstance(text, bool):
        return None
    if isinstance(text, (int, float)):
        if isinstance(text, float) and math.isnan(text):
            return None
        return float(text)
    s = str(text).replace(',', '').replace('，', '').strip()
    s = s.translate(str.maketrans('０１２３４５６７８９．', '0123456789.'))
    s = strip_dimension_tokens(s)
    m = re.search(r'(\d+(?:\.\d+)?)\s*万', s)
    if m:
        return float(m.group(1)) * 10000
    m = re.search(r'(\d+(?:\.\d+)?)\s*円', s)
    if m:
        return float(m.group(1))
    m = re.search(r'\d+(?:\.\d+)?', s)
    return float(m.group(0)) if m else None


# フィードバックのステータス（2026-08-07 MTG で合意した3段階フロー）
FB_STATUS_OPEN = '未対応'
FB_STATUS_FIXED = '高井対応済み'      # 開発側が修正・反映した
FB_STATUS_DONE = '対応完了'           # お客様が確認してOK
FB_STATUS_RETURNED = '差し戻し'       # お客様が確認してNG（コメント必須）
FB_STATUSES = [FB_STATUS_OPEN, FB_STATUS_FIXED, FB_STATUS_DONE, FB_STATUS_RETURNED]
_LEGACY_STATUS = {'対応済み': FB_STATUS_DONE, '': FB_STATUS_OPEN, None: FB_STATUS_OPEN}


def normalize_fb_status(status: Any) -> str:
    """旧2値（未対応/対応済み）を新フローの値に読み替える。未知の値は未対応扱い。"""
    if status in _LEGACY_STATUS:
        return _LEGACY_STATUS[status]
    s = _clean(status)
    return s if s in FB_STATUSES else FB_STATUS_OPEN


def fb_status_color(status: Any) -> str:
    return {FB_STATUS_OPEN: '#dc2626', FB_STATUS_FIXED: '#2563eb',
            FB_STATUS_DONE: '#16a34a', FB_STATUS_RETURNED: '#d97706'}.get(
        normalize_fb_status(status), '#dc2626')


def validate_fb_status_change(new_status: Any, note: Any) -> str | None:
    """ステータス変更の入力チェック。問題なければ None、あればエラーメッセージ。"""
    new_status = _clean(new_status)
    if new_status not in FB_STATUSES:
        return 'ステータスの値が不正です'
    if new_status == FB_STATUS_RETURNED and not _clean(note):
        return '差し戻しの場合は、どこが問題かをコメントに入力してください'
    return None


def apply_fb_status_change(fb: dict, new_status: str, note: str, actor: str, at: str) -> dict:
    """フィードバック1件にステータス変更を適用した新しい dict を返す（履歴を残す）。"""
    out = dict(fb)
    out['status'] = normalize_fb_status(new_status)
    out['status_note'] = _clean(note)
    out['status_updated_at'] = at
    out['status_updated_by'] = _clean(actor)
    history = list(out.get('status_history') or [])
    history.append({'status': out['status'], 'note': out['status_note'],
                    'by': out['status_updated_by'], 'at': at})
    out['status_history'] = history
    return out
