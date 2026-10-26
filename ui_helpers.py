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


# ---------------------------------------------------------------- 問い合わせの保持・アーカイブ
ACTIVE_FB_STATUSES = [FB_STATUS_OPEN, FB_STATUS_FIXED, FB_STATUS_RETURNED]


def fb_is_archived(fb: dict) -> bool:
    """両者の確認が済んだ（お客様が「対応完了」にした）問い合わせはアーカイブ扱い。"""
    return normalize_fb_status((fb or {}).get('status')) == FB_STATUS_DONE


def partition_feedbacks(fbs: list | None) -> tuple[list, list]:
    """(対応中, アーカイブ) に分ける。順序は保つ。削除はしない＝問い合わせは必ず残る。"""
    active, archived = [], []
    for fb in fbs or []:
        if not isinstance(fb, dict):
            continue
        (archived if fb_is_archived(fb) else active).append(fb)
    return active, archived


def _last_event(fb: dict, status: str) -> dict | None:
    for ev in reversed(list(fb.get('status_history') or [])):
        if isinstance(ev, dict) and ev.get('status') == status:
            return ev
    return None


def fb_timeline(fb: dict) -> dict:
    """「いつ受け付け・いつ何を直し・いつ確認が取れたか」を1件分まとめる。

    旧データ（status_history 無し）は received のみ埋まり、他は空文字。
    """
    fb = fb or {}
    fixed = _last_event(fb, FB_STATUS_FIXED) or {}
    done = _last_event(fb, FB_STATUS_DONE) or {}
    returned = _last_event(fb, FB_STATUS_RETURNED) or {}
    return {
        'received': _clean(fb.get('timestamp')),
        'fixed_at': _clean(fixed.get('at')),
        'fixed_note': _clean(fixed.get('note')),
        'done_at': _clean(done.get('at')),
        'done_note': _clean(done.get('note')),
        'returned_at': _clean(returned.get('at')),
        'returned_note': _clean(returned.get('note')),
        'legacy': fb_is_archived(fb) and not fixed and not done,
    }


def fb_timeline_text(fb: dict) -> str:
    """カードに出す1行の履歴。例: 受付 08/04 → 高井対応済み 10/26（変更箇所）→ 対応完了 10/27"""
    t = fb_timeline(fb)
    parts = [f"受付 {t['received']}" if t['received'] else "受付 -"]
    if t['returned_at'] and (not t['fixed_at'] or t['returned_at'] > t['fixed_at']):
        parts.append(f"差し戻し {t['returned_at']}" + (f"（{t['returned_note']}）" if t['returned_note'] else ""))
    if t['fixed_at']:
        parts.append(f"高井対応済み {t['fixed_at']}" + (f"（{t['fixed_note']}）" if t['fixed_note'] else ""))
    if t['done_at']:
        parts.append(f"対応完了 {t['done_at']}" + (f"（{t['done_note']}）" if t['done_note'] else ""))
    elif t['legacy']:
        parts.append("対応完了（以前のデータ・日時記録なし）")
    return " → ".join(parts)


def release_log_rows(fbs: list | None, project_name_of=None) -> list[dict]:
    """アーカイブ一覧用の行（新しい完了順）。どの修正をいつ出したかを一覧できる。"""
    rows = []
    for fb in fbs or []:
        if not isinstance(fb, dict) or not fb_is_archived(fb):
            continue
        t = fb_timeline(fb)
        rows.append({
            '対応完了': t['done_at'] or ('旧データ' if t['legacy'] else ''),
            '修正反映': t['fixed_at'],
            '変更箇所': t['fixed_note'],
            '案件': _clean(project_name_of(fb)) if project_name_of else _clean(fb.get('_proj_name') or fb.get('property')),
            'カテゴリ': _clean(fb.get('category')),
            '受付': t['received'],
            'ご指摘': _clean(fb.get('body') or fb.get('message'))[:60],
        })
    rows.sort(key=lambda r: (r['対応完了'] if r['対応完了'] != '旧データ' else '', r['受付']), reverse=True)
    return rows


# ---------------------------------------------------------------- 実績比較: 工事種別ごとの明細ドリルダウン
DETAIL_KIND_DIFF = '金額が違う'
DETAIL_KIND_AI_ONLY = 'AIにしかない'
DETAIL_KIND_ACTUAL_ONLY = '実績にしかない'
DETAIL_KIND_SAME = '一致'
DETAIL_KIND_ORDER = [DETAIL_KIND_DIFF, DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY, DETAIL_KIND_SAME]


def _qty_text(qty: Any, unit: Any) -> str:
    try:
        q = float(qty or 0)
    except (TypeError, ValueError):
        return ''
    if not q:
        return ''
    return f"{q:g}{_clean(unit)}"


PROJECT_SPECIFIC_LABEL = '案件特有'   # 図面に表れない物件固有の項目（全案件には該当しない）
PROJECT_SPECIFIC_CATEGORY = '案件特有（図面に表れない項目）'
NOTE_SCOPE_GLOBAL = '全案件共通'
NOTE_SCOPE_PROJECT = 'この案件のみ'


def category_detail_rows(item_diff: dict | None, special_names: set | None = None) -> list[dict]:
    """excel_diff の item_diffs[種別] を、画面に出す明細行に変換する。

    区分: 金額が違う → AIにしかない → 実績にしかない → 一致 の順、各区分内は差額の大きい順。
    special_names: 「この案件特有（全案件には該当しない）」と登録済みの品名。'扱い' 列に表示する。
    """
    d = item_diff or {}
    rows: list[dict] = []
    special = {_clean(n) for n in (special_names or set()) if _clean(n)}
    for m in d.get('matched', []) or []:
        if not isinstance(m, dict):
            continue
        ai_amt = int(m.get('ai_amount') or 0)
        ac_amt = int(m.get('actual_amount') or 0)
        rows.append({
            '区分': DETAIL_KIND_SAME if ai_amt == ac_amt else DETAIL_KIND_DIFF,
            'AI品名': _clean(m.get('ai_name')),
            'AI摘要': _clean(m.get('ai_summary')),
            'AI数量': _qty_text(m.get('ai_quantity'), m.get('unit')),
            'AI金額': ai_amt,
            '実績品名': _clean(m.get('actual_name')),
            '実績摘要': _clean(m.get('actual_summary')),
            '実績数量': _qty_text(m.get('actual_quantity'), m.get('unit')),
            '実績金額': ac_amt,
            '差額(AI−実績)': ai_amt - ac_amt,
            '登録済': bool(m.get('via_alias')),
        })
    for x in d.get('ai_only', []) or []:
        if not isinstance(x, dict):
            continue
        amt = int(x.get('amount') or 0)
        rows.append({'区分': DETAIL_KIND_AI_ONLY, 'AI品名': _clean(x.get('name')), 'AI摘要': _clean(x.get('summary')),
                     'AI数量': _qty_text(x.get('quantity'), x.get('unit')), 'AI金額': amt,
                     '実績品名': '', '実績摘要': '', '実績数量': '', '実績金額': 0, '差額(AI−実績)': amt,
                     '登録済': False})
    for x in d.get('actual_only', []) or []:
        if not isinstance(x, dict):
            continue
        amt = int(x.get('amount') or 0)
        rows.append({'区分': DETAIL_KIND_ACTUAL_ONLY, 'AI品名': '', 'AI摘要': '', 'AI数量': '', 'AI金額': 0,
                     '実績品名': _clean(x.get('name')), '実績摘要': _clean(x.get('summary')),
                     '実績数量': _qty_text(x.get('quantity'), x.get('unit')), '実績金額': amt,
                     '差額(AI−実績)': -amt, '登録済': False})
    for r in rows:
        r['扱い'] = PROJECT_SPECIFIC_LABEL if (r['AI品名'] in special or r['実績品名'] in special) else ''
    rows.sort(key=lambda r: (DETAIL_KIND_ORDER.index(r['区分']), -abs(r['差額(AI−実績)'])))
    return rows


def special_count(rows: list[dict]) -> int:
    return sum(1 for r in rows or [] if r.get('扱い') == PROJECT_SPECIFIC_LABEL)


def validate_special_selection(selected_rows: list[dict]) -> tuple[str | None, list[dict]]:
    """「この案件特有」として登録する行の入力チェック。一致行は対象外。"""
    rows = [r for r in selected_rows or [] if isinstance(r, dict)]
    if not rows:
        return '登録する行にチェックを付けてください', []
    bad = [r for r in rows if r.get('区分') == DETAIL_KIND_SAME]
    if bad:
        return '「一致」している行は案件特有として登録できません', []
    return None, rows


def special_note_name(row: dict) -> str:
    """案件特有として記録するときの品名（実績側があれば実績名、無ければAI名）"""
    return _clean(row.get('実績品名')) or _clean(row.get('AI品名'))


def detail_kind_counts(rows: list[dict]) -> dict[str, int]:
    """区分ごとの件数（順序付き・0件も含む）"""
    counts = {k: 0 for k in DETAIL_KIND_ORDER}
    for r in rows or []:
        k = r.get('区分')
        if k in counts:
            counts[k] += 1
    counts[PROJECT_SPECIFIC_LABEL] = special_count(rows)
    return counts


def detail_header(cat: dict, counts: dict[str, int]) -> str:
    """展開ボタン（expander）の見出し。例: 電気設備工事　AI 1,243,400 / 実績 1,246,550 / 差額 -3,150　金額が違う2・AIのみ1"""
    ai = int(cat.get('ai_total') or 0)
    ac = int(cat.get('actual_total') or 0)
    parts = [f"{DETAIL_KIND_DIFF} {counts.get(DETAIL_KIND_DIFF, 0)}",
             f"AIのみ {counts.get(DETAIL_KIND_AI_ONLY, 0)}",
             f"実績のみ {counts.get(DETAIL_KIND_ACTUAL_ONLY, 0)}",
             f"一致 {counts.get(DETAIL_KIND_SAME, 0)}"]
    special = counts.get(PROJECT_SPECIFIC_LABEL, 0)
    if special:
        parts.append(f"{PROJECT_SPECIFIC_LABEL} {special}")
    return (f"{_clean(cat.get('name'))}　AI {ai:,}円 / 実績 {ac:,}円 / 差額 {ai - ac:+,}円　"
            f"｜ {'・'.join(parts)}")


# ---------------------------------------------------------------- 実績比較: 「同じ商品」登録（案件横断の名寄せ辞書）
def _alias_keys(name: Any, summary: Any) -> list[str]:
    from excel_diff import item_alias_keys
    return sorted(item_alias_keys({'name': _clean(name), 'summary': _clean(summary)}))


ALIAS_SCOPE_GLOBAL = 'global'    # 普遍的な表記ゆれ（全案件に適用）
ALIAS_SCOPE_PROJECT = 'project'  # この案件特有の読み替え（その案件だけに適用）
ALIAS_SCOPE_LABEL = {ALIAS_SCOPE_GLOBAL: '全案件共通', ALIAS_SCOPE_PROJECT: 'この案件のみ'}


def make_alias_entry(ai_row: dict, actual_row: dict, category: str, property_name: str,
                     at: str, by: str = '', scope: str = ALIAS_SCOPE_GLOBAL,
                     project_id: str = '') -> dict:
    """ドリルダウン表の AIのみ行 と 実績のみ行 から辞書エントリを作る。

    scope: 'global'（普遍的＝全案件に適用）か 'project'（この案件特有＝project_id の案件だけ）。
    学習させるたびに「普遍的か案件特有か」を選び、後からも切り替えられる。
    """
    ai_name, ai_sum = _clean(ai_row.get('AI品名')), _clean(ai_row.get('AI摘要'))
    ac_name, ac_sum = _clean(actual_row.get('実績品名')), _clean(actual_row.get('実績摘要'))
    return {
        'ai_name': ai_name, 'ai_summary': ai_sum, 'ai_keys': _alias_keys(ai_name, ai_sum),
        'actual_name': ac_name, 'actual_summary': ac_sum, 'actual_keys': _alias_keys(ac_name, ac_sum),
        'category': _clean(category), 'property': _clean(property_name), 'at': _clean(at), 'by': _clean(by),
        'scope': scope if scope in ALIAS_SCOPE_LABEL else ALIAS_SCOPE_GLOBAL,
        'project_id': _clean(project_id),
    }


def alias_applies(entry: dict, project_id: str | None) -> bool:
    """このエントリを project_id の案件に適用してよいか（全案件共通 or 同じ案件の特有分）"""
    if not isinstance(entry, dict):
        return False
    if entry.get('scope', ALIAS_SCOPE_GLOBAL) != ALIAS_SCOPE_PROJECT:
        return True
    return bool(project_id) and _clean(entry.get('project_id')) == _clean(project_id)


def alias_pairs_from_library(lib: list | None, project_id: str | None = None) -> set:
    """辞書エントリ一覧 → 突合に渡す frozenset の集合（名称のみ／名称＋摘要の全組み合わせ）。

    project_id を渡すと、他案件特有のエントリは除外する。
    """
    pairs: set = set()
    for e in lib or []:
        if not alias_applies(e, project_id):
            continue
        for a in e.get('ai_keys') or []:
            for b in e.get('actual_keys') or []:
                if a and b:
                    pairs.add(frozenset((a, b)))
    return pairs


def alias_entry_exists(lib: list | None, entry: dict) -> bool:
    for e in lib or []:
        if (isinstance(e, dict) and e.get('ai_keys') == entry.get('ai_keys')
                and e.get('actual_keys') == entry.get('actual_keys')):
            return True
    return False


def validate_alias_selection(selected_rows: list[dict]) -> tuple[str | None, dict | None, dict | None]:
    """チェックされた行から (エラー, AIのみ行, 実績のみ行) を返す。AIのみ1件＋実績のみ1件のときだけ有効。"""
    ai_rows = [r for r in selected_rows or [] if r.get('区分') == DETAIL_KIND_AI_ONLY]
    ac_rows = [r for r in selected_rows or [] if r.get('区分') == DETAIL_KIND_ACTUAL_ONLY]
    others = [r for r in selected_rows or [] if r.get('区分') not in (DETAIL_KIND_AI_ONLY, DETAIL_KIND_ACTUAL_ONLY)]
    if others:
        return '「AIにしかない」と「実績にしかない」の行だけを選んでください', None, None
    if len(ai_rows) != 1 or len(ac_rows) != 1:
        return f'「AIにしかない」を1件、「実績にしかない」を1件チェックしてください（現在 AI {len(ai_rows)}件・実績 {len(ac_rows)}件）', None, None
    return None, ai_rows[0], ac_rows[0]
