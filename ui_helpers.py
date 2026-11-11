"""画面表示まわりの純粋ロジック（Streamlit に依存しない。回帰テストから直接呼べる）

2026-08-04 フィードバック対応で追加:
  - project_display_label / project_full_label: 案件の表示名
  - merge_proposal_edits: 打合せ変更タブの表編集を変更案リストへ書き戻す
"""
from __future__ import annotations

import math
import re
import unicodedata
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


_NUM = r'[0-9０-９]+(?:[.．,，][0-9０-９]+)*'
_MM = r'\s*(?:mm|㎜|cm|㎝|ｍｍ|ｃｍ)(?![A-Za-z])'
# 数字の直後が「円・万・¥」（金額）や「m・メートル」（長さの数量）なら寸法ではない
_NOT_MONEY_OR_METER = r'(?![0-9０-９.．,，])(?!\s*(?:円|万|¥|￥|[mｍ](?![mｍ])|メートル))'
_DIMENSION_TOKEN = re.compile(
    # W2700 / H850 / D650 / φ100 / UB1616（英字1〜2文字＋数字）
    r'(?<![0-9０-９A-Za-zＡ-Ｚａ-ｚ])[A-Za-zＡ-Ｚａ-ｚφΦøØ]{1,2}\s*' + _NUM + r'(?:' + _MM + r'|' + _NOT_MONEY_OR_METER + r')'
    # W:2700 / W＝2700（区切り付きは寸法記号1文字だけ。UB：400,000円 のような略号＋金額は消さない）
    r'|(?<![0-9０-９A-Za-zＡ-Ｚａ-ｚ])[WHDLwhdlＷＨＤＬφΦ]\s*[:：=＝]\s*' + _NUM + r'(?:' + _MM + r'|' + _NOT_MONEY_OR_METER + r')'
    # 幅2700 / 間口 1820 / 奥行き650 / 高さ850（長さ12m のような m 単位の数量は消さない）
    r'|(?:幅|巾|間口|奥行き?|高さ|厚み|厚さ)\s*[:：=＝]?\s*' + _NUM + r'(?:' + _MM + r'|' + _NOT_MONEY_OR_METER + r')'
    # 2700mm / 2700mmの / 2700mmタイプ
    r'|' + _NUM + _MM
)
# 型番（TKS05305JA・XCH1601RWS・TCF8GM・SC419a）。単位の無い数字を拾うときに先に消す
_MODEL_TOKEN = re.compile(
    r'(?<![0-9A-Za-z])(?:[A-Za-z]{2,}[0-9][A-Za-z0-9-]*|[A-Za-z][0-9]+[A-Za-z][A-Za-z0-9-]*|[0-9]+[A-Za-z]{2,}[A-Za-z0-9-]*)')


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


def _normalize_digits(text: Any, keep_commas: bool = False) -> str:
    s = str(text).strip().translate(str.maketrans('０１２３４５６７８９．，', '0123456789.,'))
    return s if keep_commas else s.replace(',', '')


def _drop_commas_in_numbers(s: str) -> str:
    return re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', s)


def _bare_number(s: str) -> float | None:
    """寸法・型番を除いたあとの、単位の無い最初の数字"""
    s = _MODEL_TOKEN.sub(' ', s)
    m = re.search(r'(?<![0-9.])\d+(?:\.\d+)?', s)
    return float(m.group(0)) if m else None


def parse_money_jp(text: Any) -> float | None:
    """金額だけを読む: 『円』『万』『¥』が付いた数字、または桁区切り（1,500,000）の数字を金額とみなす。

    寸法（W2700・2700mm）・型番（XCH1601RWS）・サイズ（サザナ 1616サイズ）の数字は金額にしない。
    2026-11-12 監査: 仕様変更の文の 1616 が単価になっていた。
    """
    if text is None or isinstance(text, bool):
        return None
    if isinstance(text, (int, float)):
        return None if (isinstance(text, float) and math.isnan(text)) else float(text)
    raw = strip_dimension_tokens(_normalize_digits(text, keep_commas=True))
    s = _drop_commas_in_numbers(raw)
    m = re.search(r'(\d+(?:\.\d+)?)\s*万', s)
    if m:
        return float(m.group(1)) * 10000
    m = re.search(r'[¥￥]\s*(\d+(?:\.\d+)?)', s)
    if m:
        return float(m.group(1))
    m = re.search(r'(\d+(?:\.\d+)?)\s*円', s)
    if m:
        return float(m.group(1))
    # 桁区切りのある数字は金額とみなす（1616・2550 のような区切りの無い数字は除く）
    m = re.search(r'(?<![0-9A-Za-z.])\d{1,3}(?:,\d{3})+(?![0-9])', _MODEL_TOKEN.sub(' ', raw))
    if m:
        return float(m.group(0).replace(',', ''))
    return None


_QTY_UNITS = (r'台|本|枚|個|箇所|ヶ所|ケ所|カ所|か所|ヵ所|セット|式|組|基|面|段|帖|畳|坪|㎡|m2|ｍ2|平米|'
              r'm(?![A-Za-z0-9])|ｍ|メートル|ヶ')


def parse_quantity_jp(text: Any, allow_bare: bool = False) -> float | None:
    """数量だけを読む: 『台・本・枚・箇所・㎡・m』などの単位が付いた数字を数量とみなす。

    寸法表記（W2700・2700mm・幅2700・W:2700 など）は先に取り除く。
    allow_bare=True のときは、単位が無い数字も（寸法と型番を除いたうえで）数量として読む。
    """
    if text is None or isinstance(text, bool):
        return None
    if isinstance(text, (int, float)):
        return None if (isinstance(text, float) and math.isnan(text)) else float(text)
    s = _drop_commas_in_numbers(strip_dimension_tokens(_normalize_digits(text, keep_commas=True)))
    m = re.search(r'(?<![0-9.A-Za-z])(\d+(?:\.\d+)?)\s*(?:' + _QTY_UNITS + ')', _MODEL_TOKEN.sub(' ', s))
    if m:
        return float(m.group(1))
    return _bare_number(s) if allow_bare else None


def proposal_value(change_type: str, proposed: Any, amount: Any = None) -> float | None:
    """打合せ変更の変更案から使う数字。見つからなければ None（0 と読めたときは 0.0）。

    - 表で入れた金額（amount > 0）は最優先
    - 数量変更: 単位付きの数字（無ければ単位なしの数字）
    - 単価変更: 金額表記（円・万・¥・桁区切り）、無ければ単位なしの数字
    - 仕様変更・追加: 金額表記（円・万・¥・桁区切り）だけ（サイズや型番の数字は使わない）
    """
    if isinstance(amount, (int, float)) and not isinstance(amount, bool) and not (
            isinstance(amount, float) and math.isnan(amount)) and amount > 0:
        return float(amount)
    ctype = _clean(change_type)
    if ctype == 'quantity':
        return parse_quantity_jp(proposed, allow_bare=True)
    if ctype == 'price':
        v = parse_money_jp(proposed)
        if v is None and proposed is not None:
            v = _bare_number(_drop_commas_in_numbers(strip_dimension_tokens(_normalize_digits(proposed, keep_commas=True))))
        return v
    return parse_money_jp(proposed)


def proposal_number(change_type: str, proposed: Any, amount: Any = None) -> int | float:
    """表の『金額(円)』欄に入れる初期値（見つからなければ 0）。中身は proposal_value と同じ。"""
    v = proposal_value(change_type, proposed, amount)
    if v is None:
        return 0
    return int(v) if float(v).is_integer() else v


# フィードバックのステータス（2026-08-07 MTG で合意した3段階フロー）
FB_STATUS_OPEN = '未対応'
FB_STATUS_WIP = '着手中'              # 開発側が着手・試行中（進み具合が分かるように。2026-10-26 MTG）
FB_STATUS_FIXED = '高井対応済み'      # 開発側が修正・反映した
FB_STATUS_DONE = '対応完了'           # お客様が確認してOK
FB_STATUS_RETURNED = '差し戻し'       # お客様が確認してNG（コメント必須）
FB_STATUSES = [FB_STATUS_OPEN, FB_STATUS_WIP, FB_STATUS_FIXED, FB_STATUS_DONE, FB_STATUS_RETURNED]
_LEGACY_STATUS = {'対応済み': FB_STATUS_DONE, '': FB_STATUS_OPEN, None: FB_STATUS_OPEN}


def normalize_fb_status(status: Any) -> str:
    """旧2値（未対応/対応済み）を新フローの値に読み替える。未知の値は未対応扱い。"""
    if status in _LEGACY_STATUS:
        return _LEGACY_STATUS[status]
    s = _clean(status)
    return s if s in FB_STATUSES else FB_STATUS_OPEN


def fb_status_color(status: Any) -> str:
    return {FB_STATUS_OPEN: '#dc2626', FB_STATUS_WIP: '#7c3aed', FB_STATUS_FIXED: '#2563eb',
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
ACTIVE_FB_STATUSES = [FB_STATUS_OPEN, FB_STATUS_WIP, FB_STATUS_FIXED, FB_STATUS_RETURNED]


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
    """カードに出す1行の履歴。受付から今までの変更を、起きた順に全部並べる。

    例: 受付 08/04 → 着手中 10/26（試行中）→ 高井対応済み 10/26（変更箇所）→ 差し戻し 10/27（理由）→ …
    2026-11-12 監査: 最新の1件ずつしか出さず、着手中や1回目の修正・差し戻しの経緯が見えなかった。
    """
    fb = fb or {}
    received = _clean(fb.get('timestamp'))
    parts = [f"受付 {received}" if received else "受付 -"]
    events = [ev for ev in (fb.get('status_history') or []) if isinstance(ev, dict) and _clean(ev.get('status'))]
    for ev in events:
        st = normalize_fb_status(ev.get('status'))
        at, note = _clean(ev.get('at')), _clean(ev.get('note'))
        parts.append(f"{st} {at}".strip() + (f"（{note}）" if note else ""))
    if not events and fb_is_archived(fb):
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
            '経緯': fb_timeline_text(fb),
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
EXCLUDED_LABEL = '除外'               # AIが計上したが不要な項目（理由付き）
EXCLUDED_CATEGORY = 'AI項目の除外（不要）'
EXCLUSION_REASONS = ['家の構造の違い', '地域・条例の違い', '仕様に含まない（オプション外）', '図面上は不要', 'その他']
NOTE_SCOPE_GLOBAL = '全案件共通'
NOTE_SCOPE_PROJECT = 'この案件のみ'


def category_detail_rows(item_diff: dict | None, special_names=None) -> list[dict]:
    """excel_diff の item_diffs[種別] を、画面に出す明細行に変換する。

    区分: 金額が違う → AIにしかない → 実績にしかない → 一致 の順、各区分内は差額の大きい順。
    special_names: 登録済みの扱い。set なら全て「案件特有」、dict なら {品名: 扱いラベル}
        （例: '案件特有' / '除外: 地域・条例の違い'）。'扱い' 列に表示する。
    """
    d = item_diff or {}
    rows: list[dict] = []
    if isinstance(special_names, dict):
        special = {}
        for n, v in special_names.items():
            key = (_clean(n[0]), _clean(n[1])) if isinstance(n, tuple) and len(n) == 2 else _clean(n)
            if key:
                special[key] = _clean(v) or PROJECT_SPECIFIC_LABEL
    else:
        special = {_clean(n): PROJECT_SPECIFIC_LABEL for n in (special_names or set()) if _clean(n)}
    for m in d.get('matched', []) or []:
        if not isinstance(m, dict):
            continue
        ai_amt = int(m.get('ai_amount') or 0)
        ac_amt = int(m.get('actual_amount') or 0)
        rows.append({
            '区分': DETAIL_KIND_SAME if ai_amt == ac_amt else DETAIL_KIND_DIFF,
            'グループ': _clean(m.get('group')) + ('（種別合計のみ）' if m.get('via_pseudo') else ''),
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
        rows.append({'区分': DETAIL_KIND_AI_ONLY, 'グループ': _clean(x.get('group')),
                     'AI品名': _clean(x.get('name')), 'AI摘要': _clean(x.get('summary')),
                     'AI数量': _qty_text(x.get('quantity'), x.get('unit')), 'AI金額': amt,
                     '実績品名': '', '実績摘要': '', '実績数量': '', '実績金額': 0, '差額(AI−実績)': amt,
                     '登録済': False})
    for x in d.get('actual_only', []) or []:
        if not isinstance(x, dict):
            continue
        amt = int(x.get('amount') or 0)
        rows.append({'区分': DETAIL_KIND_ACTUAL_ONLY, 'グループ': _clean(x.get('group')),
                     'AI品名': '', 'AI摘要': '', 'AI数量': '', 'AI金額': 0,
                     '実績品名': _clean(x.get('name')), '実績摘要': _clean(x.get('summary')),
                     '実績数量': _qty_text(x.get('quantity'), x.get('unit')), '実績金額': amt,
                     '差額(AI−実績)': -amt, '登録済': False})
    for r in rows:
        # 案件特有・除外は「AIにしかない／実績にしかない／金額が違う」行に登録するもの。
        # 同じ品名の「一致」行にまで付けない（2026-11-12 監査）
        if r['区分'] == DETAIL_KIND_SAME:
            r['扱い'] = ''
            continue
        label = ''
        for nm, sm in ((r['実績品名'], r.get('実績摘要', '')), (r['AI品名'], r.get('AI摘要', ''))):
            if nm and (nm, _clean(sm)) in special:          # 品名＋摘要で登録したもの
                label = special[(nm, _clean(sm))]
                break
        if not label:
            for nm in (r['実績品名'], r['AI品名']):         # 旧形式（品名だけ）
                if nm and nm in special:
                    label = special[nm]
                    break
        r['扱い'] = label
    rows.sort(key=lambda r: (DETAIL_KIND_ORDER.index(r['区分']), -abs(r['差額(AI−実績)'])))
    return rows


def special_count(rows: list[dict]) -> int:
    return sum(1 for r in rows or [] if r.get('扱い') == PROJECT_SPECIFIC_LABEL)


def excluded_count(rows: list[dict]) -> int:
    return sum(1 for r in rows or [] if str(r.get('扱い') or '').startswith(EXCLUDED_LABEL))


def validate_exclusion_selection(selected_rows: list[dict], reason: str) -> tuple[str | None, list[dict]]:
    """「不要（除外）」として登録する行の入力チェック。AIにしかない／金額が違う行が対象、理由は必須。"""
    rows = [r for r in selected_rows or [] if isinstance(r, dict)]
    if not rows:
        return '登録する行にチェックを付けてください', []
    if not _clean(reason):
        return '不要の理由を選んでください', []
    bad = [r for r in rows if r.get('区分') not in (DETAIL_KIND_AI_ONLY, DETAIL_KIND_DIFF)]
    if bad:
        return '「AIにしかない」または「金額が違う」の行だけを選んでください', []
    return None, rows


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
    counts[EXCLUDED_LABEL] = excluded_count(rows)
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
    excluded = counts.get(EXCLUDED_LABEL, 0)
    if excluded:
        parts.append(f"{EXCLUDED_LABEL} {excluded}")
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


def _alias_side_keys(name: Any, summary: Any, stored_keys: list | None) -> list[str]:
    """登録した側の照合キー。摘要がある行は『名称＋摘要』だけ、摘要が空なら『名称』だけを使う。

    2026-11-12 監査: 名称だけのキーも登録していたため、同じ品名で摘要だけ違う別の行にも
    「同じ商品」が効き、チェックした行とは別の行が突き合っていた。
    """
    from excel_diff import _normalize_name
    nm, sm = _clean(name), _clean(summary)
    if nm:
        return [_normalize_name(nm + sm)] if sm else [_normalize_name(nm)]
    return [k for k in (stored_keys or []) if k]


def alias_pairs_from_library(lib: list | None, project_id: str | None = None) -> set:
    """辞書エントリ一覧 → 突合に渡す frozenset の集合。

    摘要がある側は『名称＋摘要』キーだけを使い、同じ品名の別の行に効かないようにする。
    project_id を渡すと、他案件特有のエントリは除外する。
    """
    pairs: set = set()
    for e in lib or []:
        if not alias_applies(e, project_id):
            continue
        ai_keys = _alias_side_keys(e.get('ai_name'), e.get('ai_summary'), e.get('ai_keys'))
        ac_keys = _alias_side_keys(e.get('actual_name'), e.get('actual_summary'), e.get('actual_keys'))
        for a in ai_keys:
            for b in ac_keys:
                if a and b and a != b:
                    pairs.add(frozenset((a, b)))
        # 名称だけの組（頭に ~）。その種別に同じ名称の行が1つしか無いときだけ使う。
        # 摘要は物件ごとに変わる（CAD由来の建具の摘要など）ので、全案件共通の登録や対応表が
        # 他の物件でも効くようにする（2026-11-12 反証）
        from excel_diff import _normalize_name
        an, bn = _normalize_name(_clean(e.get('ai_name'))), _normalize_name(_clean(e.get('actual_name')))
        if an and bn and an != bn:
            pairs.add(frozenset(('~' + an, '~' + bn)))
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


# ---------------------------------------------------------------- 共有データの更新（複数人で使っても消えないように）
def _same_alias(a: dict, b: dict) -> bool:
    return (isinstance(a, dict) and isinstance(b, dict)
            and list(a.get('ai_keys') or []) == list(b.get('ai_keys') or [])
            and list(a.get('actual_keys') or []) == list(b.get('actual_keys') or []))


def alias_library_add(latest: list | None, entry: dict) -> tuple[list, bool]:
    """保存直前に読み直した最新の辞書へ1件足した新しい一覧と、足したかどうかを返す。

    2026-11-12 監査: 画面を開いた時点の古い一覧に足して丸ごと保存していたため、
    その間に別の担当者が登録した組が消えていた。
    """
    lib = [e for e in (latest or []) if isinstance(e, dict)]
    if any(_same_alias(e, entry) for e in lib):
        return lib, False
    return lib + [entry], True


def alias_library_remove(latest: list | None, entry: dict) -> list:
    """最新の辞書から、指定した組（並び順ではなく中身で特定）を取り除いた一覧を返す。"""
    return [e for e in (latest or []) if isinstance(e, dict) and not _same_alias(e, entry)]


def alias_library_set_scope(latest: list | None, entry: dict, scope: str, project_id: str) -> list:
    """最新の辞書の中で、指定した組の適用範囲だけを変えた一覧を返す。"""
    out = []
    for e in latest or []:
        if not isinstance(e, dict):
            continue
        if _same_alias(e, entry):
            e = {**e, 'scope': scope if scope in ALIAS_SCOPE_LABEL else ALIAS_SCOPE_GLOBAL,
                 'project_id': _clean(e.get('project_id')) or _clean(project_id)}
        out.append(e)
    return out


def diff_library_remove(lib: dict | None, keys, property_name: str) -> dict:
    """差異理由ライブラリから、この案件（property）が登録した理由を取り除いた新しい辞書を返す。

    差異理由を消しても他案件の「過去の理由（参考）」に残り続けていた（2026-11-12 監査）。
    """
    out: dict = {}
    drop = set(keys or [])
    for key, rec in (lib or {}).items():
        if key not in drop or not isinstance(rec, dict):
            out[key] = rec
            continue
        entries = [e for e in rec.get('entries', []) if not (isinstance(e, dict) and e.get('property') == property_name)]
        if entries:
            out[key] = {**rec, 'entries': entries}
    return out


# ---------------------------------------------------------------- 設計部の単価マスタを見積に使う（2026-11-12 監査）
DESIGN_PREVIEW_COLUMNS = [
    ('code', 'コード'), ('category', '工事種別'), ('group', '分類'), ('name', '名称'), ('summary', '摘要'),
    ('unit', '単位'), ('estimate_price', '見積単価（販価）'), ('order_price', '発注単価（設計原価）'),
    ('cost_kind', '材工区分'), ('breakdown_of', '内訳の親'), ('row_no', '行'),
]


def design_master_preview(rows: list | None, limit: int = 50) -> list[dict]:
    """取込プレビュー用に日本語の列名へ置き換えた行（英語の内部キーを画面に出さない）"""
    out = []
    for r in (rows or [])[:limit]:
        if isinstance(r, dict):
            out.append({jp: r.get(k, '') for k, jp in DESIGN_PREVIEW_COLUMNS})
    return out


def design_breakdown_rows(rows: list | None, code: Any) -> list[dict]:
    """親項目の内訳（内訳①②＝材・工の分離）の行。コードが無ければ空。"""
    code = _clean(code)
    if not code:
        return []
    return [r for r in rows or [] if isinstance(r, dict) and _clean(r.get('breakdown_of')) == code]


def search_design_master(rows: list | None, query: Any, limit: int = 20) -> list[dict]:
    """品名検索（部分一致）。内訳行は除く。見つからなければ空（呼び出し側で類似検索に回す）。"""
    q = unicodedata.normalize('NFKC', _clean(query)).lower().replace(' ', '').replace('　', '')
    if not q:
        return []
    out = []
    for r in rows or []:
        if not isinstance(r, dict) or _clean(r.get('breakdown_of')) or is_design_header_row(r):
            continue
        text = unicodedata.normalize('NFKC', f"{r.get('name', '')}{r.get('summary', '')}").lower()
        if q in text.replace(' ', '').replace('　', ''):
            out.append({'name': _clean(r.get('name')), 'summary': _clean(r.get('summary')),
                        'category': _clean(r.get('category')), 'unit': _clean(r.get('unit')),
                        'estimate_price': int(r.get('estimate_price') or 0), 'order_price': int(r.get('order_price') or 0),
                        'code': _clean(r.get('code')), 'cost_kind': _clean(r.get('cost_kind')), 'score': 1.0,
                        'reason': '品名に一致'})
            if len(out) >= limit:
                break
    return out


def design_item_to_new_item(sg: dict, cat_no: int, quantity: Any, tanka_rows: list | None = None,
                            design_rows: list | None = None) -> dict:
    """単価マスタの項目を見積の追加行（mods['new_items'] の形）にする。原価が無ければ粗利33%で入れる。

    ANDPADの材工分離が効くように、アプリの単価マスタの明細コード（andpad_code）か、
    設計部マスタの内訳①②（材・工）から作った振り分け（andpad_mapping）を付ける（2026-11-12 反証）。
    """
    try:
        qty = float(quantity)
    except (TypeError, ValueError):
        qty = 0.0
    if not qty or qty <= 0 or math.isnan(qty):
        qty = 1.0
    est = int(sg.get('estimate_price') or 0)
    cost = int(sg.get('order_price') or 0) or linked_order_price(0, 0, est)
    code = _clean(sg.get('code'))
    out = {'cat_no': int(cat_no), 'name': _clean(sg.get('name')), 'summary': _clean(sg.get('summary')),
           'quantity': qty, 'unit': _clean(sg.get('unit')) or '式', 'estimate_price': est, 'order_price': cost,
           'source': f"設計部単価マスタ {code}".strip()}
    hint = _andpad_hint(sg, tanka_rows, design_rows)
    if hint:
        out.update(hint)
    return out


def _andpad_hint(sg: dict, tanka_rows: list | None, design_rows: list | None) -> dict:
    """ANDPADの振り分けの手がかり。アプリの明細コード → 設計部の内訳①② → 材工区分 の順"""
    try:
        from andpad_converter import VENDOR_MAPPING
    except Exception:  # noqa: BLE001
        return {}
    code = _clean(sg.get('code'))
    if code.isdigit() and int(code) in VENDOR_MAPPING:
        return {'andpad_code': int(code)}
    if tanka_rows:
        try:
            import master_lookup
            best = next(iter(master_lookup.suggest_similar(_clean(sg.get('name')), _clean(sg.get('summary')),
                                                           tanka_rows, top_n=1)), None)
        except Exception:  # noqa: BLE001
            best = None
        if best and best.get('score', 0) >= 0.8 and _clean(best.get('code')).isdigit() \
                and int(best['code']) in VENDOR_MAPPING:
            return {'andpad_code': int(best['code'])}
    parts = design_breakdown_rows(design_rows, code) if design_rows else []
    if parts:
        def _amt(r):
            try:
                return max(int(r.get('order_price') or 0), 0)
            except (TypeError, ValueError):
                return 0
        total = sum(_amt(r) for r in parts)
        labor = sum(_amt(r) for r in parts
                    if '[工]' in _clean(r.get('name')) or _clean(r.get('cost_kind')) in ('工事費', '工'))
        if total > 0 and 0 < labor < total:
            return {'andpad_mapping': {'vendor': '', 'kind': '材',
                                       'parts': [{'vendor': '', 'kind': '工', 'ratio': labor / total, 'note': '施工費分'}]}}
        if total > 0 and labor == total:
            return {'andpad_mapping': {'vendor': '', 'kind': '工'}}
    if _clean(sg.get('cost_kind')) in ('工事費', '工'):
        return {'andpad_mapping': {'vendor': '', 'kind': '工'}}
    return {}


def validate_alias_pair(ai_name: Any, ai_summary: Any, actual_name: Any, actual_summary: Any) -> str | None:
    """「同じ商品」として登録してよい組か。両側が同じ表記なら登録しても意味がないので止める。"""
    from excel_diff import _normalize_name
    if not _clean(ai_name) or not _clean(actual_name):
        return 'AI側と実績側の名称を両方選んでください'
    if _normalize_name(_clean(ai_name) + _clean(ai_summary)) == _normalize_name(_clean(actual_name) + _clean(actual_summary)):
        return ('AI側と実績側が同じ表記なので、登録しても突き合わせは変わりません。'
                'AIの明細に無い項目なら「見積に追加」を使ってください')
    return None


def special_note_key(category: Any, name: Any, summary: Any = '') -> str:
    """案件特有・除外の登録キー。摘要があれば摘要まで含める（同じ品名の別の行に付かないように）"""
    c, n, sm = _clean(category), _clean(name), _clean(summary)
    return f"{c}|{n}|{sm}" if sm else f"{c}|{n}"


def parse_special_note_key(key: Any) -> tuple[str, str, str]:
    """登録キー → (工事種別, 品名, 摘要)。旧形式（工事種別|品名）は摘要が空"""
    parts = _clean(key).split('|', 2)
    while len(parts) < 3:
        parts.append('')
    return parts[0], parts[1], parts[2]


def withdrawn_note_keys(old_notes: dict | None, new_notes: dict | None) -> list[str]:
    """他案件の参考（差異理由ライブラリ）から取り下げるキー。

    消した行・理由の文を空にした行・「この案件のみ」に変えた行（2026-11-12 反証）。
    """
    old, new = old_notes or {}, new_notes or {}
    out = []
    for k, v in old.items():
        nv = new.get(k)
        if nv is None:
            out.append(k)
        elif isinstance(nv, dict) and not _clean(nv.get('reason')) and isinstance(v, dict) and _clean(v.get('reason')):
            out.append(k)
    for k, v in new.items():
        if isinstance(v, dict) and v.get('scope') == 'project' and k not in out:
            out.append(k)
    return out


# 設計部マスタの「種別」（可動棚・キッチン・サッシ…）から見積の工事種別を決めるための対応
_DESIGN_CATEGORY_HINTS = [
    (('キッチン', 'UB', 'ユニットバス', 'ﾕﾆｯﾄﾊﾞｽ', 'トイレ', '洗面', '給湯', '住宅設備', 'カップボード'), '住宅設備'),
    (('サッシ', '玄関', '窓', '外部建具', 'シャッター'), '外部建具'),
    (('可動棚', '棚', 'カウンター', '収納', '内部建具', 'ドア', '建具', '枕棚', 'クローゼット'), '内部建具'),
    (('クロス', '床', '内装', 'フローリング', 'タイル貼'), '内装'),
    (('外壁', 'サイディング', '外装'), '外装'),
    (('電気', '照明', 'コンセント'), '電気'),
    (('給排水', '配管'), '給排水'),
    (('基礎', '深基礎'), '基礎'),
    (('造作', '大工', '木工', '下地'), '木工'),
    (('左官', 'タイル'), '左官'),
    (('デッキ', 'ウッドデッキ'), 'デッキ'),
]


def design_category_no(sg: dict, categories: list, default: int = 21) -> int:
    """単価マスタの項目を見積のどの工事種別に入れるか（分類 → 種別 → 品名の手がかり → 雑工事）。

    本番の設計部マスタの「種別」は 可動棚・キッチン・サッシ などで工事種別名と一致しないため、
    追加先の初期値が『1. 仮設工事』になっていた（2026-11-12 反証）。
    """
    from excel_diff import _normalize_name
    pairs = [(int(no), _clean(name)) for no, name in (categories or []) if _clean(name)]
    if not pairs:
        return default

    def _find(text: str):
        t = _normalize_name(re.sub(r'^[0-9０-９]+[.．、\s]*', '', _clean(text)))
        if not t:
            return None
        for no, name in pairs:
            if _normalize_name(name) == t:
                return no
        for no, name in pairs:
            nn = _normalize_name(name)
            if nn and (t in nn or nn in t) and len(t) >= 2:
                return no
        return None
    for text in (sg.get('group'), sg.get('category')):
        no = _find(text or '')
        if no is not None:
            return no
    hay = f"{_clean(sg.get('group'))} {_clean(sg.get('category'))} {_clean(sg.get('name'))}"
    for words, target in _DESIGN_CATEGORY_HINTS:
        if any(w in hay for w in words):
            for no, name in pairs:
                if target in name:
                    return no
    return default if any(no == default for no, _ in pairs) else pairs[0][0]


def is_design_header_row(r: dict) -> bool:
    """販価0円で単位も無い行（見出し・別途見積の案内）。候補や検索に出さない"""
    try:
        price = int(r.get('estimate_price') or 0)
    except (TypeError, ValueError):
        price = 0
    return price == 0 and not _clean(r.get('unit'))


def design_vocabulary(rows: list | None) -> list:
    """類似候補に使う行（内訳行と0円の見出し行を除く）"""
    return [r for r in rows or [] if isinstance(r, dict) and not _clean(r.get('breakdown_of')) and not is_design_header_row(r)]
