"""設計部の単価マスタ取込・類似候補・対応表取込（純粋ロジック。Streamlit に依存しない）

実績比較で「AIにしかない」「実績にしかない」の突合率を上げるための裏方。

  - parse_price_master : 設計部の単価マスタ（CSV / xlsx）を {code, category, name, ...} の行に正規化
  - suggest_similar    : 実績明細の名称・摘要に似たマスタ項目を上位 N 件返す
  - parse_alias_table  : 図面表記⇔実績名称の対応表（CSV / xlsx）を「同じ商品」辞書エントリに変換
  - merge_alias_entries: 既存の「同じ商品」辞書へ重複なく追加

戻り値はすべて JSON にそのまま保存できるプリミティブ（str / int / list / dict）だけで構成する。
依存: openpyxl + 標準ライブラリ、excel_diff（名称の正規化・類似度）、ui_helpers（辞書エントリ）
"""
from __future__ import annotations

import csv
import io
import math
import os
import re
import unicodedata
from difflib import SequenceMatcher
from functools import lru_cache
from typing import Any, Callable

import openpyxl

from excel_diff import CATEGORY_MATCH_RATIO, _normalize_name, _similarity
from ui_helpers import ALIAS_SCOPE_GLOBAL, ALIAS_SCOPE_PROJECT, alias_entry_exists, make_alias_entry

# ---------------------------------------------------------------- 定数
HEADER_SEARCH_ROWS = 40          # 見出し行を探す最大行数
SUGGEST_MIN_SCORE = 0.35         # 類似候補として返す最低スコア（工事種別の加点前）
SUGGEST_CATEGORY_BONUS = 0.08    # 同じ工事種別の候補への加点
CATEGORY_CODE_MAX = 10000        # 工事コードがこれ未満なら工事種別（見出し）行とみなす

_TEXT_EXTENSIONS = ('.csv', '.txt', '.tsv')
_EXCEL_EXTENSIONS = ('.xlsx', '.xlsm')

# 列名のゆれ: (内部キー, 完全一致する見出し, 部分一致の正規表現)。
# 完全一致を全列で先に試し、残った列に部分一致を上から順に当てる（「発注単価」が「単価」に食われない）。
_MASTER_COLUMN_RULES: list[tuple[str, set[str], str]] = [
    ('order_price', {'発注単価', '原価', '仕入単価', '仕入価格', '発注価格', '原単価', '仕入値', '仕入', '発注',
                     '原価単価', '発注金額'}, r'原価|仕入|発注'),
    ('estimate_price', {'見積単価', '売単価', '販売単価', '単価', '売価', '見積価格', '販売価格', '提出単価',
                        '見積', '金額', '販価'}, r'販価|単価|売価|価格|金額'),
    ('unit', {'単位'}, r'^単位'),
    ('summary', {'摘要', '規格', '仕様', '備考', '規格仕様', '摘要規格', '仕様規格', '型番', '型式', '摘要備考'},
     r'摘要|規格|仕様|備考'),
    ('category', {'工事種別', '工種', '種別', 'カテゴリ', '工事区分', '工種名', '種別名'},
     r'工事種別|工種|種別|カテゴリ'),
    ('group', {'分類', '大分類', '区分'}, r'分類|区分'),
    ('cost_kind', {'工事費', '材工区分', '材工', '区分材工'}, r'^工事費$|材工'),
    ('name', {'名称', '品名', '工事名', '項目', '項目名', '品目', '名前', '内容', '工事項目', '工事内容', '商品名',
              '工事名称', '品名称', '名称品名', '品名名称'}, r'名称|品名|項目|工事名|品目'),
    ('code', {'工事コード', '工種コード', 'コード', 'code', 'no', '項目no', '番号', '項目番号', 'id'},
     r'コード|code|^no$|番号'),
]

_ALIAS_COLUMN_RULES: list[tuple[str, set[str], str]] = [
    ('ai_summary', {'ai摘要', '図面摘要', '図面規格', 'ai規格', 'cad摘要', 'ai仕様'}, r'ai摘要|図面摘要|cad摘要|ai規格'),
    ('actual_summary', {'実績摘要', 'マスタ摘要', 'マスター摘要', '実績規格', '見積摘要', 'マスタ規格'},
     r'実績摘要|マスタ摘要|マスター摘要|見積摘要|実績規格'),
    ('ai_name', {'図面表記', 'ai品名', 'ai名称', '図面名称', 'cad表記', '図面', 'ai', '図面の表記', 'cad名称',
                 '図面品名', 'ai側'}, r'図面|ai品名|ai名称|cad|^ai'),
    ('actual_name', {'実績品名', 'マスタ名', 'マスタ名称', 'マスター名', 'マスター名称', '実績名称', '実績',
                     'マスタ品名', '見積品名', '見積名称', '実績見積の名称', '実績側', 'マスタ'},
     r'実績|マスタ|マスター|見積'),
    ('category', {'工事種別', '工種', '種別', '分類', 'カテゴリ'}, r'工事種別|工種|種別|分類|カテゴリ'),
    ('scope', {'適用範囲', '範囲', 'スコープ', '適用', '共通'}, r'適用|範囲|スコープ'),
]

# アプリ同梱 tankamaster_updated.csv の採番のずれ:
#   22 管理費 は経費率で計算するため明細を持たず、旧22 屋外付帯工事の明細（220001〜）と
#   旧23 諸費用の明細（230001〜）はコードがそのまま残っている（estimate_calculator.py の
#   cat23/cat24 と同じ対応）。見出しがこの並びのときだけ接頭辞を読み替える。
_LEGACY_PREFIX_SHIFT = {'22': '23', '23': '24'}
_LEGACY_PREFIX_SIGNATURE = {'22': '管理費', '23': '屋外付帯工事', '24': '諸費用'}


# ---------------------------------------------------------------- セル・見出しの正規化
_RE_INVISIBLE = re.compile(r'[\u3164\u200b\u200c\u200d\u202a-\u202e\u2060\ufeff]')


def _cell_text(value: Any) -> str:
    """セル値を文字列化（数値セルの 10001.0 → '10001'、NaN/None → ''）。前後の空白・全角空白を除く。"""
    if value is None:
        return ''
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ''
        return str(int(value)) if value.is_integer() else str(value)
    if isinstance(value, int):
        return str(value)
    # 見えない文字（ハングル埋め草 U+3164・双方向制御 U+202A〜E・ゼロ幅空白・BOM）を除く
    return _RE_INVISIBLE.sub('', str(value)).strip(' \t\r\n　')


def _norm_header(value: Any) -> str:
    """見出しセルの比較用正規化。『見積単価(円)』→『見積単価』、『AI 品名』→『ai品名』。"""
    s = unicodedata.normalize('NFKC', _cell_text(value)).lower()
    s = re.sub(r'[\(\[【].*?[\)\]】]', '', s)
    s = re.sub(r'[\s_\-/・:.。、,※*]', '', s)
    return s


def _is_numeric_code(text: str) -> bool:
    return bool(re.fullmatch(r'\d+', text or ''))


def _parse_price(value: Any) -> tuple[int, bool]:
    """単価セル → (整数円, 読めたか)。空・'-' は 0 扱いで読めた扱い。"""
    if value is None:
        return 0, True
    if isinstance(value, bool):
        return 0, False
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return 0, True
        return int(round(value)), True
    s = unicodedata.normalize('NFKC', str(value)).strip()
    s = re.sub(r'[,\s円¥\\]', '', s)
    if s in ('', '-', '—', '―'):
        return 0, True
    try:
        return int(round(float(s))), True
    except ValueError:
        return 0, False


# ---------------------------------------------------------------- 表の読み込み（CSV / xlsx 共通）
def _decode_text(file_bytes: bytes) -> str:
    if file_bytes.startswith((b'\xff\xfe', b'\xfe\xff')):
        return file_bytes.decode('utf-16')
    for enc in ('utf-8-sig', 'cp932'):   # utf-8-sig は BOM 無しの utf-8 も読める。cp932 は最後
        try:
            return file_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode('utf-8', errors='replace')


def _read_csv_rows(file_bytes: bytes) -> list[list[Any]]:
    text = _decode_text(file_bytes)
    first = next((ln for ln in text.splitlines() if ln.strip()), '')
    delimiter = '\t' if first.count('\t') > first.count(',') else ','
    return [list(r) for r in csv.reader(io.StringIO(text, newline=''), delimiter=delimiter)]


def _read_xlsx_tables(file_bytes: bytes) -> list[list[list[Any]]]:
    """全シートを行リストにして、データ行の多い順（同数なら先頭から）に並べて返す。"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    try:
        tables = []
        for idx, ws in enumerate(wb.worksheets):
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            n = sum(1 for r in rows if any(_cell_text(c) for c in r))
            tables.append((-n, idx, rows))
    finally:
        wb.close()
    tables.sort(key=lambda t: (t[0], t[1]))
    return [t[2] for t in tables]


def _load_tables(file_bytes: bytes, filename: str) -> tuple[list[list[list[Any]]], list[str]]:
    """ファイル → 候補の表（xlsx は複数シート）。読めなければ ([], [理由])。"""
    if not file_bytes:
        return [], ['ファイルが空です']
    ext = os.path.splitext(filename or '')[1].lower()
    if not ext:
        ext = '.xlsx' if file_bytes.startswith(b'PK') else '.csv'
    if ext in _EXCEL_EXTENSIONS:
        try:
            return _read_xlsx_tables(file_bytes), []
        except Exception as exc:  # noqa: BLE001 - 壊れたファイルは理由を返して終わる
            return [], [f'Excel ファイルとして読めませんでした: {exc}']
    if ext in _TEXT_EXTENSIONS:
        try:
            return [_read_csv_rows(file_bytes)], []
        except (csv.Error, UnicodeError) as exc:
            return [], [f'CSV として読めませんでした: {exc}']
    return [], [f'対応していないファイル形式です（CSV または Excel .xlsx を選んでください）: {filename}']


_PREFER_HEADER_RE = re.compile(r'33')   # 30%/33% 販価が並ぶマスタでは 33%（標準の粗利率）を採用


def _prefer_33(cells: list[Any], cols: dict[str, int], key: str, rules: list[tuple[str, set[str], str]]) -> None:
    """key に割り当てた列以外にも同じ規則に合う列があれば、見出しに「33」を含む列を優先する"""
    if key not in cols:
        return
    rule = next((r for r in rules if r[0] == key), None)
    if rule is None:
        return
    exact, partial = rule[1], re.compile(rule[2])
    cands = [i for i, c in enumerate(cells)
             if _norm_header(c) in exact or partial.search(_norm_header(c))]
    pref = [i for i in cands if _PREFER_HEADER_RE.search(_norm_header(cells[i]))]
    if pref and cols[key] not in pref:
        cols[key] = pref[0]


def _assign_columns(cells: list[Any], rules: list[tuple[str, set[str], str]]) -> dict[str, int]:
    headers = [_norm_header(c) for c in cells]
    mapping: dict[str, int] = {}
    used: set[int] = set()
    for key, exact, _rx in rules:                      # 1) 完全一致
        for idx, h in enumerate(headers):
            if h and idx not in used and h in exact:
                mapping[key], _ = idx, used.add(idx)
                break
    for key, _exact, rx in rules:                      # 2) 残った列に部分一致
        if key in mapping:
            continue
        for idx, h in enumerate(headers):
            if h and idx not in used and re.search(rx, h):
                mapping[key], _ = idx, used.add(idx)
                break
    return mapping


def _find_header(rows: list[list[Any]], rules: list[tuple[str, set[str], str]],
                 is_header: Callable[[dict[str, int]], bool]) -> tuple[int | None, dict[str, int]]:
    for i, cells in enumerate(rows[:HEADER_SEARCH_ROWS]):
        mapping = _assign_columns(cells, rules)
        if is_header(mapping):
            return i, mapping
    return None, {}


def _getter(cells: list[Any], cols: dict[str, int]) -> Callable[[str], str]:
    def get(key: str) -> str:
        idx = cols.get(key)
        if idx is None or idx >= len(cells):
            return ''
        return _cell_text(cells[idx])
    return get


# ---------------------------------------------------------------- 1. 単価マスタの取込
def _is_category_header(name: str, summary: str, unit: str, code: str, order_price: int, estimate_price: int) -> bool:
    """『1,仮設工事,,式,0,0』のような工事種別の見出し行か（摘要・単価が無く名称だけ）。"""
    if summary or order_price or estimate_price or name.startswith('※'):
        return False
    if code:
        return _is_numeric_code(code) and int(code) < CATEGORY_CODE_MAX
    return unit in ('', '式')


def _category_for_code(code: str, cat_by_code: dict[str, str]) -> str:
    """工事コードの上位桁（10001 → 1）から工事種別を引く。引けなければ ''。"""
    if not (_is_numeric_code(code) and int(code) >= CATEGORY_CODE_MAX and cat_by_code):
        return ''
    prefix = str(int(code) // CATEGORY_CODE_MAX)
    legacy = all(cat_by_code.get(k) == v for k, v in _LEGACY_PREFIX_SIGNATURE.items())
    if legacy and prefix in _LEGACY_PREFIX_SHIFT:
        prefix = _LEGACY_PREFIX_SHIFT[prefix]
    return cat_by_code.get(prefix, '')


def _parse_master_table(rows: list[list[Any]], errors: list[str]) -> list[dict] | None:
    """1つの表から明細行を取り出す。見出しが無ければ None（別シートを試す合図）。"""
    header_idx, cols = _find_header(
        rows, _MASTER_COLUMN_RULES,
        lambda m: 'name' in m and ('estimate_price' in m or 'order_price' in m))
    if header_idx is None:
        return None
    _prefer_33(rows[header_idx], cols, 'estimate_price', _MASTER_COLUMN_RULES)
    out: list[dict] = []
    current_category = ''
    current_group = ''
    cat_by_code: dict[str, str] = {}
    last_parent_code = ''   # 内訳行（内訳①…）の親コード
    for i in range(header_idx + 1, len(rows)):
        cells, row_no = rows[i], i + 1
        get = _getter(cells, cols)
        name, summary, unit, code = get('name'), get('summary'), get('unit'), get('code')
        cat_cell = get('category')
        group_cell = get('group')
        cost_kind = get('cost_kind')
        raw_order, raw_est = get('order_price'), get('estimate_price')
        order_price, ok_order = _parse_price(raw_order)
        estimate_price, ok_est = _parse_price(raw_est)

        if group_cell and not _is_numeric_code(group_cell):
            current_group = group_cell
        explicit_category = ''
        if cat_cell:
            if _is_numeric_code(cat_cell) and not code:
                code = cat_cell                      # 「工種」列が番号だけのマスタ
            elif not _is_numeric_code(cat_cell):
                explicit_category = cat_cell
                current_category = cat_cell          # 空欄は上の行の種別を引き継ぐ
        elif group_cell and not _is_numeric_code(group_cell) and 'group' in cols and 'category' in cols:
            # 「種別」が空の行は「分類」を種別として使う
            explicit_category = group_cell
            current_category = group_cell

        if not name:
            if summary or raw_order or raw_est or code:
                errors.append(f'{row_no}行目: 名称が空のため読み飛ばしました')
            continue
        if not explicit_category and _is_category_header(name, summary, unit, code, order_price, estimate_price):
            current_category = name
            if code:
                cat_by_code[code] = name
            continue
        if not ok_est:
            errors.append(f'{row_no}行目「{name}」: 見積単価『{raw_est}』を数値として読めません（0円として取り込みました）')
        if not ok_order:
            errors.append(f'{row_no}行目「{name}」: 発注単価『{raw_order}』を数値として読めません（0円として取り込みました）')
        category = explicit_category or _category_for_code(code, cat_by_code) or current_category
        # 内訳行（「内訳①: [材] …」や親コード + '-1'）は親項目の材／工の内訳。
        # 候補の語彙には親だけを使うため breakdown_of に親コードを残す
        breakdown_of = ''
        if re.match(r'^\s*内訳', name) or (last_parent_code and code and re.fullmatch(re.escape(last_parent_code) + r'-\d+', code)):
            breakdown_of = last_parent_code
        else:
            last_parent_code = code or last_parent_code
        out.append({
            'code': code, 'category': category, 'group': current_group, 'name': name, 'summary': summary,
            'unit': unit, 'order_price': order_price, 'estimate_price': estimate_price, 'row_no': row_no,
            'cost_kind': cost_kind, 'breakdown_of': breakdown_of,
        })
    return out


def parse_price_master(file_bytes: bytes, filename: str) -> tuple[list[dict], list[str]]:
    """設計部の単価マスタ（CSV: cp932/utf-8/utf-8-sig 自動判定、xlsx: データ行の多いシート）を読む。

    見出し行は「名称/品名/工事名…」と「単価」系の列がある行を自動検出し、列名のゆれ
    （名称/品名/項目、摘要/規格/仕様、見積単価/売単価/単価、発注単価/原価/仕入単価、
    工事コード/工種/工事種別）を吸収する。

    工事種別は次の優先順で決める:
      1. 「工事種別」列に文字で入っていればその値（空欄は上の行を引き継ぐ）
      2. 工事コードの上位桁 → 先に出てきた種別見出し行（『1,仮設工事,,式,0,0』）の名称
      3. 直前に出てきた種別見出し行の名称
    種別見出し行そのものは rows に含めない。

    Returns:
        (rows, errors)
        rows: [{code, category, name, summary, unit, order_price(int), estimate_price(int), row_no}]
        errors: 読めなかった行・ファイルの説明（利用者向けの日本語）
    """
    tables, errors = _load_tables(file_bytes, filename)
    if not tables:
        return [], errors
    for rows in tables:
        parsed = _parse_master_table(rows, errors)
        if parsed is not None:
            return parsed, errors
    errors.append('見出し行が見つかりません（「名称」と「単価」の列がある行が必要です）')
    return [], errors


def master_categories(rows: list[dict] | None) -> list[str]:
    """取り込んだマスタ行の工事種別を出現順に（重複なく）返す。"""
    seen: list[str] = []
    for r in rows or []:
        c = _cell_text((r or {}).get('category')) if isinstance(r, dict) else ''
        if c and c not in seen:
            seen.append(c)
    return seen


# ---------------------------------------------------------------- 2. 類似候補
_TOKEN_RE = re.compile(r'[a-z]+|[぀-ヿ㐀-鿿豈-﫿]+')


def _tokens(*texts: Any) -> set[str]:
    """NFKC 後に数字・記号・空白で切った語の集合（1文字のひらがな＝助詞は捨てる）。"""
    toks: set[str] = set()
    for t in texts:
        s = unicodedata.normalize('NFKC', str(t or '')).lower().replace('・', ' ')
        for tok in _TOKEN_RE.findall(s):
            if len(tok) == 1 and '぀' <= tok <= 'ゟ':
                continue
            toks.add(tok)
    return toks


def _bigrams(s: str) -> frozenset[str]:
    if len(s) < 2:
        return frozenset({s}) if s else frozenset()
    return frozenset(s[i:i + 2] for i in range(len(s) - 1))


def _dice(a: frozenset[str], b: frozenset[str]) -> float:
    return 2 * len(a & b) / (len(a) + len(b)) if a and b else 0.0


_GENERIC_WORDS = ('工事', '一式')   # どの明細にも付く語。残すと「ガス工事」が「ｸﾛｽ工事」に寄ってしまう


def _strip_generic(key: str) -> str:
    """正規化キーから『工事』『一式』を除く（除くと空になる場合は元のまま）。"""
    stripped = key
    for w in _GENERIC_WORDS:
        stripped = stripped.replace(w, '')
    return stripped or key


def _strip_generic_text(text: str) -> str:
    out = text
    for w in _GENERIC_WORDS:
        out = out.replace(w, ' ')
    return out if out.strip() else text


@lru_cache(maxsize=8192)
def _prep(name: str, summary: str) -> tuple[str, str, str, frozenset[str], frozenset[str], frozenset[str]]:
    """(名称キー, 名称＋摘要キー, 摘要キー, 名称bigram, 名称＋摘要bigram, 語集合)。

    excel_diff._match_keys と同じ正規化のうえで『工事』『一式』を除いて比べる。
    """
    name_key = _strip_generic(_normalize_name(name))
    full_key = _strip_generic(_normalize_name(name + summary))
    sum_key = _strip_generic(_normalize_name(summary))
    tokens = frozenset(_tokens(_strip_generic_text(name), _strip_generic_text(summary)))
    return name_key, full_key, sum_key, _bigrams(name_key), _bigrams(full_key), tokens


def _token_weights(preps: list[tuple]) -> dict[str, float]:
    """語の重み（マスタ内で珍しい語ほど重い＝ log(1 + 行数/出現行数)）。『材工』『予算組』『LIXIL』の
    ような多くの行に付く語だけで候補に挙がるのを防ぐ。マスタに無い語は最大の重み。"""
    n = len(preps)
    df: dict[str, int] = {}
    for p in preps:
        for t in p[5]:
            df[t] = df.get(t, 0) + 1
    return {t: math.log(1 + n / c) for t, c in df.items()}


def _same_category(a: str, b: str) -> bool:
    ka, kb = _normalize_name(a), _normalize_name(b)
    if not ka or not kb:
        return False
    return ka == kb or SequenceMatcher(None, ka, kb).ratio() >= CATEGORY_MATCH_RATIO


def _score(q: tuple, m: tuple, weights: dict[str, float], max_w: float) -> tuple[float, str]:
    """(0〜1 のスコア, 根拠文)。difflib 類似度 0.6 + 文字bigram 0.25 + 語の重なり（重み付き）0.15。"""
    q_name, q_full, q_sum, q_nb, q_fb, q_tok = q
    m_name, m_full, _m_sum, m_nb, m_fb, m_tok = m
    sim = max(_similarity(q_name, m_name), _similarity(q_full, m_full),
              _similarity(q_name, m_full), _similarity(q_full, m_name))
    if q_sum:   # 実績側が「名称=1階床材変更 / 摘要=1階床 無垢ﾌﾛｰﾘﾝｸﾞ」と摘要に品名を書く場合
        sim = max(sim, 0.95 * max(_similarity(q_sum, m_name), _similarity(q_sum, m_full)))
    dice = max(_dice(q_fb, m_fb), _dice(q_nb, m_nb))
    common = q_tok & m_tok
    if q_tok and m_tok:
        w_common = sum(weights.get(t, max_w) for t in common)
        w_q = sum(weights.get(t, max_w) for t in q_tok)
        w_m = sum(weights.get(t, max_w) for t in m_tok)
        tok = w_common / min(w_q, w_m) if min(w_q, w_m) > 0 else 0.0
    else:
        tok = 0.0
    score = 0.6 * sim + 0.25 * dice + 0.15 * tok

    if q_name and q_name == m_name:
        reason = '名称が一致'
    elif q_full == m_full or q_name == m_full or q_full == m_name:
        reason = '名称＋摘要が一致'
    elif q_sum and q_sum in (m_name, m_full):
        reason = '摘要がマスタ名称と一致'
    elif sim >= 0.9:
        reason = '名称を含む'
    elif common:
        top = sorted(common, key=lambda t: -weights.get(t, max_w))[:3]
        reason = '共通語: ' + '、'.join(top)
    else:
        reason = f'文字の並びが近い（{int(round(sim * 100))}%）'
    return min(1.0, score), reason


def suggest_similar(name: str, summary: str, master_rows: list[dict], top_n: int = 5,
                    category: str | None = None) -> list[dict]:
    """実績明細の名称・摘要に似た単価マスタ項目を、スコアの高い順に top_n 件返す。

    スコア = difflib 類似度（名称／名称＋摘要の組み合わせの最大）0.6
           + 文字 bigram の Dice 係数 0.25 + 語（NFKC 後に数字・記号で分割）の重なり 0.15。
    category を渡すと同じ工事種別の候補に +SUGGEST_CATEGORY_BONUS（他種別も候補に残る）。
    加点前のスコアが SUGGEST_MIN_SCORE 未満の候補は返さない。

    Returns:
        [{name, summary, category, unit, estimate_price, score(0〜1), reason}] score 降順
    """
    master_rows = [r for r in (master_rows or []) if isinstance(r, dict) and not r.get('breakdown_of')]
    name, summary = _cell_text(name), _cell_text(summary)
    if not (name or summary) or not master_rows or top_n <= 0:
        return []
    q = _prep(name, summary)
    preps: list[tuple[int, dict, tuple]] = []
    seen: set[tuple[str, str, str]] = set()
    for order, row in enumerate(master_rows):
        if not isinstance(row, dict):
            continue
        m_name, m_sum, m_cat = _cell_text(row.get('name')), _cell_text(row.get('summary')), _cell_text(row.get('category'))
        if not m_name or (m_name, m_sum, m_cat) in seen:
            continue
        seen.add((m_name, m_sum, m_cat))
        preps.append((order, row, _prep(m_name, m_sum)))
    weights = _token_weights([p[2] for p in preps])
    max_w = math.log(1 + len(preps)) if preps else 1.0
    cands: list[tuple[float, int, int, dict, str]] = []
    for order, row, mp in preps:
        base, reason = _score(q, mp, weights, max_w)
        if base < SUGGEST_MIN_SCORE:
            continue
        same_cat = bool(category) and _same_category(category, _cell_text(row.get('category')))
        score = min(1.0, base + SUGGEST_CATEGORY_BONUS) if same_cat else base
        if same_cat:
            reason += '・同じ工事種別'
        cands.append((score, 0 if same_cat else 1, order, row, reason))
    cands.sort(key=lambda c: (-c[0], c[1], c[2]))
    out = []
    for score, _flag, _order, row, reason in cands[:top_n]:
        price, _ = _parse_price(row.get('estimate_price'))
        out.append({
            'name': _cell_text(row.get('name')), 'summary': _cell_text(row.get('summary')),
            'category': _cell_text(row.get('category')), 'unit': _cell_text(row.get('unit')),
            'estimate_price': price, 'score': round(score, 3), 'reason': reason,
        })
    return out


# ---------------------------------------------------------------- 3. 対応表（図面表記 ⇔ 実績名称）の取込
_SCOPE_GLOBAL_WORDS = {'', '全案件共通', '全案件', '共通', 'global', '全体', '全て', 'すべて', '汎用', '普遍', '共通項目'}
_SCOPE_PROJECT_WORDS = {'この案件のみ', '案件のみ', 'この案件', 'project', '案件特有', 'この案件特有', '案件限定', '個別',
                        '案件固有', 'この案件だけ'}


def parse_scope_label(value: Any) -> str | None:
    """対応表の「適用範囲」セル → ALIAS_SCOPE_GLOBAL / ALIAS_SCOPE_PROJECT。読めなければ None。空は全案件共通。"""
    s = unicodedata.normalize('NFKC', _cell_text(value)).lower().replace(' ', '')
    if s in _SCOPE_GLOBAL_WORDS:
        return ALIAS_SCOPE_GLOBAL
    if s in _SCOPE_PROJECT_WORDS:
        return ALIAS_SCOPE_PROJECT
    if '案件' in s and any(k in s for k in ('のみ', '特有', '限定', '個別', '固有', 'だけ')):
        return ALIAS_SCOPE_PROJECT
    if '共通' in s or '全案件' in s:
        return ALIAS_SCOPE_GLOBAL
    return None


def _parse_alias_rows(rows: list[list[Any]], property_name: str, project_id: str, at: str, by: str,
                      errors: list[str]) -> list[dict] | None:
    header_idx, cols = _find_header(rows, _ALIAS_COLUMN_RULES, lambda m: 'ai_name' in m and 'actual_name' in m)
    if header_idx is None:
        return None
    out: list[dict] = []
    for i in range(header_idx + 1, len(rows)):
        cells, row_no = rows[i], i + 1
        get = _getter(cells, cols)
        ai_name, ai_sum = get('ai_name'), get('ai_summary')
        ac_name, ac_sum = get('actual_name'), get('actual_summary')
        cat, scope_raw = get('category'), get('scope')
        if not any((ai_name, ai_sum, ac_name, ac_sum, cat, scope_raw)):
            continue                                        # 空行は黙って飛ばす
        if not ai_name and not ac_name:
            errors.append(f'{row_no}行目: 図面表記（AI品名）と実績品名の両方が空のためスキップ')
            continue
        if not ai_name:
            errors.append(f'{row_no}行目「{ac_name}」: 図面表記（AI品名）が空のためスキップ')
            continue
        if not ac_name:
            errors.append(f'{row_no}行目「{ai_name}」: 実績品名が空のためスキップ')
            continue
        scope = parse_scope_label(scope_raw)
        if scope is None:
            errors.append(f'{row_no}行目「{ai_name}」: 適用範囲『{scope_raw}』が読めません'
                          f'（「全案件共通」か「この案件のみ」を入れてください）のためスキップ')
            continue
        if scope == ALIAS_SCOPE_PROJECT and not _cell_text(project_id):
            errors.append(f'{row_no}行目「{ai_name}」: 「この案件のみ」ですが案件が選ばれていないためスキップ')
            continue
        entry = make_alias_entry({'AI品名': ai_name, 'AI摘要': ai_sum}, {'実績品名': ac_name, '実績摘要': ac_sum},
                                 cat, property_name, at, by=by, scope=scope, project_id=project_id)
        if entry['ai_keys'] == entry['actual_keys']:
            errors.append(f'{row_no}行目「{ai_name}」: 両側が同じ表記のためスキップ（登録不要）')
            continue
        if alias_entry_exists(out, entry):
            errors.append(f'{row_no}行目「{ai_name}」⇔「{ac_name}」: 同じ内容の行が先にあるためスキップ')
            continue
        out.append(entry)
    return out


def parse_alias_table(file_bytes: bytes, filename: str, property_name: str, project_id: str, at: str,
                      by: str = '') -> tuple[list[dict], list[str]]:
    """対応表（CSV / xlsx）を「同じ商品」辞書エントリ（ui_helpers.make_alias_entry の形）に変換する。

    列（名前のゆれは吸収・見出し行は自動検出）:
      図面表記 または AI品名   … AI見積側の表記（必須）
      実績品名 または マスタ名 … 実績見積側の表記（必須）
      AI摘要 / 実績摘要 / 工事種別 … 任意
      適用範囲 … 「全案件共通」か「この案件のみ」。空なら全案件共通

    空行は無視、片側が空・適用範囲が読めない行はスキップして errors に行番号付きで理由を返す。
    同じファイル内の重複行は 1 件にまとめる（既存辞書との重複は merge_alias_entries で除く）。
    """
    tables, errors = _load_tables(file_bytes, filename)
    if not tables:
        return [], errors
    for rows in tables:
        parsed = _parse_alias_rows(rows, _cell_text(property_name), _cell_text(project_id), _cell_text(at),
                                   _cell_text(by), errors)
        if parsed is not None:
            return parsed, errors
    errors.append('見出し行が見つかりません（「図面表記（またはAI品名）」と「実績品名（またはマスタ名）」の列が必要です）')
    return [], errors


def merge_alias_entries(lib: list | None, entries: list | None) -> tuple[list, int]:
    """既存の「同じ商品」辞書に entries を重複なく追加した新しいリストと追加件数を返す（元のリストは変えない）。"""
    new_lib = list(lib or [])
    added = 0
    for e in entries or []:
        if not isinstance(e, dict) or alias_entry_exists(new_lib, e):
            continue
        new_lib.append(e)
        added += 1
    return new_lib, added
