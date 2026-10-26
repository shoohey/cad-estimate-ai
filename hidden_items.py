"""図面に現れない項目（機器・工事）の統計・パターン登録・フィードバック学習（純粋ロジック）

CAD 数量からは拾えないが実績見積には必ず載る項目（例: エコキュート、仮設電気、
浄化槽、外構の基本工事など）を、過去物件の実績比較（excel_diff.diff_estimates の
``actual_only``）から統計的に抽出し、概算見積へ追加候補として提示するための計算。

Streamlit・DB には依存しない。app 側は

  1. 実績比較済みの物件から ``samples`` を組み立てて ``derive_hidden_item_stats``
  2. 保存済みフィードバックを ``apply_feedback`` で重ね
  3. ``build_candidates`` でしきい値・登録パターンから追加候補を作り
  4. ``to_new_items`` で mods['new_items'] 形式に変換して見積へ反映
  5. ユーザーの「必要／不要」判定を ``record_judgments`` で蓄積

という流れで使う。全関数は入力を変更せず、新しいオブジェクトを返す。
"""
from __future__ import annotations

import hashlib
import unicodedata
from collections import Counter
from statistics import median
from typing import Any

from excel_diff import _normalize_name
from ui_helpers import linked_order_price

MISC_CATEGORY_NO = 21          # 種別が特定できない項目は雑工事へ
SOURCE_STATS = '統計'
SOURCE_PATTERN_PREFIX = 'パターン: '
JUDGMENT_NEED = '必要'
JUDGMENT_UNNEED = '不要'

_ITEM_KEYS = ('key', 'category', 'name', 'summary', 'quantity', 'unit', 'estimate_price', 'order_price')


# --- 内部ヘルパー -------------------------------------------------------

def _text(value: Any) -> str:
    return str(value).strip() if value is not None else ''


def _to_float(value: Any, default: float = 0.0) -> float:
    if isinstance(value, bool):
        return default
    try:
        f = float(value)
    except (TypeError, ValueError):
        return default
    if f != f:   # NaN
        return default
    return f


def _to_int(value: Any, default: int = 0) -> int:
    if isinstance(value, bool):
        return default
    try:
        f = float(value)
    except (TypeError, ValueError):
        return default
    if f != f:
        return default
    return int(round(f))


def _nonneg_int(value: Any) -> int:
    return max(_to_int(value), 0)


def make_key(category: Any, name: Any) -> str:
    """統計・パターン・フィードバックを結ぶキー: '種別|正規化名'"""
    return f"{_text(category)}|{_normalize_name(name)}"


def _most_common(values: list[str]) -> str:
    """最頻の表記（同数なら最初に現れたもの）。空文字は候補から外す。"""
    filtered = [v for v in values if v]
    if not filtered:
        return ''
    return Counter(filtered).most_common(1)[0][0]


def _quantity_or_one(value: Any) -> float:
    q = _to_float(value)
    return q if q > 0 else 1.0


def _normalize_item(raw: dict) -> dict | None:
    """パターン行を正規化する。名称が空なら None。"""
    if not isinstance(raw, dict):
        return None
    name = _text(raw.get('name'))
    if not name:
        return None
    category = _text(raw.get('category'))
    key = _text(raw.get('key')) or make_key(category, name)
    estimate_price = _nonneg_int(raw.get('estimate_price'))
    order_raw = raw.get('order_price')
    order_price = (linked_order_price(0, 0, estimate_price) if order_raw is None
                   else _nonneg_int(order_raw))
    return {
        'key': key,
        'category': category,
        'name': name,
        'summary': _text(raw.get('summary')),
        'quantity': _quantity_or_one(raw.get('quantity')),
        'unit': _text(raw.get('unit')),
        'estimate_price': estimate_price,
        'order_price': order_price,
    }


# --- 1. 統計 ------------------------------------------------------------

def derive_hidden_item_stats(samples: list[dict]) -> list[dict]:
    """過去物件の actual_only 明細を (種別, 正規化名) で集約し、出現率と中央値を出す。

    samples: [{'property': str,
               'actual_only': [{'category', 'name', 'summary', 'quantity', 'unit', 'amount'}, ...]}, ...]

    戻り値（prob 降順 → median_amount 降順 → key 昇順）:
      {key, category, name, summary, unit, n_properties, freq, prob,
       median_amount, median_quantity, properties}

    同一物件内に同じ項目が複数行あっても出現は1回と数え、金額・数量はその物件内で合算する。
    """
    if not samples:
        return []
    n_properties = len(samples)
    groups: dict[str, dict] = {}

    for idx, sample in enumerate(samples):
        if not isinstance(sample, dict):
            continue
        prop_name = _text(sample.get('property')) or f'物件{idx + 1}'
        seen_in_property: dict[str, dict] = {}
        for raw in sample.get('actual_only') or []:
            if not isinstance(raw, dict):
                continue
            name = _text(raw.get('name'))
            if not name:
                continue
            category = _text(raw.get('category'))
            key = make_key(category, name)
            g = groups.setdefault(key, {
                'category': category, 'names': [], 'summaries': [], 'units': [],
                'amounts': [], 'quantities': [], 'properties': [],
            })
            g['names'].append(name)
            g['summaries'].append(_text(raw.get('summary')))
            g['units'].append(_text(raw.get('unit')))
            acc = seen_in_property.get(key)
            if acc is None:
                acc = {'amount': 0, 'quantity': 0.0}
                seen_in_property[key] = acc
                g['properties'].append(prop_name)
            acc['amount'] += _to_int(raw.get('amount'))
            acc['quantity'] += _to_float(raw.get('quantity'))
        for key, acc in seen_in_property.items():
            groups[key]['amounts'].append(acc['amount'])
            groups[key]['quantities'].append(acc['quantity'])

    stats = []
    for key, g in groups.items():
        freq = len(g['properties'])
        stats.append({
            'key': key,
            'category': g['category'],
            'name': _most_common(g['names']),
            'summary': _most_common(g['summaries']),
            'unit': _most_common(g['units']),
            'n_properties': n_properties,
            'freq': freq,
            'prob': freq / n_properties if n_properties else 0.0,
            'median_amount': int(round(median(g['amounts']))) if g['amounts'] else 0,
            'median_quantity': float(median(g['quantities'])) if g['quantities'] else 0.0,
            'properties': list(g['properties']),
        })
    stats.sort(key=lambda s: (-s['prob'], -s['median_amount'], s['key']))
    return stats


# --- 2. フィードバック -------------------------------------------------

def apply_feedback(stats: list[dict], feedback: dict) -> list[dict]:
    """「必要／不要」の判定件数を統計に重ねて prob_adj を付けた新しいリストを返す。

    feedback: {key: {'need': int, 'unneed': int}}
    prob_adj = (freq + need) / (n_properties + need + unneed) を 0〜1 に丸める。
    n_properties が 0 の統計は need / (need + unneed)（両方 0 なら 0）。
    """
    feedback = feedback if isinstance(feedback, dict) else {}
    result = []
    for stat in stats or []:
        if not isinstance(stat, dict):
            continue
        fb = feedback.get(stat.get('key'))
        fb = fb if isinstance(fb, dict) else {}
        need = _nonneg_int(fb.get('need'))
        unneed = _nonneg_int(fb.get('unneed'))
        freq = _nonneg_int(stat.get('freq'))
        n_props = _nonneg_int(stat.get('n_properties'))
        denominator = n_props + need + unneed
        if n_props > 0:
            prob_adj = (freq + need) / denominator if denominator else 0.0
        else:
            prob_adj = need / (need + unneed) if (need + unneed) else 0.0
        prob_adj = min(max(prob_adj, 0.0), 1.0)
        new_stat = dict(stat)
        new_stat['properties'] = list(stat.get('properties') or [])
        new_stat.update({'need': need, 'unneed': unneed, 'prob_adj': prob_adj})
        result.append(new_stat)
    return result


def record_judgments(feedback: dict, judgments: list[dict]) -> dict:
    """判定（'必要' / '不要'）をカウントに加算した新しい feedback を返す。元は変更しない。

    judgments: [{'key': str, 'judgment': '必要' | '不要'}, ...]  未判定・不正値は無視。
    """
    result: dict = {}
    if isinstance(feedback, dict):
        for key, counts in feedback.items():
            counts = counts if isinstance(counts, dict) else {}
            result[key] = {'need': _nonneg_int(counts.get('need')),
                           'unneed': _nonneg_int(counts.get('unneed'))}
    for j in judgments or []:
        if not isinstance(j, dict):
            continue
        key = _text(j.get('key'))
        judgment = _text(j.get('judgment'))
        if not key or judgment not in (JUDGMENT_NEED, JUDGMENT_UNNEED):
            continue
        counts = result.setdefault(key, {'need': 0, 'unneed': 0})
        counts['need' if judgment == JUDGMENT_NEED else 'unneed'] += 1
    return result


def summarize_feedback(stats_adj: list[dict]) -> dict:
    """画面用の小さな集計: 統計項目数・必要／不要の累計・必要率。"""
    rows = [s for s in (stats_adj or []) if isinstance(s, dict)]
    n_need = sum(_nonneg_int(s.get('need')) for s in rows)
    n_unneed = sum(_nonneg_int(s.get('unneed')) for s in rows)
    total = n_need + n_unneed
    return {
        'n_items': len(rows),
        'n_need': n_need,
        'n_unneed': n_unneed,
        'need_rate': n_need / total if total else 0.0,
    }


# --- 3. 候補の組み立て ----------------------------------------------------

def _candidate_from_stat(stat: dict) -> dict:
    quantity = _quantity_or_one(stat.get('median_quantity'))
    median_amount = _nonneg_int(stat.get('median_amount'))
    estimate_price = int(round(median_amount / quantity)) if quantity > 0 else median_amount
    prob = stat.get('prob_adj', stat.get('prob', 0.0))
    return {
        'key': _text(stat.get('key')),
        'category': _text(stat.get('category')),
        'name': _text(stat.get('name')),
        'summary': _text(stat.get('summary')),
        'unit': _text(stat.get('unit')),
        'quantity': quantity,
        'estimate_price': estimate_price,
        'order_price': linked_order_price(0, 0, estimate_price),
        'prob': min(max(_to_float(prob), 0.0), 1.0),
        'source': SOURCE_STATS,
    }


def _candidate_from_pattern_item(item: dict, pattern_name: str) -> dict | None:
    norm = _normalize_item(item)
    if norm is None:
        return None
    if norm['order_price'] <= 0 < norm['estimate_price']:
        norm['order_price'] = linked_order_price(0, 0, norm['estimate_price'])
    norm['prob'] = 1.0
    norm['source'] = f'{SOURCE_PATTERN_PREFIX}{pattern_name}'
    return norm


def build_candidates(stats_adj: list[dict], threshold: float, patterns: list[dict]) -> list[dict]:
    """統計（prob_adj >= threshold）と登録パターンから追加候補を組み立てる。

    同じ key が統計とパターンの両方にあればパターンを優先し1件にまとめる。
    並び: パターン由来（登録順）→ 統計由来（stats_adj の順）。
    戻り値の各要素: {key, category, name, summary, unit, quantity,
                     estimate_price, order_price, prob, source}
    """
    threshold = _to_float(threshold)
    candidates: list[dict] = []
    seen: set[str] = set()

    for pattern in patterns or []:
        if not isinstance(pattern, dict):
            continue
        pattern_name = _text(pattern.get('name'))
        for item in pattern.get('items') or []:
            cand = _candidate_from_pattern_item(item, pattern_name)
            if cand is None or cand['key'] in seen:
                continue
            seen.add(cand['key'])
            candidates.append(cand)

    for stat in stats_adj or []:
        if not isinstance(stat, dict):
            continue
        prob = _to_float(stat.get('prob_adj', stat.get('prob')))
        if prob < threshold:
            continue
        cand = _candidate_from_stat(stat)
        if not cand['name'] or cand['key'] in seen:
            continue
        seen.add(cand['key'])
        candidates.append(cand)
    return candidates


# --- 4. パターン登録 -----------------------------------------------------

def pattern_id(name: str) -> str:
    """パターン名から決定的に作る短いID（同じ名前なら常に同じ）。"""
    normalized = unicodedata.normalize('NFKC', _text(name))
    return 'pat_' + hashlib.sha1(normalized.encode('utf-8')).hexdigest()[:10]


def new_pattern(name: str, description: str, items: list[dict], at: str = '') -> dict:
    """パターン（よく一緒に追加する項目の組）を作る。name が空なら ValueError。

    items の各行は不足キーを補い、数値を int/float に正規化する。名称の無い行は捨てる。
    """
    name = _text(name)
    if not name:
        raise ValueError('パターン名を入力してください')
    normalized_items = []
    for raw in items or []:
        item = _normalize_item(raw)
        if item is not None:
            normalized_items.append(item)
    return {
        'id': pattern_id(name),
        'name': name,
        'description': _text(description),
        'items': normalized_items,
        'at': _text(at),
    }


def pattern_from_candidates(name: str, description: str, candidates: list[dict], at: str = '') -> dict:
    """候補（build_candidates の戻り値）をそのままパターンの items にして登録する。"""
    items = [{k: c.get(k) for k in _ITEM_KEYS} for c in (candidates or []) if isinstance(c, dict)]
    return new_pattern(name, description, items, at)


# --- 5. 見積への反映 -------------------------------------------------------

def _strip_koji(norm: str) -> str:
    return norm.replace('工事', '')


def resolve_category_no(category: Any, categories: list[tuple[int, str]],
                        default: int = MISC_CATEGORY_NO) -> int:
    """種別名から工事種別 No を決める（正規化後の完全一致 → 部分一致 → default）。"""
    target = _normalize_name(category)
    pairs = []
    for entry in categories or []:
        try:
            no, cname = entry[0], entry[1]
        except (TypeError, IndexError, KeyError):
            continue
        pairs.append((_to_int(no), _normalize_name(cname)))
    if not target:
        return default
    for no, cname in pairs:
        if cname and cname == target:
            return no
    for no, cname in pairs:
        if cname and (target in cname or cname in target):
            return no
    short = _strip_koji(target)
    if short:
        for no, cname in pairs:
            cshort = _strip_koji(cname)
            if cshort and (short in cshort or cshort in short):
                return no
    return default


def to_new_items(candidates: list[dict], categories: list[tuple[int, str]]) -> list[dict]:
    """候補を mods['new_items'] 形式へ変換する。

    categories: [(no, name), ...]（WorkCategory の .no / .name）。
    種別が特定できない候補は 21（雑工事）へ。各行に 'hidden_key' / 'hidden_source' を付ける。
    """
    result = []
    for cand in candidates or []:
        if not isinstance(cand, dict):
            continue
        name = _text(cand.get('name'))
        if not name:
            continue
        estimate_price = _nonneg_int(cand.get('estimate_price'))
        order_raw = cand.get('order_price')
        order_price = (linked_order_price(0, 0, estimate_price) if order_raw is None
                       else _nonneg_int(order_raw))
        result.append({
            'cat_no': resolve_category_no(cand.get('category'), categories),
            'name': name,
            'summary': _text(cand.get('summary')),
            'quantity': _quantity_or_one(cand.get('quantity')),
            'unit': _text(cand.get('unit')) or '式',
            'estimate_price': estimate_price,
            'order_price': order_price,
            'hidden_key': _text(cand.get('key')) or make_key(cand.get('category'), name),
            'hidden_source': _text(cand.get('source')),
        })
    return result
