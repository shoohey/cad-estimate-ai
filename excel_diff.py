"""
実績見積Excel差分比較モジュール (CAD→概算見積もり生成AI / 日本住建株式会社)

AIが生成した概算見積 (estimate_calculator.Estimate) と、実際に顧客へ提出された
見積Excel（「16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx」形式）を
突合し、差分を明らかにする。目的は以下の2つ:
  1. 画面での差分表示（総額・工事種別・明細の3階層）
  2. 見積精度較正用データの蓄積（パース結果はJSONシリアライズ可能な
     プリミティブのみで構成し、クラウドDBへそのまま保存できる）

提出見積xlsxの構造（「明細書」シート A1:H722 で確認済み）:
  - ヘッダ行: 項目№/名称/摘要/見積数量/単位/見積単価/見積金額
  - サマリ部: 項目№=1〜24 の工事種別サマリ行（単位=式）に続き、
    「端数値引き」（項目№=25等）「消費税等」「総合計」行が並ぶ
  - 明細部: 「総合計」行の後に種別ごとのブロックが続く。3階層構造:
      * 種別ヘッダ行  … 項目№=純数値（例 '1'）・金額なし
      * レベル1明細行 … 項目№='N-M'（例 '1-1'）・金額あり ← 種別合計と一致
      * レベル2内訳行 … 項目№='N-M-K'（例 '1-1-1'）… レベル1の内訳。
        サブヘッダ行（'N-M'・金額なし）に続いて現れ、集計すると
        レベル1と二重計上になるため通常は集計しない。ただし内訳を展開した
        出力（レベル1に金額が無い）では、この行から小計を積み上げて復元する
      * 「合計」行・注記行（※〜）はスキップ
  - 22管理費はサマリ行のみで明細ブロックを持たない

読み取り時の耐性（0円問題への対策）:
  - 列位置は見出し行のラベルから検出する（列構成が違うファイルにも追随）
  - 金額欄は数式（=ROUNDDOWN(数量*単価,0)）で、Excelが保存した計算結果の
    キャッシュを読んでいる。Excel以外のツールを経由してキャッシュを失った
    ファイルは金額が全て空になるため、数量×単価から再計算して復元する
  - それでも合計が0円になる場合は取り込まず、原因を添えて ValueError を送出する

依存: openpyxl + 標準ライブラリのみ
"""

from __future__ import annotations

import io
import re
import unicodedata
from difflib import SequenceMatcher
from typing import Any

import openpyxl

# --- 定数 ---------------------------------------------------------------

DETAIL_SHEET_NAME = "明細書"          # 明細を読むシート名
COVER_SHEET_NAME = "表紙"             # 物件名を読むシート名
HEADER_SEARCH_ROWS = 30               # ヘッダ行を探す最大行数
HEADER_SEARCH_COLS = 8                # ヘッダ行を探す最大列数

# 列見出しラベル → 内部キー。見出し行から実際の列位置を検出するために使う
# （内訳を展開した出力など、列構成が標準と異なるファイルへの耐性）
_COLUMN_LABELS = {
    "名称": "name",
    "摘要": "summary", "備考": "summary",
    "見積数量": "quantity", "数量": "quantity",
    "単位": "unit",
    "見積単価": "price", "単価": "price",
    "見積金額": "amount", "金額": "amount",
}
# 見出しが読めなかった場合の既定列（標準の提出見積フォーマット A〜G列）
_DEFAULT_COLUMNS = {"no": 1, "name": 2, "summary": 3,
                    "quantity": 4, "unit": 5, "price": 6, "amount": 7}

# 明細マッチングのしきい値（過剰マッチより未マッチのほうが良い＝保守的に設定）
ITEM_MATCH_RATIO = 0.80               # difflib類似度の下限
_RE_CODE = re.compile(r"(?<!\d)\d{4,6}(?:-\d)?(?!\d)")   # 名称中の数字コード（サッシ呼称等）
_RE_ANNEX = re.compile(r"追加分|変更|差額|追加")              # 実績側だけの追加・変更行
ITEM_CONTAIN_MIN_LEN = 4              # 部分一致を認める正規化後の最小文字数
ITEM_CONTAIN_MIN_COVER = 0.6          # 部分一致で短い側が長い側を覆う最低比率
                                      # （「大工手間」⊂「●33調整分（大工手間）」の
                                      #   ような別項目への過剰マッチを防ぐ）
_FLOAT_EPS = 1e-6                     # 数量×単価の再計算時に許容する二進小数の誤差
CATEGORY_MATCH_RATIO = 0.85           # 工事種別名のfuzzyマッチ下限
                                      # （外部建具/内部建具の誤マッチ0.83を弾く）

# 項目№のパターン
_RE_CAT_NO = re.compile(r"^\d+$")             # 種別ヘッダ (例 '1')
_RE_ITEM_NO = re.compile(r"^\d+-\d+$")        # レベル1明細 (例 '1-1')
_RE_SUBITEM_NO = re.compile(r"^\d+-\d+-\d+$")  # レベル2内訳 (例 '1-1-1')


# --- 内部ヘルパー -------------------------------------------------------

_RE_LEADING_MARK = re.compile(r"^[●★■◆▲※☆□◇△]+")
_RE_TRAILING_PAREN = re.compile(r"[（(]([^（）()]*)[）)]\s*$")
_RE_ANNOTATION = re.compile(r"令和|平成|\d{4}年|\d{1,2}/\d{1,2}|見積no|見積番号|no\.")
_RE_KEEP = re.compile(r"[^0-9a-z぀-ヿ㐀-鿿豈-﫿]")


def _core(text: Any) -> str:
    return _RE_KEEP.sub("", unicodedata.normalize("NFKC", str(text or "")).lower()).replace("・", "")


def _is_annotation(inner: str, groups: tuple) -> bool:
    """名称末尾の括弧の中身が『付記』（親グループ名・日付・見積No）か"""
    core = _core(inner)
    if not core:
        return True
    if _RE_ANNOTATION.search(unicodedata.normalize("NFKC", inner).lower()):
        return True
    return any(_core(g) and core == _core(g) for g in groups)


def _normalize_name(text: Any, *groups: str) -> str:
    """名称マッチング用に文字列を正規化する。

    NFKC正規化（半角カナ→全角・全角英数→半角）→ 小文字化 → 先頭の記号（●★■等）除去 →
    末尾の括弧付記（親グループ名・日付・見積No）除去 → 空白/記号/「・」/「%」を除去し、
    英数字・かな・カタカナ・漢字のみを残す。
    例: 'ﾍﾞﾀ基礎' → 'ベタ基礎', '板金・樋工事' → '板金樋工事',
        '●33調整分（大工手間）' + groups=('大工手間',) → '33調整分'
    """
    if text is None:
        return ""
    t = unicodedata.normalize("NFKC", str(text)).lower().strip()
    t = _RE_LEADING_MARK.sub("", t).strip()
    m = _RE_TRAILING_PAREN.search(t)
    if m and _is_annotation(m.group(1), tuple(groups)):
        t = t[:m.start()].strip()
    t = t.replace("%", "")
    # 「・」(U+30FB)はカタカナ範囲に含まれるので明示的に除去する
    return _RE_KEEP.sub("", t).replace("・", "")


def _to_int(value: Any) -> int:
    """セル値を int に変換する（None・文字列は0）。"""
    if isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return int(round(value))
    return 0


def _to_float(value: Any) -> float:
    """セル値を float に変換する（None・文字列は0.0）。"""
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


def _cell_str(value: Any) -> str:
    """セル値を文字列化して前後空白を除去する。"""
    if value is None:
        return ""
    return str(value).strip()


def _extract_property_name(wb: openpyxl.Workbook) -> str:
    """表紙シートから「工事名」ラベルの右隣セルを物件名として取得する。"""
    if COVER_SHEET_NAME not in wb.sheetnames:
        return ""
    try:
        ws = wb[COVER_SHEET_NAME]
        for row in ws.iter_rows():
            cells = [c.value for c in row]
            for idx, val in enumerate(cells):
                if _cell_str(val) == "工事名":
                    # 同一行の右側で最初に見つかった非空セルを採用
                    for nxt in cells[idx + 1:]:
                        name = _cell_str(nxt)
                        if name:
                            return name
    except Exception:  # noqa: BLE001 - 表紙は装飾が多いため失敗しても致命でない
        pass
    return ""


def _detect_layout(ws) -> dict:
    """明細シートの見出し行と各列の位置を検出する。

    「項目№」ラベルのあるセルを起点に、同じ行の見出し文字列
    （名称/摘要/見積数量/単位/見積単価/見積金額）から列位置を割り出す。
    見出しが読めない列は標準フォーマットの既定列で補う。

    Returns:
        dict: {'header_row': int, 'no': int, 'name': int, ...} の列番号マップ。

    Raises:
        ValueError: 「項目№」見出し行が見つからない場合。
    """
    header_row = 0
    no_col = 0
    for r in range(1, HEADER_SEARCH_ROWS + 1):
        for c in range(1, HEADER_SEARCH_COLS + 1):
            if _normalize_name(ws.cell(row=r, column=c).value).startswith("項目"):
                header_row, no_col = r, c
                break
        if header_row:
            break
    if not header_row:
        raise ValueError(
            f"「{DETAIL_SHEET_NAME}」シートに項目№の見出し行が見つかりません"
            f"（{HEADER_SEARCH_ROWS}行×{HEADER_SEARCH_COLS}列まで探索）")

    layout = dict(_DEFAULT_COLUMNS)
    layout["no"] = no_col
    found: dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        label = _cell_str(ws.cell(row=header_row, column=c).value)
        key = _COLUMN_LABELS.get(label)
        # 同じ意味の見出しが複数あるときは最初の1つを採用する
        if key and key not in found:
            found[key] = c
    layout.update(found)
    layout["header_row"] = header_row
    return layout


def _has_formula_amounts(file_bytes: bytes, amount_col: int) -> bool:
    """金額列に「計算結果が保存されていない数式」が入っているかを判定する。

    openpyxl は data_only=True のとき Excel が保存したキャッシュ値を読む。
    Excel以外のツールが書き出したブックはキャッシュを持たないため、数式セルが
    すべて None として読まれ、結果として金額が全て0になる。その識別に使う。
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        if DETAIL_SHEET_NAME not in wb.sheetnames:
            return False
        ws = wb[DETAIL_SHEET_NAME]
        for row in ws.iter_rows(min_col=amount_col, max_col=amount_col):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    return True
    except Exception:  # noqa: BLE001 - 診断用のため失敗しても致命ではない
        return False
    return False


def _zero_amount_reason(file_bytes: bytes, ws, layout: dict,
                        n_summary: int, n_item_cats: int) -> str:
    """税抜合計が0円になった原因を推定し、利用者向けの説明文を組み立てる。"""
    col_letter = openpyxl.utils.get_column_letter(layout["amount"])
    head = (f"実績見積の金額を1件も読み取れませんでした（税抜合計が0円）。"
            f"見出し行={layout['header_row']}行目 / 金額列={col_letter}列 / "
            f"サマリ行{n_summary}件・明細のある種別{n_item_cats}件。")
    if _has_formula_amounts(file_bytes, layout["amount"]):
        return (head + " 金額欄が計算式のままで計算結果がファイルに保存されていません。"
                "Excelで開いて上書き保存（または値貼り付け）してから、"
                "もう一度取り込んでください。")
    # 金額列に数値が1つも無い場合は列の取り違えを疑う
    for row in ws.iter_rows(min_row=layout["header_row"] + 1,
                            min_col=layout["amount"], max_col=layout["amount"]):
        for cell in row:
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                return (head + " 金額の数値はありますが、項目№の並びが想定と異なるため"
                        "集計できませんでした。お手数ですがこのファイルをご共有ください。")
    return (head + f" 金額列（{col_letter}列）に数値が見つかりません。"
            "列構成が標準の提出見積と異なる可能性があります。"
            "お手数ですがこのファイルをご共有ください。")


def _similarity(a: str, b: str) -> float:
    """正規化済み文字列同士の類似度スコアを返す。

    完全一致=1.0 / 一定条件を満たす部分一致=0.90 / それ以外はdifflib比。
    部分一致は「短い側が一定長以上」かつ「長い側の6割以上を占める」場合のみ
    認める（短い名称が別項目の名称の一部に飲み込まれる過剰マッチを防ぐ）。
    """
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    if (len(shorter) >= ITEM_CONTAIN_MIN_LEN
            and len(shorter) / len(longer) >= ITEM_CONTAIN_MIN_COVER
            and shorter in longer):
        return 0.90
    return SequenceMatcher(None, a, b).ratio()


# --- 公開 API: 実績見積のパース ------------------------------------------

def parse_actual_estimate(file_bytes: bytes) -> dict:
    """提出見積xlsxのバイト列をパースし、種別・明細を構造化して返す。

    Args:
        file_bytes: 提出見積Excelファイル (.xlsx) のバイト列。

    Returns:
        dict: JSONシリアライズ可能なプリミティブのみで構成された辞書。
            {
                'property_name': str,        # 表紙の工事名（取れなければ空文字）
                'total_excl_tax': int,       # 税抜合計（端数値引き込み・提出値）
                'total_incl_tax': int,       # 税込合計（総合計）
                'tax': int,                  # 消費税等
                'rounding_adjustment': int,  # 端数値引き（例 -206）
                'categories': [
                    {'no': int, 'name': str, 'total': int,
                     'items': [{'name': str, 'summary': str, 'quantity': float,
                                'unit': str, 'price': int, 'amount': int}]}
                ],
                'warnings': [str],           # 整合チェックで検出した注意事項
            }

    Raises:
        ValueError: 「明細書」シートが存在しない／見出し行が見つからない場合、
            または金額を1件も読み取れず税抜合計が0円になった場合
            （原因の推定を含むメッセージを添える）。
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if DETAIL_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"シート「{DETAIL_SHEET_NAME}」が見つかりません: {wb.sheetnames}")
    ws = wb[DETAIL_SHEET_NAME]

    # --- ヘッダ行と列位置を検出する ---
    layout = _detect_layout(ws)
    header_row = layout["header_row"]
    col_no, col_name, col_amount = layout["no"], layout["name"], layout["amount"]

    warnings: list[str] = []
    summary: dict[int, dict[str, Any]] = {}   # {種別no: {'name','total'}}
    items_by_cat: dict[int, list[dict]] = {}  # {種別no: [明細dict]}（レベル1・金額あり）
    # 内訳を展開した出力（レベル1に金額が無くレベル2にのみ金額がある形式）への対応。
    # レベル1が金額を持つ通常形式では未使用のまま捨てられる。
    sub_names: dict[str, str] = {}            # {'N-M': サブヘッダ行の名称}
    subitems: dict[str, list[dict]] = {}      # {'N-M': [レベル2明細dict]}
    level1_with_amount: set[str] = set()      # 金額を持つレベル1の項目№
    level1_item_by_no: dict[str, dict] = {}   # {'N-M': レベル1明細dict}（内訳展開で親を外すため）
    cur_parent = ""                           # 直前に出たレベル1の項目№（内訳行の親判定に使う）
    cat_names: dict[int, str] = {}            # {種別no: 明細部ヘッダ行の名称}
    rounding_adjustment = 0
    tax = 0
    total_incl_tax = 0
    in_detail = False  # False=サマリ部 / True=明細部

    recomputed = 0  # 数量×単価から金額を再計算した行数

    def _amount_of(row: int) -> tuple[int, bool]:
        """行の金額と「金額として使えるか」を返す。

        提出見積の金額欄は数式（=ROUNDDOWN(数量*単価,0)）で、Excelが保存した
        計算結果のキャッシュを読んでいる。Excel以外のツールを経由したファイルは
        このキャッシュを持たず全額が空になるため、数量と単価が数値で読めるときは
        そこから金額を再計算する（見出し行・小計行は数量も単価も空なので対象外）。
        """
        nonlocal recomputed
        raw = ws.cell(row=row, column=col_amount).value
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return int(round(raw)), True
        qty = ws.cell(row=row, column=layout["quantity"]).value
        price = ws.cell(row=row, column=layout["price"]).value
        if (isinstance(qty, (int, float)) and not isinstance(qty, bool)
                and isinstance(price, (int, float)) and not isinstance(price, bool)):
            value = qty * price
            # ExcelのROUNDDOWNは0方向への切り捨て。
            # 62.02*28200 が 1748963.9999… になるような二進小数の誤差で
            # 1円ずれるのを防ぐため、微小な許容差を入れてから切り捨てる。
            recomputed += 1
            adjusted = value + _FLOAT_EPS if value >= 0 else value - _FLOAT_EPS
            return int(adjusted), True
        return 0, False

    def _read_item(row: int, name: str, amount: int) -> dict:
        """明細行を辞書化する（列位置は検出済みレイアウトに従う）。"""
        return {
            "name": name,
            "summary": _cell_str(ws.cell(row=row, column=layout["summary"]).value),
            "quantity": _to_float(ws.cell(row=row, column=layout["quantity"]).value),
            "unit": _cell_str(ws.cell(row=row, column=layout["unit"]).value),
            "price": _to_int(ws.cell(row=row, column=layout["price"]).value),
            "amount": amount,
        }

    for r in range(header_row + 1, ws.max_row + 1):
        no_raw = _cell_str(ws.cell(row=r, column=col_no).value)
        name = _cell_str(ws.cell(row=r, column=col_name).value)

        if not no_raw and not name:
            continue  # 空行

        amount, has_amount = _amount_of(r)

        # 項目№なし行: 消費税等・総合計・注記など
        if not no_raw:
            if "消費税" in name:
                tax = amount
            elif name == "総合計":
                total_incl_tax = amount
                in_detail = True  # 総合計以降が明細部
            # 「合計」行・「※〜」注記行はスキップ
            continue

        if _RE_CAT_NO.fullmatch(no_raw):
            # サマリ部（総合計より前）の種別行だけを合計として採用する。
            # 明細部の種別ヘッダをサマリとして拾うと、金額を再計算したときに
            # 正しい合計を0円で上書きしてしまうため。
            if has_amount and not in_detail:
                # 「端数値引き」等の調整行は種別から分離する
                if "値引" in name or "端数" in name:
                    rounding_adjustment += amount
                else:
                    summary[int(no_raw)] = {"name": name, "total": amount}
            else:
                # 明細部の種別ヘッダ行
                in_detail = True
                cat_names.setdefault(int(no_raw), name)
            continue

        if in_detail and _RE_ITEM_NO.fullmatch(no_raw):
            cur_parent = no_raw
            if not has_amount:
                # レベル2内訳のサブヘッダ行（'N-M'・金額なし）。
                # 展開形式ではここが唯一の名称の出どころになるため控えておく
                sub_names.setdefault(no_raw, name)
                continue
            cat_no = int(no_raw.split("-")[0])
            item = _read_item(r, name, amount)
            item["no"] = no_raw
            items_by_cat.setdefault(cat_no, []).append(item)
            level1_with_amount.add(no_raw)
            level1_item_by_no[no_raw] = item
            continue

        # レベル2内訳行（'N-M-K'）。通常形式ではレベル1と二重計上になるため
        # 集計しないが、親のレベル1に金額が無い場合の復元用に保持する
        if _RE_SUBITEM_NO.fullmatch(no_raw):
            if has_amount:
                prefix = "-".join(no_raw.split("-")[:2])
                # 親は直前のレベル1行で決める（「3-3」配下の行が「3-2-23」と誤採番されている物件への対策）
                parent = cur_parent if (cur_parent and cur_parent.split("-")[0] == prefix.split("-")[0]) else prefix
                child = _read_item(r, name, amount)
                child["no"] = no_raw
                subitems.setdefault(parent, []).append(child)
            continue

    # --- 内訳展開: レベル2(N-M-K)の内訳を突合対象の明細として採用する ---
    # 提出見積は「1-1 安全対策費 387,300」のような親（レベル1・小計）の下に
    # 「1-1-1 駐車場代 171,500 …」の内訳を持つ。AI概算は内訳の粒度で出すため、
    # 親だけを残すと内訳行が全て突合できない（8物件で実績の約36%が親行に隠れていた）。
    # 親は候補から外して group 名として内訳に付け、親金額−内訳合計の残差は
    # 「親名（内訳外）」の1行にして種別合計との整合を保つ。
    groups_by_cat: dict[int, list[dict]] = {}
    expanded = 0
    for parent, children in subitems.items():
        cat_no = int(parent.split("-")[0])
        cat_items = items_by_cat.setdefault(cat_no, [])
        if parent in level1_with_amount:
            p = level1_item_by_no[parent]
            if p in cat_items:
                cat_items.remove(p)
            group_name = p["name"]
            children_sum = int(sum(c["amount"] for c in children))
            groups_by_cat.setdefault(cat_no, []).append(
                {"name": group_name, "amount": int(p["amount"]), "children_sum": children_sum, "no": parent})
            for c in children:
                c["group"] = group_name
                cat_items.append(c)
            residual = int(p["amount"]) - children_sum
            if residual:
                cat_items.append({
                    "name": f"{group_name}（内訳外）", "summary": "", "quantity": 1.0, "unit": "式",
                    "price": residual, "amount": residual, "group": group_name, "no": parent,
                })
        else:
            # 展開形式（親に金額なし）: 内訳をそのまま明細にする
            group_name = sub_names.get(parent, "") or (children[0]["name"] if children else parent)
            children_sum = int(sum(c["amount"] for c in children))
            for c in children:
                c["group"] = group_name
                cat_items.append(c)
            groups_by_cat.setdefault(cat_no, []).append(
                {"name": group_name, "amount": children_sum, "children_sum": children_sum, "no": parent})
        expanded += 1
    if expanded:
        warnings.append(
            f"{expanded}件の小計行を内訳（N-M-K行）に展開して比較します"
            "（親の小計と内訳合計に差がある場合は「（内訳外）」の行として残しています）")

    if recomputed:
        warnings.append(
            f"金額欄の計算結果がファイルに保存されていないため、"
            f"数量×単価から{recomputed}件を再計算しました"
            f"（端数処理の都合で数円の差が出る場合があります）")

    # --- 種別リストを構築（サマリ部を正とし、明細を紐付ける）---
    # 種別名は明細部の種別ヘッダ行からも拾えるようにする（サマリ部が読めない
    # ファイルでも「種別1」のような仮名称にならないようにするため）
    categories: list[dict] = []
    for cat_no in sorted(set(summary) | set(items_by_cat)):
        info = summary.get(cat_no, {})
        items = items_by_cat.get(cat_no, [])
        total = _to_int(info.get("total", 0)) if info else sum(i["amount"] for i in items)
        name = info.get("name", "") or cat_names.get(cat_no, "") or f"種別{cat_no}"
        if not info and summary:
            # サマリ部そのものが読めていない場合は種別ごとに警告を出さない
            warnings.append(f"種別{cat_no}「{name}」はサマリ部に存在しません（明細のみ）")
        # 整合チェック: レベル1明細の合計 = サマリ金額
        if items and info:
            item_sum = sum(i["amount"] for i in items)
            if item_sum != total:
                warnings.append(
                    f"種別{cat_no}「{name}」の明細合計{item_sum:,}円が"
                    f"サマリ金額{total:,}円と一致しません"
                )
        categories.append({"no": cat_no, "name": name, "total": total, "items": items,
                           "groups": groups_by_cat.get(cat_no, [])})

    # 金額0円の種別は読み取り漏れの可能性が高い（実際の見積で0円の種別は稀）
    zero_cats = [c["name"] for c in categories if c["total"] == 0]
    if zero_cats:
        warnings.append(
            "金額が0円として読み取られた工事種別があります: "
            + " / ".join(zero_cats[:8])
            + "（実際に0円でない場合は、Excelで開いて上書き保存してからお試しください）")

    # --- 総額の確定と整合チェック ---
    category_sum = sum(c["total"] for c in categories)
    total_excl_tax = category_sum + rounding_adjustment  # 提出値（端数値引き込み）
    if total_incl_tax and tax and total_incl_tax - tax != total_excl_tax:
        warnings.append(
            f"税込{total_incl_tax:,}円-消費税{tax:,}円={total_incl_tax - tax:,}円が"
            f"税抜合計{total_excl_tax:,}円と一致しません"
        )
    if not total_incl_tax:
        total_incl_tax = total_excl_tax + tax

    # --- 0円ガード ---
    # 金額が1円も読めなかったファイルを「0円の実績見積」として取り込むと、
    # 差分画面が丸ごと誤った数字になる。取り込まずに原因を伝えて中断する。
    if total_excl_tax <= 0:
        raise ValueError(_zero_amount_reason(file_bytes, ws, layout,
                                             len(summary), len(items_by_cat)))

    return {
        "property_name": _extract_property_name(wb),
        "total_excl_tax": total_excl_tax,
        "total_incl_tax": total_incl_tax,
        "tax": tax,
        "rounding_adjustment": rounding_adjustment,
        "categories": categories,
        "warnings": warnings,
    }


# --- 内部ヘルパー: マッチング ---------------------------------------------

def _match_categories(ai_cats: list, actual_cats: list) -> list[tuple]:
    """AI種別と実績種別を名称ベースで対応付ける。

    Returns:
        list[tuple]: (ai_category | None, actual_category(dict) | None) のリスト。
            片方にしか無い種別は相手側を None として含める。
    """
    pairs: list[tuple] = []
    used_actual: set[int] = set()

    # 1) 正規化名の完全一致
    actual_by_norm = {}
    for idx, ac in enumerate(actual_cats):
        actual_by_norm.setdefault(_normalize_name(ac["name"]), idx)
    unmatched_ai = []
    for cat in ai_cats:
        idx = actual_by_norm.get(_normalize_name(cat.name))
        if idx is not None and idx not in used_actual:
            pairs.append((cat, actual_cats[idx]))
            used_actual.add(idx)
        else:
            unmatched_ai.append(cat)

    # 2) 残りをfuzzyマッチ（高類似のみ・スコア降順に貪欲確定）
    candidates = []
    for cat in unmatched_ai:
        for idx, ac in enumerate(actual_cats):
            if idx in used_actual:
                continue
            score = _similarity(_normalize_name(cat.name), _normalize_name(ac["name"]))
            if score >= CATEGORY_MATCH_RATIO:
                candidates.append((score, id(cat), cat, idx))
    matched_ai_ids: set[int] = set()
    for score, cat_id, cat, idx in sorted(candidates, key=lambda t: -t[0]):
        if cat_id in matched_ai_ids or idx in used_actual:
            continue
        pairs.append((cat, actual_cats[idx]))
        matched_ai_ids.add(cat_id)
        used_actual.add(idx)

    # 3) どちらにも対応が付かなかった種別
    for cat in unmatched_ai:
        if id(cat) not in matched_ai_ids:
            pairs.append((cat, None))
    for idx, ac in enumerate(actual_cats):
        if idx not in used_actual:
            pairs.append((None, ac))

    return pairs


def _ai_items_to_dicts(ai_items: list) -> list[dict]:
    """AI見積の明細オブジェクトを比較用の辞書に揃える。"""
    return [{
        "name": it.name,
        "summary": _cell_str(getattr(it, "summary", "")),
        "quantity": _to_float(getattr(it, "quantity", 0)),
        "unit": _cell_str(getattr(it, "unit", "")),
        "amount": int(it.estimate_amount),
        "group": _cell_str(getattr(it, "category", "")),
    } for it in ai_items]


def _match_keys(item: dict, *extra_groups: str) -> tuple[str, str]:
    """明細の照合キー（名称のみ／名称＋摘要）を正規化して返す。

    名称末尾の括弧が自分または相手側の group 名（親の小計名）なら付記として外す
    （例: 「33調整分（大工手間）」⇔ 3-3 大工手間 配下の「33調整分」）。
    """
    name = _cell_str(item.get("name"))
    summary = _cell_str(item.get("summary"))
    groups = tuple(g for g in (_cell_str(item.get("group")), *extra_groups) if g)
    return _normalize_name(name, *groups), _normalize_name(name + summary, *groups)


def item_alias_keys(item: dict) -> set[str]:
    """同一商品辞書で使う照合キー（名称のみ／名称＋摘要の正規化文字列）"""
    return {k for k in _match_keys(item) if k}


def is_alias_pair(ai: dict, ac: dict, aliases: set | None) -> bool:
    """利用者が「同じ商品」と登録した組か。aliases は frozenset({key_a, key_b}) の集合。"""
    if not aliases:
        return False
    for a in item_alias_keys(ai):
        for b in item_alias_keys(ac):
            if frozenset((a, b)) in aliases:
                return True
    return False


def _item_similarity(ai: dict, ac: dict) -> float:
    """名称・摘要の組み合わせで最も高い類似度を返す。

    実績見積は「名称=鋼板屋根 / 摘要=荷揚げ費」と列を分けて書くのに対し、
    AI側は「鋼板屋根　荷揚げ費」と名称に含めることがある。名称だけで比べると
    同じ項目が別物として扱われるため、摘要を連結した表記どうしも突き合わせる。
    """
    ai_name, ai_full = _match_keys(ai, _cell_str(ac.get("group")))
    ac_name, ac_full = _match_keys(ac, _cell_str(ai.get("group")))
    return max(
        _similarity(ai_name, ac_name),
        _similarity(ai_full, ac_full),
        _similarity(ai_name, ac_full),
        _similarity(ai_full, ac_name),
    )


def _rollup_items(items: list[dict]) -> list[dict]:
    """同一名称・同一単位の明細を1行に集約する。

    CADの建具リストは同じ製品でも1本ずつ別行で出てくるため、そのまま並べると
    実績見積（同じ製品を1行にまとめて数量で表す）と1対1に対応せず、
    同じ品名が何行も並んで重複計上のように見えてしまう。
    比較の粒度を実績側に合わせるため、数量・金額を合算して1行にまとめる。
    金額の合計は変わらないので、総額・種別合計には影響しない。
    """
    merged: dict[tuple, dict] = {}
    order: list[tuple] = []
    for it in items:
        # 摘要までキーに含める（同じ名称でも摘要が違えば別の項目として扱う）
        key = (_match_keys(it)[1], _cell_str(it.get("unit", "")), _cell_str(it.get("group", "")))
        cur = merged.get(key)
        if cur is None:
            merged[key] = {
                "name": it["name"],
                "summary": _cell_str(it.get("summary", "")),
                "quantity": _to_float(it.get("quantity", 0)),
                "unit": _cell_str(it.get("unit", "")),
                "amount": int(it["amount"]),
                "rows": 1,
                "group": _cell_str(it.get("group", "")),
                "pseudo": bool(it.get("pseudo", False)),
            }
            order.append(key)
        else:
            cur["quantity"] += _to_float(it.get("quantity", 0))
            cur["amount"] += int(it["amount"])
            cur["rows"] += 1
    return [merged[k] for k in order]


def _match_items(ai_items: list[dict], actual_items: list[dict],
                 aliases: set | None = None) -> dict:
    """種別内の明細同士を名称類似度で対応付ける。

    引数はいずれも _rollup_items() で集約済みの辞書リスト。
    完全一致 → 部分一致（正規化後4文字以上）→ difflib類似度0.80以上の順で
    スコアリングし、スコア降順の貪欲法で1対1に確定する。
    しきい値未満は無理にマッチさせない（保守的）。
    """
    scored = []
    alias_hits: set[tuple[int, int]] = set()
    for i, ai in enumerate(ai_items):
        for j, ac in enumerate(actual_items):
            # 利用者が「同じ商品」と登録した組は、名称が似ていなくても最優先で突合する
            if is_alias_pair(ai, ac, aliases):
                score = 1.0
                alias_hits.add((i, j))
            else:
                score = _item_similarity(ai, ac)
                # 名称中の数字コード（サッシ呼称 16005 等）が両側にあり一致しない組は、
                # 文字の並びが似ていても別品目の可能性が高いので順位を下げる
                codes_ai = set(_RE_CODE.findall(unicodedata.normalize("NFKC", _cell_str(ai.get("name")))))
                codes_ac = set(_RE_CODE.findall(unicodedata.normalize("NFKC", _cell_str(ac.get("name")))))
                if codes_ai and codes_ac and codes_ai.isdisjoint(codes_ac) and score < 1.0:
                    score = min(score, ITEM_MATCH_RATIO)
            if score < ITEM_MATCH_RATIO:
                continue
            # 順位付け（マッチ可否の判定には使わない）:
            #   名称＋摘要の完全一致 > 登録済みの組 > 名称のみ一致 > 部分一致 > 類似
            #   同じ group（親の小計名）・単位・数量が揃う組を優先し、
            #   実績側だけが「追加分／変更／差額」の行は後回し、最後は金額の近さ
            rank = score
            _, ai_full_key = _match_keys(ai, _cell_str(ac.get("group")))
            _, ac_full_key = _match_keys(ac, _cell_str(ai.get("group")))
            if ai_full_key and ai_full_key == ac_full_key:
                rank += 0.02
            elif (i, j) in alias_hits:
                rank += 0.01
            if _cell_str(ai.get("group")) and _core(ai.get("group")) == _core(ac.get("group")):
                rank += 0.003
            if _cell_str(ai.get("unit")) == _cell_str(ac.get("unit")):
                rank += 0.002
            if abs(_to_float(ai.get("quantity")) - _to_float(ac.get("quantity"))) < 1e-6:
                rank += 0.001
            ac_text = _cell_str(ac.get("name")) + _cell_str(ac.get("summary"))
            ai_text = _cell_str(ai.get("name")) + _cell_str(ai.get("summary"))
            if _RE_ANNEX.search(ac_text) and not _RE_ANNEX.search(ai_text):
                rank -= 0.005
            ai_amt, ac_amt = abs(int(ai.get("amount", 0))), abs(int(ac.get("amount", 0)))
            rank += 0.0005 * (1 - min(1.0, abs(ai_amt - ac_amt) / max(ai_amt, ac_amt, 1)))
            scored.append((rank, i, j))

    matched = []
    used_ai: set[int] = set()
    used_actual: set[int] = set()
    for _rank, i, j in sorted(scored, key=lambda t: -t[0]):
        if i in used_ai or j in used_actual:
            continue
        ai, ac = ai_items[i], actual_items[j]
        matched.append({
            "ai_name": ai["name"],
            "actual_name": ac["name"],
            "ai_summary": ai.get("summary", ""),
            "actual_summary": ac.get("summary", ""),
            "ai_amount": int(ai["amount"]),
            "actual_amount": int(ac["amount"]),
            "diff": int(ai["amount"]) - int(ac["amount"]),
            "ai_quantity": ai["quantity"],
            "actual_quantity": ac["quantity"],
            "unit": ai.get("unit") or ac.get("unit", ""),
            "ai_rows": ai.get("rows", 1),
            "via_alias": (i, j) in alias_hits,
            "via_pseudo": bool(ac.get("pseudo", False)),
            "group": ac.get("group") or ai.get("group", ""),
        })
        used_ai.add(i)
        used_actual.add(j)

    def _side(items: list[dict], used: set[int]) -> list[dict]:
        return [{"name": it["name"], "amount": int(it["amount"]),
                 "summary": it.get("summary", ""),
                 "quantity": it["quantity"], "unit": it.get("unit", ""),
                 "rows": it.get("rows", 1), "group": it.get("group", "")}
                for k, it in enumerate(items) if k not in used]

    return {"matched": matched,
            "ai_only": _side(ai_items, used_ai),
            "actual_only": _side(actual_items, used_actual)}


# --- 公開 API: 差分計算 ---------------------------------------------------

def diff_estimates(estimate, actual: dict, aliases: set | None = None) -> dict:
    """AI概算見積と実績見積パース結果の差分を計算する。

    Args:
        estimate: estimate_calculator.Estimate オブジェクト（AI概算見積）。
        actual: parse_actual_estimate() の戻り値（実績見積）。
        aliases: 利用者が登録した「同じ商品」の組（frozenset({key_a, key_b}) の集合）。
            名称が似ていなくてもこの組は突合する。

    Returns:
        dict: JSONシリアライズ可能な差分構造。
            {
                'total_ai': int, 'total_actual': int,
                'diff': int,                  # total_ai - total_actual
                'rate': float,                # diff / total_actual
                'category_diffs': [           # 名称マッピング済み・片側のみも含む
                    {'no', 'name', 'ai_total', 'actual_total', 'diff', 'rate'}
                ],
                'item_diffs': {
                    種別名: {'matched': [{'ai_name','actual_name',
                                          'ai_amount','actual_amount','diff'}],
                             'ai_only': [{'name','amount'}],
                             'actual_only': [{'name','amount'}]}
                },
            }
    """
    total_ai = int(estimate.total_estimate_excl_tax)
    total_actual = int(actual["total_excl_tax"])
    diff = total_ai - total_actual
    rate = diff / total_actual if total_actual else 0.0

    category_diffs: list[dict] = []
    item_diffs: dict[str, dict] = {}

    for ai_cat, actual_cat in _match_categories(list(estimate.categories),
                                                list(actual["categories"])):
        ai_total = int(ai_cat.estimate_total) if ai_cat is not None else 0
        actual_total = int(actual_cat["total"]) if actual_cat is not None else 0
        cat_no = ai_cat.no if ai_cat is not None else actual_cat["no"]
        cat_name = ai_cat.name if ai_cat is not None else actual_cat["name"]
        cat_diff = ai_total - actual_total
        category_diffs.append({
            "no": int(cat_no),
            "name": cat_name,
            "ai_total": ai_total,
            "actual_total": actual_total,
            "diff": cat_diff,
            "rate": cat_diff / actual_total if actual_total else 0.0,
        })
        # 明細差分（どちらか一方しか無い種別も ai_only / actual_only で表現）
        # 双方とも同一名称の明細を1行に集約してから突合する（粒度合わせ）
        ai_items = _rollup_items(
            _ai_items_to_dicts(ai_cat.items)) if ai_cat is not None else []
        actual_raw = list(actual_cat["items"]) if actual_cat is not None else []
        if actual_cat is not None and not actual_raw and int(actual_cat["total"]) > 0:
            # 管理費のようにサマリ部にしか無い種別は、種別合計を1行の擬似明細にして突合する
            actual_raw = [{"name": actual_cat["name"], "summary": "", "quantity": 1.0, "unit": "式",
                           "price": int(actual_cat["total"]), "amount": int(actual_cat["total"]),
                           "pseudo": True}]
        actual_items = _rollup_items(actual_raw)
        if ai_items or actual_items:
            item_diffs[cat_name] = _match_items(ai_items, actual_items, aliases)

    category_diffs.sort(key=lambda d: d["no"])

    return {
        "total_ai": total_ai,
        "total_actual": total_actual,
        "diff": diff,
        "rate": rate,
        "category_diffs": category_diffs,
        "item_diffs": item_diffs,
    }


# --- 公開 API: 日本語サマリ -----------------------------------------------

def diff_summary_text(diff: dict) -> str:
    """差分結果からフィードバック送信用の日本語サマリテキストを生成する。

    総額差・率、差の大きい工事種別トップ5、代表的な明細差（金額差の大きい
    マッチ明細＋片側にしか無い主要明細）をプレーンテキストで整形する。
    """
    lines: list[str] = []
    lines.append("【実績見積との差分サマリ】")
    lines.append("■ 総額（税抜）")
    lines.append(f"  AI概算: {diff['total_ai']:,}円 / 実績: {diff['total_actual']:,}円")
    lines.append(f"  差額: {diff['diff']:+,}円（{diff['rate']:+.1%}）")

    # 差の大きい工事種別トップ5
    top_cats = sorted(diff["category_diffs"], key=lambda d: -abs(d["diff"]))[:5]
    lines.append("■ 差の大きい工事種別 TOP5")
    for rank, cat in enumerate(top_cats, start=1):
        rate_str = f" / {cat['rate']:+.1%}" if cat["actual_total"] else ""
        lines.append(
            f"  {rank}. {cat['name']}: AI {cat['ai_total']:,}円 / "
            f"実績 {cat['actual_total']:,}円（差 {cat['diff']:+,}円{rate_str}）"
        )

    # 代表的な明細差: マッチ済みで金額差の大きいものトップ5
    matched_all = []
    ai_only_all = []
    actual_only_all = []
    for cat_name, detail in diff["item_diffs"].items():
        for m in detail["matched"]:
            if m["diff"] != 0:
                matched_all.append((cat_name, m))
        for it in detail["ai_only"]:
            ai_only_all.append((cat_name, it))
        for it in detail["actual_only"]:
            actual_only_all.append((cat_name, it))

    lines.append("■ 代表的な明細差")
    for cat_name, m in sorted(matched_all, key=lambda t: -abs(t[1]["diff"]))[:5]:
        lines.append(
            f"  [{cat_name}] {m['ai_name']}: AI {m['ai_amount']:,}円 / "
            f"実績 {m['actual_amount']:,}円（差 {m['diff']:+,}円）"
        )
    for cat_name, it in sorted(ai_only_all, key=lambda t: -abs(t[1]["amount"]))[:3]:
        lines.append(f"  （AIのみ）[{cat_name}] {it['name']}: {it['amount']:,}円")
    for cat_name, it in sorted(actual_only_all, key=lambda t: -abs(t[1]["amount"]))[:3]:
        lines.append(f"  （実績のみ）[{cat_name}] {it['name']}: {it['amount']:,}円")
    if not matched_all and not ai_only_all and not actual_only_all:
        lines.append("  （明細レベルの差分はありません）")

    return "\n".join(lines)


# --- CLI エントリポイント -----------------------------------------------

if __name__ == "__main__":
    # 古居様邸の実データで動作確認する（開発用）
    import json
    from pathlib import Path

    base = Path(__file__).resolve().parent
    xlsx_path = base / "16-古居徹朗様・古居百恵様邸新築工事【提出見積】.xlsx.xlsx"

    print("=== excel_diff.py 動作確認 ===")
    actual = parse_actual_estimate(xlsx_path.read_bytes())
    print(f"物件名: {actual['property_name']}")
    print(f"税抜合計: {actual['total_excl_tax']:,}円 / 税込合計: {actual['total_incl_tax']:,}円")
    print(f"種別数: {len(actual['categories'])} / "
          f"明細数: {sum(len(c['items']) for c in actual['categories'])}")
    for w in actual["warnings"]:
        print(f"  警告: {w}")

    from cad_parser import load_cad_file
    from estimate_calculator import calculate_estimate, load_unit_prices

    cad = load_cad_file(str(base / "16-古居徹朗様・古居百恵様邸新築工事【数量データ】.TXT"))
    est = calculate_estimate(cad, load_unit_prices(str(base / "tankamaster_updated.csv")))
    result = diff_estimates(est, actual)
    print(f"\n差分率: {result['rate']:+.2%}")
    print()
    print(diff_summary_text(result))
    # JSONシリアライズ可能であることの確認
    json.dumps(actual, ensure_ascii=False)
    json.dumps(result, ensure_ascii=False)
    print("\nJSONシリアライズ: OK")
