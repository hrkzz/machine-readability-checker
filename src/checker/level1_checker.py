from pathlib import Path
import re
from typing import Tuple
from openpyxl.workbook.workbook import Workbook

from src.processor.context import TableContext
from src.checker.utils import (
    has_any_drawing_xlsx,
    detect_platform_characters,
    get_excel_column_letter,
    MAX_EXAMPLES
)
from src.llm.llm_client import call_llm

# L1-01: ファイル形式はExcelかCSV
def check_valid_file_format(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    ext = Path(filepath).suffix.lower()
    if ext not in {".csv", ".xlsx"}:
        return False, f"サポート外のファイル形式です: {ext}"
    return True, "ファイル形式はCSVまたはExcelです"

# L1-02: オブジェクトや画像を使わない
def check_no_images_or_objects(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    path = Path(filepath)
    if path.suffix.lower() != ".xlsx":
        return True, "csvファイルのためオブジェクトチェック不要"
    if has_any_drawing_xlsx(path):
        return False, "図形・テキストボックスが検出されました"
    return True, "図形・テキストボックスは見つかりませんでした"

# L1-03: 1シートに1つの表を入れる
def check_one_table_per_sheet(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    ws = workbook[ctx.sheet_name]
    start = min(ctx.row_indices["column_rows"])
    end = ctx.row_indices["data_end"]

    flags = [
        any(cell not in (None, "") for cell in row)
        for row in ws.iter_rows(min_row=start + 1, max_row=end + 1, values_only=True)
    ]

    in_block = False
    blocks = 0
    for has_data in flags:
        if has_data and not in_block:
            blocks += 1
            in_block = True
        elif not has_data:
            in_block = False

    if blocks > 1:
        return False, f"複数テーブルの疑いがあります（検出ブロック数: {blocks}）"
    return True, "1つのテーブルのみです"

# L1-04: 行や列を非表示にしない
def check_no_hidden_rows_or_columns(
    ctx: TableContext,
    workbook: Workbook,
    filepath: str
) -> Tuple[bool, str]:
    """
    対象シートで非表示の行／列が存在しないことを確認。
    注釈などは含まない。
    """
    ws = workbook[ctx.sheet_name]
    hidden_rows = [d.index for d in ws.row_dimensions.values() if d.hidden]
    hidden_cols = [d.index for d in ws.column_dimensions.values() if d.hidden]

    if hidden_rows or hidden_cols:
        row_str = hidden_rows if hidden_rows else "該当なし"
        col_str = hidden_cols if hidden_cols else "該当なし"
        return False, f"非表示行／列があります（行: {row_str}, 列: {col_str}）"

    return True, "非表示行／列はありません"

# L1-05: 表の外側にメモや備考を記載しない
def check_no_notes_outside_table(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    if not ctx.upper_annotations.empty or not ctx.lower_annotations.empty:
        top = ctx.row_indices["column_rows"][0]
        bot = ctx.row_indices["data_end"]
        return False, f"注釈行が検出されました（{top}行目より前、または{bot+2}行目以降）"
    return True, "表外の注釈や備考はありません"

# L1-06: セルを結合しない（データ + ヘッダーの範囲のみ）
def check_no_merged_cells(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    ws = workbook[ctx.sheet_name]
    start = min(ctx.row_indices["column_rows"]) + 1  # 1-indexed
    end = ctx.row_indices["data_end"] + 1

    relevant_merges = []
    for rng in ws.merged_cells.ranges:
        # セル範囲（例: A3:D5）を openpyxl の RangeBound で判定
        if rng.min_row >= start and rng.max_row <= end:
            relevant_merges.append(str(rng))

    if relevant_merges:
        return False, f"結合セルが検出されました: {relevant_merges}"
    return True, "結合セルはありません"

# L1-07: 書式でデータの違いを表現しない（網掛けなど）
def check_no_format_based_semantics(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    """
    セルの視覚的書式（塗りつぶし、文字色、太字、イタリック、下線、フォントサイズ）で
    意味の違いを表現していないかを検出する。
    """
    ws = workbook[ctx.sheet_name]
    start = min(ctx.row_indices["column_rows"]) + 1
    end = ctx.row_indices["data_end"] + 1

    flagged = []

    for row in ws.iter_rows(min_row=start, max_row=end):
        for cell in row:
            coord = cell.coordinate

            # 塗りつぶしの色
            fill = cell.fill
            if fill and fill.fgColor:
                fg = fill.fgColor
                if hasattr(fg, "rgb") and isinstance(fg.rgb, str):
                    rgb = fg.rgb.upper()
                    if rgb not in ("00000000", "FFFFFFFF", "FF000000"):
                        flagged.append(f"{coord}（塗りつぶし）")

            # 文字色・太字・イタリック・下線・フォントサイズ
            font = cell.font
            if font:
                # 文字色
                if font.color:
                    color = font.color
                    if hasattr(color, "rgb") and isinstance(color.rgb, str):
                        rgb = color.rgb.upper()
                        if rgb not in ("00000000", "FF000000"):
                            flagged.append(f"{coord}（文字色）")

                # 太字
                if font.bold:
                    flagged.append(f"{coord}（太字）")

                # イタリック
                if font.italic:
                    flagged.append(f"{coord}（イタリック）")

                # 下線
                if font.underline:
                    flagged.append(f"{coord}（下線）")

                # フォントサイズ（通常より極端に大 or 小）
                if font.sz and (font.sz < 9 or font.sz > 13):
                    flagged.append(f"{coord}（フォントサイズ {font.sz}）")

    if flagged:
        return False, f"視覚的装飾による意味付けが検出されました（例: {flagged[:MAX_EXAMPLES]}）"
    return True, "書式ベースの意味づけは検出されませんでした"

# L1-08: スペースや改行で体裁を整えない
def check_no_whitespace_formatting(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    """
    LLMを用いて、体裁調整目的の空白（特に全角スペース）が含まれていないかを判定。
    対象：カラム行 + データ行
    """
    from src.llm.llm_client import call_llm

    ws = workbook[ctx.sheet_name]
    start = min(ctx.row_indices["column_rows"]) + 1
    end = ctx.row_indices["data_end"] + 1

    sample_cells = []

    for r_idx, row in enumerate(ws.iter_rows(min_row=start, max_row=end, values_only=True), start=start):
        for c_idx, val in enumerate(row, start=1):
            if isinstance(val, str) and "　" in val:
                col_letter = get_excel_column_letter(c_idx)
                cell_ref = f"{col_letter}{r_idx}"
                sample_cells.append(f"{cell_ref}: '{val.strip()}'")
                if len(sample_cells) >= 10:
                    break
        if len(sample_cells) >= 10:
            break

    if not sample_cells:
        return True, "体裁調整目的の空白は見つかりませんでした"

    prompt = f"""
            以下はExcelのセル値の一部です。これらの中に、見た目を整える目的（位置揃え・スペース調整など）で
            **空白（特に全角スペース）が使われているものがあるか**を判定してください。

            データ:
            {chr(10).join(sample_cells)}

            判断結果を次のいずれか一語で返してください：
            - 調整目的あり
            - 調整目的なし
            """

    result = call_llm(prompt)
    if "調整目的あり" in result:
        return False, f"体裁調整目的の空白が含まれている可能性があります（例: {sample_cells[:MAX_EXAMPLES]}）"
    return True, "体裁調整目的の空白は見つかりませんでした"


# L1-09: 1セルに1データしか入れない
def check_single_data_per_cell(
    ctx: TableContext,
    workbook: Workbook,
    filepath: str
) -> Tuple[bool, str]:
    """
    各セルに複数データ（改行・カンマ・スラッシュなど）が混在していないか確認。
    """
    pattern = re.compile(r"[\n,;/]")
    problems = []

    for row_idx, row in ctx.data.iterrows():
        for col_idx, val in enumerate(row):
            if isinstance(val, str) and pattern.search(val):
                excel_row = row_idx + 1 + ctx.row_indices["data_start"]  # 実Excel上の行
                excel_col_letter = get_excel_column_letter(col_idx + 1)        # A, B, C...
                coord = f"{excel_col_letter}{excel_row}"
                problems.append(f"{coord}: {repr(val)}")

    if problems:
        return False, f"複数データセルが検出されました（例: {problems[:MAX_EXAMPLES]}）"
    return True, "各セルに1データのみです"


# L1-10: 機種依存文字を使わない
def check_no_platform_dependent_characters(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    ws = workbook[ctx.sheet_name]
    start = min(ctx.row_indices["column_rows"]) + 1
    end = ctx.row_indices["data_end"] + 1

    issues = []
    for r, row in enumerate(ws.iter_rows(min_row=start, max_row=end, values_only=True), start=start):
        for c, val in enumerate(row, start=1):
            if isinstance(val, str) and detect_platform_characters(val):
                coord = f"{get_excel_column_letter(c)}{r}"
                issues.append(f"{coord}: '{val}'")

    if issues:
        return False, f"機種依存文字が含まれています（例: {issues[:MAX_EXAMPLES]}）"
    return True, "機種依存文字は含まれていません"
