from pathlib import Path
import re
from typing import Tuple, cast
from openpyxl.workbook.workbook import Workbook
import pandas as pd
import xlrd
from loguru import logger

from src.processor.context import TableContext
from src.checker.utils import (
    has_any_drawing_xlsx,
    detect_platform_characters,
    get_excel_column_letter,
    MAX_EXAMPLES,
)
from src.llm.llm_client import call_llm

logger.add("logs/checker1.log", rotation="10 MB", retention="30 days", level="DEBUG")


def check_xls_merged_cells(file_path: Path) -> list:
    """xlsファイルの結合セルをチェック"""
    try:
        workbook = xlrd.open_workbook(str(file_path))
        merged_cells = []
        
        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)
            if hasattr(sheet, 'merged_cells'):
                for row_start, row_end, col_start, col_end in sheet.merged_cells:
                    merged_cells.append({
                        'sheet': sheet_name,
                        'range': f"{row_start}:{row_end-1}, {col_start}:{col_end-1}"
                    })
        
        return merged_cells
    except Exception as e:
        logger.error(f"結合セルチェックでエラー: {e}")
        return []


def check_xls_cell_formats(file_path: Path, sheet_name: str, data_start: int, data_end: int) -> list:
    """xlsファイルのセル書式をチェック（修正版）"""
    try:
        logger.debug(f"check_xls_cell_formats: 開始 - {file_path}, sheet: {sheet_name}")
        workbook = xlrd.open_workbook(str(file_path), formatting_info=True)
        sheet = workbook.sheet_by_name(sheet_name)
        flagged = []

        for row_idx in range(data_start, min(data_end + 1, sheet.nrows)):
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                xf_index = cell.xf_index

                if xf_index >= len(workbook.xf_list):
                    continue  # 異常なインデックスはスキップ

                xf = workbook.xf_list[xf_index]
                font_index = xf.font_index

                if font_index >= len(workbook.font_list):
                    continue

                font = workbook.font_list[font_index]
                coord = f"{get_excel_column_letter(col_idx + 1)}{row_idx + 1}"

                # 太字
                if font.bold:
                    flagged.append(f"{coord}（太字）")
                # イタリック
                if font.italic:
                    flagged.append(f"{coord}（イタリック）")
                # 下線
                if font.underline_type != 0:
                    flagged.append(f"{coord}（下線）")
                # 文字色
                if font.colour_index not in (0, 1, 7, 8):  # 自動・黒・白以外
                    flagged.append(f"{coord}（文字色）")
                # 背景色
                bg_index = xf.background.pattern_colour_index
                if bg_index not in (64, 0):  # 標準色以外
                    flagged.append(f"{coord}（背景色）")

        return flagged

    except Exception as e:
        logger.exception(f"書式チェックでエラー: {e}")
        return []

def detect_multiple_tables_dataframe(df: pd.DataFrame, sheet_name: str = "") -> tuple:
    """
    DataFrameベースで複数テーブルを検出する
    
    Args:
        df: 対象のDataFrame
        sheet_name: シート名（ログ用）
    
    Returns:
        tuple: (has_multiple_tables: bool, details: str)
    """
    try:
        if df.empty or len(df) < 3:
            return False, "データが少ないため複数テーブルの検出をスキップ"
        
        # 完全に空の行を検索
        empty_rows = []
        for idx, row in df.iterrows():
            if row.isna().all() or (row.astype(str).str.strip() == "").all():
                empty_rows.append(idx)
        
        # 連続する空行を検索（テーブル区切りの可能性）
        if len(empty_rows) > 0:
            consecutive_groups = []
            current_group = [empty_rows[0]]
            
            for i in range(1, len(empty_rows)):
                if empty_rows[i] == empty_rows[i-1] + 1:
                    current_group.append(empty_rows[i])
                else:
                    if len(current_group) >= 2:  # 2行以上の連続空行
                        consecutive_groups.append(current_group)
                    current_group = [empty_rows[i]]
            
            if len(current_group) >= 2:
                consecutive_groups.append(current_group)
            
            if consecutive_groups:
                return True, f"複数の連続空行グループが見つかりました: {len(consecutive_groups)}箇所"
        
        # ヘッダー様の行の検出
        header_like_rows = []
        for idx, row in df.iterrows():
            non_na_values = row.dropna().astype(str).str.strip()
            if len(non_na_values) > 0:
                # 数値以外が多い行をヘッダー候補とする
                numeric_count = sum(1 for val in non_na_values if val.replace('.', '').replace('-', '').isdigit())
                if numeric_count / len(non_na_values) < 0.5:
                    header_like_rows.append(idx)
        
        # 複数のヘッダー様行が離れて存在する場合
        if len(header_like_rows) >= 2:
            gaps = [header_like_rows[i+1] - header_like_rows[i] for i in range(len(header_like_rows)-1)]
            if any(gap > 3 for gap in gaps):  # 3行以上離れたヘッダーがある
                return True, f"離れた位置に複数のヘッダー様行が検出されました: {header_like_rows}"
        
        return False, "単一テーブルと判定"
        
    except Exception as e:
        logger.error(f"複数テーブル検出でエラー: {e}")
        return False, f"検出処理でエラーが発生: {str(e)}"


def check_valid_file_format(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    ext = Path(filepath).suffix.lower()
    if ext not in {".csv", ".xlsx", ".xls"}:
        return False, f"サポート外のファイル形式です: {ext}"
    elif ext == ".xls":
        return True, "旧Excel（.xls）形式のため、一部の自動チェック（書式・図形など）が制限されます。必要に応じて目視での確認を行ってください"
    return True, "ファイル形式はCSVまたはExcel（.xlsx）です"


def check_no_images_or_objects(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    path = Path(filepath)
    ext = path.suffix.lower()
    if ext == ".csv":
        return True, "csvファイルのためオブジェクトチェック不要"
    elif ext == ".xls":
        # .xlsファイルでは図形・オブジェクトの詳細検出は困難だが、
        # 一般的に統計表では図形は使用されないため、警告として扱う
        return False, "xlsファイルでは図形や画像の自動判定ができません。必要に応じて目視でご確認ください"
    elif ext == ".xlsx":
        if has_any_drawing_xlsx(path):
            return False, "図形・テキストボックスが検出されました"
        return True, "図形・テキストボックスは見つかりませんでした"
    else:
        return True, "サポート外形式のためオブジェクトチェック不要"


def check_one_table_per_sheet(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    # .xlsファイルの場合はDataFrameベースでチェック
    if workbook is None:
        # 元のDataFrameを使用して複数テーブル検出
        is_multiple, details = detect_multiple_tables_dataframe(ctx.data, ctx.sheet_name)
        
        if is_multiple:
            return False, f"複数テーブルの疑いがあります: {details}"
        return True, "1つのテーブルのみです"
    
    ws = workbook[ctx.sheet_name]
    column_rows = ctx.row_indices.get("column_rows")
    data_end = ctx.row_indices.get("data_end")

    if column_rows is None or data_end is None:
        return False, "シート範囲情報が不足しているためチェックできません"

    start = min(column_rows) if isinstance(column_rows, list) else cast(int, column_rows)
    end = cast(int, data_end)

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


def check_xls_hidden_rows_columns(file_path: Path) -> tuple:
    """xlsファイルの非表示行・列をチェック（修正版）"""
    try:
        logger.debug(f"check_xls_hidden_rows_columns: 開始 - {file_path}")
        workbook = xlrd.open_workbook(str(file_path), formatting_info=True)
        hidden_rows = []
        hidden_cols = []

        for sheet_name in workbook.sheet_names():
            logger.debug(f"シート処理中: {sheet_name}")
            sheet = workbook.sheet_by_name(sheet_name)

            # 行の高さが0 → 非表示
            for row_idx in range(sheet.nrows):
                rowinfo = sheet.rowinfo_map.get(row_idx)
                if rowinfo:
                    logger.debug(f"  row {row_idx}: height={rowinfo.height}")
                if rowinfo and rowinfo.height == 0:
                    logger.info(f"  非表示行検出: {sheet_name} 行{row_idx}")
                    hidden_rows.append((sheet_name, row_idx))

            # 列の幅が0 → 非表示（colinfo_map は sheet 単位）
            for col_idx, colinfo in sheet.colinfo_map.items():
                logger.debug(f"  col {col_idx}: width={colinfo.width}")
                if colinfo.width == 0:
                    logger.info(f"  非表示列検出: {sheet_name} 列{col_idx}")
                    hidden_cols.append((sheet_name, col_idx))

        return hidden_rows, hidden_cols

    except Exception as e:
        logger.exception(f"非表示行・列チェックでエラー: {e}")
        return [], []


def check_no_hidden_rows_or_columns(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    # .xls の場合
    if workbook is None:
        hidden_rows, hidden_cols = check_xls_hidden_rows_columns(Path(filepath))

        row_str = (
            ", ".join(f"{sheet}シートの{r+1}行" for sheet, r in hidden_rows)
            if hidden_rows else "該当なし"
        )
        col_str = (
            ", ".join(f"{sheet}シートの{get_excel_column_letter(c+1)}列" for sheet, c in hidden_cols)
            if hidden_cols else "該当なし"
        )

        if hidden_rows or hidden_cols:
            return False, f"非表示行／列があります（行: {row_str}, 列: {col_str}）"
        return True, "非表示行／列はありません"

    # .xlsx の場合
    ws = workbook[ctx.sheet_name]
    hidden_rows = [d.index for d in ws.row_dimensions.values() if d.hidden]
    hidden_cols = [d.index for d in ws.column_dimensions.values() if d.hidden]

    row_str = (
        ", ".join(f"{r}行" for r in hidden_rows) if hidden_rows else "該当なし"
    )
    col_str = (
        ", ".join(f"{get_excel_column_letter(c)}列" for c in hidden_cols) if hidden_cols else "該当なし"
    )

    if hidden_rows or hidden_cols:
        return False, f"非表示行／列があります（行: {row_str}, 列: {col_str}）"
    return True, "非表示行／列はありません"


def check_no_notes_outside_table(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    if not ctx.upper_annotations.empty or not ctx.lower_annotations.empty:
        column_rows = ctx.row_indices.get("column_rows")
        data_end = ctx.row_indices.get("data_end")

        if column_rows is None or data_end is None:
            return False, "注釈行のチェックに必要な情報が不足しています"

        top = column_rows[0] if isinstance(column_rows, list) else cast(int, column_rows)
        bottom = cast(int, data_end) + 2
        return False, f"注釈行が検出されました（{top}行目より前、または{bottom}行目以降）"
    return True, "表外の注釈や備考はありません"


def check_no_merged_cells(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    # .xlsファイルの場合はxlrdを使用してチェック
    if workbook is None:
        merged_ranges = check_xls_merged_cells(Path(filepath))
        
        if merged_ranges:
            return False, f"結合セルが検出されました: {merged_ranges}"
        return True, "結合セルはありません"
    
    ws = workbook[ctx.sheet_name]
    column_rows = ctx.row_indices.get("column_rows")
    data_end = ctx.row_indices.get("data_end")

    if column_rows is None or data_end is None:
        return False, "結合セルチェックに必要な情報が不足しています"

    start = min(column_rows) + 1 if isinstance(column_rows, list) else cast(int, column_rows) + 1
    end = cast(int, data_end) + 1

    relevant_merges = [
        str(rng)
        for rng in ws.merged_cells.ranges
        if rng.min_row >= start and rng.max_row <= end
    ]

    if relevant_merges:
        return False, f"結合セルが検出されました: {relevant_merges}"
    return True, "結合セルはありません"


def check_no_format_based_semantics(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    # .xlsファイルの場合はxlrdを使用してチェック
    if workbook is None:
        data_start = ctx.row_indices.get("data_start", 0)
        data_end = ctx.row_indices.get("data_end", len(ctx.data) - 1)
        
        flagged = check_xls_cell_formats(Path(filepath), ctx.sheet_name, data_start, data_end)
        
        if flagged:
            return False, f"視覚的装飾による意味付けが検出されました（例: {flagged[:MAX_EXAMPLES]}）"
        return True, "書式ベースの意味づけは検出されませんでした"
    
    ws = workbook[ctx.sheet_name]
    column_rows = ctx.row_indices.get("column_rows")
    data_end = ctx.row_indices.get("data_end")

    if column_rows is None or data_end is None:
        return False, "書式チェックに必要な情報が不足しています"

    start = min(column_rows) + 1 if isinstance(column_rows, list) else cast(int, column_rows) + 1
    end = cast(int, data_end) + 1

    flagged = []
    for row in ws.iter_rows(min_row=start, max_row=end):
        for cell in row:
            coord = cell.coordinate
            fill = cell.fill
            if fill and fill.fgColor:
                fg = fill.fgColor
                if hasattr(fg, "rgb") and isinstance(fg.rgb, str):
                    rgb = fg.rgb.upper()
                    if rgb not in ("00000000", "FFFFFFFF", "FF000000"):
                        flagged.append(f"{coord}（塗りつぶし）")

            font = cell.font
            if font:
                if font.color and hasattr(font.color, "rgb") and isinstance(font.color.rgb, str):
                    rgb = font.color.rgb.upper()
                    if rgb not in ("00000000", "FF000000"):
                        flagged.append(f"{coord}（文字色）")

                if font.bold:
                    flagged.append(f"{coord}（太字）")
                if font.italic:
                    flagged.append(f"{coord}（イタリック）")
                if font.underline:
                    flagged.append(f"{coord}（下線）")
                if font.sz and (font.sz < 9 or font.sz > 13):
                    flagged.append(f"{coord}（フォントサイズ {font.sz}）")

    if flagged:
        return False, f"視覚的装飾による意味付けが検出されました（例: {flagged[:MAX_EXAMPLES]}）"
    return True, "書式ベースの意味づけは検出されませんでした"


def check_no_whitespace_formatting(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    # .xlsファイルの場合はワークブックがNoneになるため、DataFrameベースでチェック
    if workbook is None:
        # DataFrameを使用した簡易版の空白チェック
        sample_cells = []
        for row_idx, row in ctx.data.iterrows():
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and "　" in val:
                    col_letter = get_excel_column_letter(col_idx + 1)
                    cell_ref = f"{col_letter}{row_idx + 1}"
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
    
    ws = workbook[ctx.sheet_name]
    column_rows = ctx.row_indices.get("column_rows")
    data_end = ctx.row_indices.get("data_end")

    if column_rows is None or data_end is None:
        return False, "空白チェックに必要な情報が不足しています"

    start = min(column_rows) + 1 if isinstance(column_rows, list) else cast(int, column_rows) + 1
    end = cast(int, data_end) + 1

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


def check_single_data_per_cell(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    pattern = re.compile(r"[\n,;/]")
    problems = []

    data_start = ctx.row_indices.get("data_start")
    if data_start is None:
        return False, "データ開始位置が不明です"

    start: int = cast(int, data_start)
    for row_idx_raw, row in ctx.data.iterrows():
        row_idx: int = cast(int, row_idx_raw)
        for col_idx, val in enumerate(row):
            if isinstance(val, str) and pattern.search(val):
                excel_row = row_idx + 1 + start
                excel_col_letter = get_excel_column_letter(col_idx + 1)
                coord = f"{excel_col_letter}{excel_row}"
                problems.append(f"{coord}: {repr(val)}")

    if problems:
        return False, f"複数データセルが検出されました（例: {problems[:MAX_EXAMPLES]}）"
    return True, "各セルに1データのみです"


def check_no_platform_dependent_characters(
    ctx: TableContext, workbook: Workbook, filepath: str
) -> Tuple[bool, str]:
    # .xlsファイルの場合はワークブックがNoneになるため、DataFrameベースでチェック
    if workbook is None:
        # DataFrameを使用した簡易版の機種依存文字チェック
        issues = []
        for row_idx, row in ctx.data.iterrows():
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and detect_platform_characters(val):
                    coord = f"{get_excel_column_letter(col_idx + 1)}{row_idx + 1}"
                    issues.append(f"{coord}: '{val}'")

        if issues:
            return False, f"機種依存文字が含まれています（例: {issues[:MAX_EXAMPLES]}）"
        return True, "機種依存文字は含まれていません"
    
    ws = workbook[ctx.sheet_name]
    column_rows = ctx.row_indices.get("column_rows")
    data_end = ctx.row_indices.get("data_end")

    if column_rows is None or data_end is None:
        return False, "機種依存文字チェックに必要な情報が不足しています"

    if isinstance(column_rows, list):
        start = min(column_rows) + 1
    else:
        start = cast(int, column_rows) + 1
    end = cast(int, data_end) + 1

    issues = []
    for r, row in enumerate(ws.iter_rows(min_row=start, max_row=end, values_only=True), start=start):
        for c, val in enumerate(row, start=1):
            if isinstance(val, str) and detect_platform_characters(val):
                coord = f"{get_excel_column_letter(c)}{r}"
                issues.append(f"{coord}: '{val}'")

    if issues:
        return False, f"機種依存文字が含まれています（例: {issues[:MAX_EXAMPLES]}）"
    return True, "機種依存文字は含まれていません"
