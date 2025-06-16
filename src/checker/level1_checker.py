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
from .factory import checker_factory

logger.add("logs/checker1.log", rotation="10 MB", retention="30 days", level="DEBUG")


def check_xls_merged_cells(
    file_path: Path,
    sheet_name: str,
    row_start: int,
    row_end: int
) -> list:
    """xlsファイルの結合セルをチェック（指定範囲内のみ）"""
    try:
        wb = xlrd.open_workbook(str(file_path), formatting_info=True)
        sheet = wb.sheet_by_name(sheet_name)
        merged_cells = []

        # xlrdのmerged_cellsは (r0, r1, c0, c1) 形式。r1/c1 は「1つ先」のインデックス。
        for r0, r1, c0, c1 in getattr(sheet, 'merged_cells', []):
            last_row = r1 - 1
            # テーブル本体の行範囲内かどうか
            if r0 >= row_start and last_row <= row_end:
                top_left  = f"{get_excel_column_letter(c0 + 1)}{r0 + 1}"
                bottom_right = f"{get_excel_column_letter(c1)}{r1}"
                merged_cells.append(f"{top_left}:{bottom_right}")

        return merged_cells

    except Exception as e:
        logger.error(f"check_xls_merged_cells エラー ({sheet_name}): {e}")
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


def check_valid_file_format(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """ファイル形式の妥当性チェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_valid_file_format(ctx, workbook, filepath)


def check_no_images_or_objects(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """画像・オブジェクトの存在チェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_images_or_objects(ctx, workbook, filepath)


def check_one_table_per_sheet(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """1シート1テーブルのチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_one_table_per_sheet(ctx, workbook, filepath)


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


def check_no_hidden_rows_or_columns(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """非表示行・列のチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_hidden_rows_or_columns(ctx, workbook, filepath)


def check_no_notes_outside_table(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """表外注釈のチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_notes_outside_table(ctx, workbook, filepath)


def check_no_merged_cells(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """結合セルのチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_merged_cells(ctx, workbook, filepath)


def check_no_format_based_semantics(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """書式ベースの意味づけチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_format_based_semantics(ctx, workbook, filepath)


def check_no_whitespace_formatting(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """空白による体裁調整のチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_whitespace_formatting(ctx, workbook, filepath)


def check_single_data_per_cell(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """1セル1データのチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_single_data_per_cell(ctx, workbook, filepath)


def check_no_platform_dependent_characters(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """機種依存文字のチェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level1_checker(file_path)
    return checker.check_no_platform_dependent_characters(ctx, workbook, filepath)
