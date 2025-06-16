import xlrd
from pathlib import Path
from loguru import logger
from .common import get_excel_column_letter


def get_xls_workbook_info(file_path: Path) -> dict:
    """xlsファイルの基本情報を取得"""
    try:
        workbook = xlrd.open_workbook(str(file_path))
        
        sheet_info = []
        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)
            sheet_info.append({
                'name': sheet_name,
                'nrows': sheet.nrows,
                'ncols': sheet.ncols
            })
        
        return {
            'file_path': str(file_path),
            'nsheets': workbook.nsheets,
            'sheet_names': workbook.sheet_names(),
            'sheet_info': sheet_info
        }
    except Exception as e:
        logger.error(f"xlsファイルの詳細情報取得でエラー: {e}")
        return {}


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
    """xlsファイルのセル書式をチェック"""
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


def check_xls_hidden_rows_columns(file_path: Path) -> tuple:
    """xlsファイルの非表示行・列をチェック"""
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