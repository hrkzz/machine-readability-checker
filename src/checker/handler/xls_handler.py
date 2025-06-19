import xlrd
from pathlib import Path
from loguru import logger
from src.checker.common import get_excel_column_letter


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

        # 結果サマリー
        if merged_cells:
            logger.warning(f"check_xls_merged_cells: 結合セル検出 - {sheet_name}: {merged_cells}")
        else:
            logger.info(f"check_xls_merged_cells: OK - {sheet_name}")

        return merged_cells

    except Exception as e:
        logger.error(f"check_xls_merged_cells: 処理中にエラーが発生 ({sheet_name}): {e}")
        return []


def check_xls_cell_formats(file_path: Path, sheet_name: str, data_start: int, data_end: int) -> list:
    """xlsファイルのセル書式をチェック（座標修正版）"""
    try:
        logger.debug(f"check_xls_cell_formats: 開始 - {file_path}, sheet: {sheet_name}")
        workbook = xlrd.open_workbook(str(file_path), formatting_info=True)
        sheet = workbook.sheet_by_name(sheet_name)
        flagged = []

        # data_startとdata_endは0ベースのDataFrameインデックスなので、
        # 実際のExcel行番号に変換する必要がある
        for df_row_idx in range(data_start, min(data_end + 1, len(sheet.rows))):
            # Excel行番号は df_row_idx + 実際のデータ開始行
            # ここでは簡易的にdf_row_idx + 2（ヘッダー考慮）とする
            excel_row_idx = df_row_idx + 2  # 通常、1行目がヘッダー、2行目からデータ
            
            if excel_row_idx >= sheet.nrows:
                break
                
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(excel_row_idx, col_idx)
                xf_index = cell.xf_index

                if xf_index >= len(workbook.xf_list):
                    continue  # 異常なインデックスはスキップ

                xf = workbook.xf_list[xf_index]
                font_index = xf.font_index

                if font_index >= len(workbook.font_list):
                    continue

                font = workbook.font_list[font_index]
                # Excel座標（1ベース）で表示
                coord = f"{get_excel_column_letter(col_idx + 1)}{excel_row_idx + 1}"

                # 太字
                if font.bold:
                    flagged.append(f"{coord}（太字）")
                # イタリック
                if font.italic:
                    flagged.append(f"{coord}（イタリック）")
                # 下線
                if font.underline_type != 0:
                    flagged.append(f"{coord}（下線）")
                # 文字色（デフォルト色以外）
                if font.colour_index not in (0, 1, 7, 8):  # 自動・黒・白以外
                    flagged.append(f"{coord}（文字色）")
                # 背景色（デフォルト色以外）
                bg_index = xf.background.pattern_colour_index
                if bg_index not in (64, 0):  # 標準色以外
                    flagged.append(f"{coord}（背景色）")

        # 結果サマリー
        if flagged:
            logger.warning(f"check_xls_cell_formats: 書式付きセル検出 - {sheet_name}: {len(flagged)}件")
        else:
            logger.info(f"check_xls_cell_formats: OK - {sheet_name}")

        return flagged

    except Exception as e:
        logger.error(f"check_xls_cell_formats: 処理中にエラーが発生 - {sheet_name}: {e}")
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
            sheet_hidden_rows = []
            for row_idx in range(sheet.nrows):
                rowinfo = sheet.rowinfo_map.get(row_idx)
                if rowinfo and rowinfo.height == 0:
                    sheet_hidden_rows.append(row_idx)
                    hidden_rows.append((sheet_name, row_idx))
            
            if sheet_hidden_rows:
                logger.info(f"非表示行検出: {sheet_name} 行{sheet_hidden_rows}")

            # 列の幅が0 → 非表示（colinfo_map は sheet 単位）
            sheet_hidden_cols = []
            for col_idx, colinfo in sheet.colinfo_map.items():
                if colinfo.width == 0:
                    col_letter = get_excel_column_letter(col_idx + 1)
                    sheet_hidden_cols.append(col_letter)
                    hidden_cols.append((sheet_name, col_idx))
            
            if sheet_hidden_cols:
                logger.info(f"非表示列検出: {sheet_name} 列{sheet_hidden_cols}")

        # チェック結果サマリー
        if hidden_rows or hidden_cols:
            logger.warning(f"check_xls_hidden_rows_columns: 非表示要素検出 - 行:{len(hidden_rows)}件, 列:{len(hidden_cols)}件")
        else:
            logger.info("check_xls_hidden_rows_columns: OK")

        return hidden_rows, hidden_cols

    except Exception as e:
        logger.error(f"check_xls_hidden_rows_columns: 処理中にエラーが発生: {e}")
        return [], [] 