from pathlib import Path
from typing import Dict, Any
import pandas as pd
import xlrd

from .base import BaseLoader


class XLSLoader(BaseLoader):
    """
    XLSファイル専用のローダー
    """
    
    def get_supported_extensions(self) -> set:
        return {".xls"}
    
    def load_file_internal(self, file_path: Path) -> Dict[str, Any]:
        """
        XLSファイルの読み込み処理
        .xls形式はxlrd 1.2.0で処理する（pandasを使わない）
        """
        self.logger.info(f"=== .xls ファイル読み込み開始: {file_path} ===")
        
        sheets = []
        try:
            wb = xlrd.open_workbook(str(file_path), formatting_info=True)
            self.logger.info(f"xlrdで開いたワークブック: {wb.nsheets} シート")

            for sheet in wb.sheets():
                self.logger.info(f"=== シート '{sheet.name}' を処理中 ===")
                rows = []
                for row_idx in range(sheet.nrows):
                    row = sheet.row_values(row_idx)
                    rows.append(row)

                df = pd.DataFrame(rows)
                self.logger.info(f"xlrdで構築したDataFrame: shape={df.shape}")

                sheet_info = self.create_sheet_info(sheet.name, df)
                sheets.append(sheet_info)

        except Exception as e:
            self.logger.error(f".xls読み込みでエラー: {e}")
            self.logger.exception("詳細なエラー情報:")
            # エラー時は空のデータフレームを返す
            sheet_info = self.create_sheet_info("Sheet1", pd.DataFrame())
            sheets.append(sheet_info)
        
        return self.create_result_structure(file_path, ".xls", sheets) 