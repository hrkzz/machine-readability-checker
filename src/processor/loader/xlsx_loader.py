from pathlib import Path
from typing import Dict, Any
import pandas as pd
import openpyxl

from .base import BaseLoader


class XLSXLoader(BaseLoader):
    """
    XLSXファイル専用のローダー
    """
    
    def get_supported_extensions(self) -> set:
        return {".xlsx"}
    
    def load_file_internal(self, file_path: Path) -> Dict[str, Any]:
        """
        XLSXファイルの読み込み処理
        """
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheets = []
        
        for ws in wb.worksheets:
            try:
                # ヘッダーは後で LLM が判定するため、ここでは header=None
                df = pd.read_excel(file_path, sheet_name=ws.title, header=None)
                self.logger.info(f"xlsx シート '{ws.title}' 読み込み完了: shape={df.shape}")
            except Exception as e:
                self.logger.error(f"xlsx シート '{ws.title}' の読み込みでエラー: {e}")
                # 読み込みに失敗した場合は空 DataFrame
                df = pd.DataFrame()

            sheet_info = self.create_sheet_info(ws.title, df)
            sheets.append(sheet_info)
        
        return self.create_result_structure(file_path, ".xlsx", sheets) 