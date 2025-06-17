from pathlib import Path
from typing import Dict, Any
import pandas as pd

from .base import BaseLoader


class CSVLoader(BaseLoader):
    """
    CSVファイル専用のローダー
    """
    
    def get_supported_extensions(self) -> set:
        return {".csv"}
    
    def load_file_internal(self, file_path: Path) -> Dict[str, Any]:
        """
        CSVファイルの読み込み処理
        """
        try:
            df = pd.read_csv(file_path, header=None, encoding="cp932")
            self.logger.info(f"CSV読み込み完了: shape={df.shape}")
        except UnicodeDecodeError as e:
            self.logger.error(f"CSVのエンコーディングエラー: {e}")
            raise ValueError("CSVファイルのエンコーディングに問題があります。Shift_JIS（cp932）などを試してください。")
        
        sheet_info = self.create_sheet_info("CSV", df)
        return self.create_result_structure(file_path, ".csv", [sheet_info]) 