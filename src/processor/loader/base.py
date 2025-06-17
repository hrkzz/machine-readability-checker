from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, Any
import pandas as pd
from loguru import logger

from src.config import PREVIEW_ROW_COUNT


class BaseLoader(ABC):
    """
    ファイルローダーの基底クラス
    各ファイル形式（CSV、XLS、XLSX）に共通するインターフェースを定義
    """
    
    def __init__(self):
        self.logger = logger
        self.logger.add("logs/file_loader.log", rotation="10 MB", retention="30 days", level="DEBUG")
    
    @abstractmethod
    def get_supported_extensions(self) -> set:
        """
        サポートするファイル拡張子を返す
        """
        pass
    
    @abstractmethod
    def load_file_internal(self, file_path: Path) -> Dict[str, Any]:
        """
        ファイル形式固有の読み込み処理
        """
        pass
    
    def can_handle(self, file_path: Path) -> bool:
        """
        このローダーがファイルを処理できるかチェック
        """
        return file_path.suffix.lower() in self.get_supported_extensions()
    
    def load_file(self, file_path: Path) -> Dict[str, Any]:
        """
        ファイルを読み込み、各シートのデータを返す
        """
        if not self.can_handle(file_path):
            raise ValueError(f"Unsupported file format: {file_path.suffix}")
        
        try:
            return self.load_file_internal(file_path)
        except Exception as e:
            self.logger.error(f"ファイル読み込みエラー ({file_path}): {e}")
            raise
    
    def create_sheet_info(self, sheet_name: str, df: pd.DataFrame) -> Dict[str, Any]:
        """
        シート情報の共通フォーマットを作成
        """
        return {
            "sheet_name": sheet_name,
            "dataframe": df,
            "preview_top": df.head(PREVIEW_ROW_COUNT),
            "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
        }
    
    def create_result_structure(self, file_path: Path, file_type: str, sheets: list) -> Dict[str, Any]:
        """
        結果の共通構造を作成
        """
        return {
            "file_path": file_path,
            "file_type": file_type,
            "sheets": sheets
        } 