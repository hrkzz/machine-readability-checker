from abc import ABC, abstractmethod
from typing import Tuple, Optional
import pandas as pd
from pathlib import Path
from loguru import logger

from src.processor.context import TableContext


class BaseChecker(ABC):
    """
    チェッカーの基底クラス
    各ファイル形式（CSV、XLS、XLSX）に共通するインターフェースを定義
    """
    
    def __init__(self):
        self.logger = logger
    
    @abstractmethod
    def get_supported_file_types(self) -> set:
        """
        サポートするファイル拡張子を返す
        """
        pass
    
    def can_handle(self, file_path: Path) -> bool:
        """
        このチェッカーがファイルを処理できるかチェック
        """
        return file_path.suffix.lower() in self.get_supported_file_types()
    
    @abstractmethod
    def check_valid_file_format(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """ファイル形式の妥当性チェック"""
        pass
    
    @abstractmethod
    def check_no_images_or_objects(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """画像・オブジェクトの存在チェック"""  
        pass
    
    @abstractmethod
    def check_one_table_per_sheet(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """1シート1テーブルのチェック"""
        pass
    
    @abstractmethod
    def check_no_hidden_rows_or_columns(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """非表示行・列のチェック"""
        pass
    
    @abstractmethod
    def check_no_notes_outside_table(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """表外注釈のチェック"""
        pass
    
    @abstractmethod
    def check_no_merged_cells(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """結合セルのチェック"""
        pass
    
    @abstractmethod
    def check_no_format_based_semantics(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """書式による意味付けのチェック"""
        pass
    
    @abstractmethod
    def check_no_whitespace_formatting(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """空白による体裁調整のチェック"""
        pass
    
    @abstractmethod
    def check_single_data_per_cell(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """1セル1データのチェック"""
        pass
    
    @abstractmethod
    def check_no_platform_dependent_characters(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """機種依存文字のチェック"""
        pass


class BaseLevel2Checker(ABC):
    """
    Level2チェッカーの基底クラス
    """
    
    def __init__(self):
        self.logger = logger
    
    @abstractmethod
    def get_supported_file_types(self) -> set:
        """
        サポートするファイル拡張子を返す
        """
        pass
    
    def can_handle(self, file_path: Path) -> bool:
        """
        このチェッカーがファイルを処理できるかチェック
        """
        return file_path.suffix.lower() in self.get_supported_file_types()
    
    @abstractmethod
    def check_numeric_columns_only(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """数値列の妥当性チェック"""
        pass
    
    @abstractmethod
    def check_separate_other_detail_columns(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """選択肢列と自由記述の分離チェック"""
        pass
    
    @abstractmethod
    def check_no_missing_column_headers(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """列ヘッダーの明確性チェック"""
        pass
    
    @abstractmethod
    def check_handling_of_missing_values(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """欠損値の統一性チェック"""
        pass


class BaseLevel3Checker(ABC):
    """
    Level3チェッカーの基底クラス
    """
    
    def __init__(self):
        self.logger = logger
    
    @abstractmethod
    def get_supported_file_types(self) -> set:
        """
        サポートするファイル拡張子を返す
        """
        pass
    
    def can_handle(self, file_path: Path) -> bool:
        """
        このチェッカーがファイルを処理できるかチェック
        """
        return file_path.suffix.lower() in self.get_supported_file_types()
    
    @abstractmethod
    def check_code_format_for_choices(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """選択肢のコード化チェック"""
        pass
    
    @abstractmethod
    def check_codebook_exists(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """コード表の存在チェック"""
        pass
    
    @abstractmethod
    def check_question_master_exists(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """設問マスターの存在チェック"""
        pass
    
    @abstractmethod
    def check_metadata_presence(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """メタデータの存在チェック"""
        pass
    
    @abstractmethod
    def check_long_format_if_many_columns(self, ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """long format形式のチェック"""
        pass 