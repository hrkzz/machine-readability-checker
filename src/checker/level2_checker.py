from typing import Tuple
import pandas as pd
from openpyxl.workbook.workbook import Workbook
from pathlib import Path

from src.processor.context import TableContext
from src.llm.llm_client import call_llm
from src.checker.utils import (
    get_excel_column_letter,
    MAX_EXAMPLES,
    is_clean_numeric,
    FREE_TEXT_PATTERN,
    MISSING_VALUE_EXPRESSIONS
)
from .factory import checker_factory


def check_numeric_columns_only(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """数値列の妥当性チェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level2_checker(file_path)
    return checker.check_numeric_columns_only(ctx, workbook, filepath)


def check_separate_other_detail_columns(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """選択肢列と自由記述の分離チェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level2_checker(file_path)
    return checker.check_separate_other_detail_columns(ctx, workbook, filepath)


def check_no_missing_column_headers(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """列ヘッダーの明確性チェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level2_checker(file_path)
    return checker.check_no_missing_column_headers(ctx, workbook, filepath)


def check_handling_of_missing_values(ctx: TableContext, workbook, filepath: str) -> Tuple[bool, str]:
    """欠損値の統一性チェック"""
    file_path = Path(filepath)
    checker = checker_factory.get_level2_checker(file_path)
    return checker.check_handling_of_missing_values(ctx, workbook, filepath)
