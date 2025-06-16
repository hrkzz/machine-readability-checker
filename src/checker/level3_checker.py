from typing import Tuple, Optional
import pandas as pd
from openpyxl.workbook.workbook import Workbook
from pathlib import Path

from src.processor.context import TableContext
from .factory import checker_factory


def check_code_format_for_choices(
    ctx: TableContext, workbook: Optional[Workbook] = None, filepath: Optional[str] = None
) -> Tuple[bool, str]:
    """選択肢のコード形式チェック"""
    file_path = Path(filepath) if filepath else None
    if file_path is None:
        return False, "ファイルパスが指定されていません"
    
    checker = checker_factory.get_level3_checker(file_path)
    return checker.check_code_format_for_choices(ctx, workbook, filepath)


def check_codebook_exists(
    ctx: TableContext, workbook: Optional[Workbook], filepath: Optional[str]
) -> Tuple[bool, str]:
    """コード表の存在チェック"""
    file_path = Path(filepath) if filepath else None
    if file_path is None:
        return False, "ファイルパスが指定されていません"
    
    checker = checker_factory.get_level3_checker(file_path)
    return checker.check_codebook_exists(ctx, workbook, filepath)


def check_question_master_exists(
    ctx: TableContext, workbook: Optional[Workbook], filepath: Optional[str]
) -> Tuple[bool, str]:
    """設問マスターの存在チェック"""
    file_path = Path(filepath) if filepath else None
    if file_path is None:
        return False, "ファイルパスが指定されていません"
    
    checker = checker_factory.get_level3_checker(file_path)
    return checker.check_question_master_exists(ctx, workbook, filepath)


def check_metadata_presence(
    ctx: TableContext, workbook: Optional[Workbook], filepath: Optional[str]
) -> Tuple[bool, str]:
    """メタデータの存在チェック"""
    file_path = Path(filepath) if filepath else None
    if file_path is None:
        return False, "ファイルパスが指定されていません"
    
    checker = checker_factory.get_level3_checker(file_path)
    return checker.check_metadata_presence(ctx, workbook, filepath)


def check_long_format_if_many_columns(
    ctx: TableContext, workbook: Optional[Workbook] = None, filepath: Optional[str] = None
) -> Tuple[bool, str]:
    """列数が多い場合の縦型形式チェック"""
    file_path = Path(filepath) if filepath else None
    if file_path is None:
        return False, "ファイルパスが指定されていません"
    
    checker = checker_factory.get_level3_checker(file_path)
    return checker.check_long_format_if_many_columns(ctx, workbook, filepath)