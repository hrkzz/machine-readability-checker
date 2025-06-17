from typing import Tuple
import pandas as pd
import openpyxl

from src.processor.context import TableContext
from src.checker.base.base_checker import BaseLevel3Checker
from src.checker.utils.common import is_likely_long_format, is_sheet_likely


class XLSXLevel3Checker(BaseLevel3Checker):
    """
    XLSX専用のLevel3チェッカー
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/xlsx_checker3.log", rotation="10 MB", retention="30 days", level="DEBUG")
    
    def get_supported_file_types(self) -> set:
        return {".xlsx"}
    
    def check_code_format_for_choices(self, ctx: TableContext, workbook: openpyxl.Workbook, filepath: str) -> Tuple[bool, str]:
        """XLSX専用の選択肢のコード化チェック"""
        df = ctx.data
        candidate_cols = []

        for col in df.columns:
            try:
                series = df[col]
                if isinstance(series, pd.DataFrame):
                    continue

                unique_vals = series.dropna().unique()
                if len(unique_vals) < 10:
                    if any(not str(val).isdigit() for val in unique_vals):
                        candidate_cols.append(col)
            except Exception as e:
                return False, f"列 '{col}' でエラー発生: {e}"

        if candidate_cols:
            return False, f"コード表記ではない可能性のある列: {candidate_cols}"
        return True, "選択肢はコード表記されています"
    
    def check_codebook_exists(self, ctx: TableContext, workbook: openpyxl.Workbook, filepath: str) -> Tuple[bool, str]:
        """XLSX専用のコード表の存在チェック"""
        if workbook is None:
            return False, "ワークブック情報が不足しているためチェックできません"
        
        for sheet in workbook.worksheets:
            if sheet.title == ctx.sheet_name:
                continue
            if is_sheet_likely(sheet, "コード表"):
                return True, f"コード表とみられるシート: {sheet.title}"
        return False, "コード表が見つかりません"
    
    def check_question_master_exists(self, ctx: TableContext, workbook: openpyxl.Workbook, filepath: str) -> Tuple[bool, str]:
        """XLSX専用の設問マスターの存在チェック"""
        if workbook is None:
            return False, "ワークブック情報が不足しているためチェックできません"
        
        for sheet in workbook.worksheets:
            if sheet.title == ctx.sheet_name:
                continue
            if is_sheet_likely(sheet, "設問マスター"):
                return True, f"設問マスターとみられるシート: {sheet.title}"
        return False, "設問マスター（変数定義表）が見つかりません"
    
    def check_metadata_presence(self, ctx: TableContext, workbook: openpyxl.Workbook, filepath: str) -> Tuple[bool, str]:
        """XLSX専用のメタデータの存在チェック"""
        if workbook is None:
            return False, "ワークブック情報が不足しているためチェックできません"
        
        for sheet in workbook.worksheets:
            if sheet.title == ctx.sheet_name:
                continue
            if is_sheet_likely(sheet, "メタ情報"):
                return True, f"メタ情報とみられるシート: {sheet.title}"
        return False, "調査概要やメタデータが確認できません"
    
    def check_long_format_if_many_columns(self, ctx: TableContext, workbook: openpyxl.Workbook, filepath: str) -> Tuple[bool, str]:
        """XLSX専用のlong format形式のチェック"""
        if is_likely_long_format(ctx.data):
            return True, "縦型（long format）とみなされます"
        return False, "wide型であり、long型形式ではありません" 