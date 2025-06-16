from typing import Tuple
import pandas as pd
from pathlib import Path
from loguru import logger

from src.processor.context import TableContext
from src.checker.base.base_checker import BaseLevel3Checker
from src.checker.utils.common import is_likely_long_format


class XLSLevel3Checker(BaseLevel3Checker):
    """
    XLS専用のLevel3チェッカー
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/xls_checker3.log", rotation="10 MB", retention="30 days", level="DEBUG")
    
    def get_supported_file_types(self) -> set:
        return {".xls"}
    
    def check_code_format_for_choices(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS専用の選択肢のコード化チェック"""
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
    
    def check_codebook_exists(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS専用のコード表の存在チェック"""
        # .xlsファイルの場合、pandas ExcelFileを使用してシート情報を取得
        if filepath and filepath.lower().endswith('.xls'):
            try:
                xl_file = pd.ExcelFile(filepath, engine='xlrd')
                for sheet_name in xl_file.sheet_names:
                    if sheet_name == ctx.sheet_name:
                        continue
                    
                    # 簡易的なシート内容チェック
                    try:
                        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=10, engine='xlrd')
                        if not df.empty:
                            # シート名や内容からコード表らしさを判定
                            if any(keyword in sheet_name.lower() for keyword in ['code', 'コード', 'master', 'マスタ']):
                                return True, f"コード表とみられるシート: {sheet_name}"
                    except Exception:
                        continue
                        
                return False, "コード表が見つかりません（.xlsファイルでは詳細検索は制限されます）"
            except Exception as e:
                return False, f".xlsファイルのシート検索でエラー: {e}"
        else:
            return False, ".xlsファイルでないためコード表チェックをスキップ"
    
    def check_question_master_exists(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS専用の設問マスターの存在チェック"""
        # .xlsファイルの場合、pandas ExcelFileを使用してシート情報を取得
        if filepath and filepath.lower().endswith('.xls'):
            try:
                xl_file = pd.ExcelFile(filepath, engine='xlrd')
                for sheet_name in xl_file.sheet_names:
                    if sheet_name == ctx.sheet_name:
                        continue
                    
                    # 簡易的なシート内容チェック
                    try:
                        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=10, engine='xlrd')
                        if not df.empty:
                            # シート名や内容から設問マスターらしさを判定
                            if any(keyword in sheet_name.lower() for keyword in ['question', '設問', 'master', 'マスタ', 'variable', '変数']):
                                return True, f"設問マスターとみられるシート: {sheet_name}"
                    except Exception:
                        continue
                        
                return False, "設問マスター（変数定義表）が見つかりません（.xlsファイルでは詳細検索は制限されます）"
            except Exception as e:
                return False, f".xlsファイルのシート検索でエラー: {e}"
        else:
            return False, ".xlsファイルでないため設問マスターチェックをスキップ"
    
    def check_metadata_presence(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS専用のメタデータの存在チェック"""
        # .xlsファイルの場合、pandas ExcelFileを使用してシート情報を取得
        if filepath and filepath.lower().endswith('.xls'):
            try:
                xl_file = pd.ExcelFile(filepath, engine='xlrd')
                for sheet_name in xl_file.sheet_names:
                    if sheet_name == ctx.sheet_name:
                        continue
                    
                    # 簡易的なシート内容チェック
                    try:
                        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=10, engine='xlrd')
                        if not df.empty:
                            # シート名や内容からメタ情報らしさを判定
                            if any(keyword in sheet_name.lower() for keyword in ['meta', 'メタ', 'info', '情報', '概要', 'readme']):
                                return True, f"メタ情報とみられるシート: {sheet_name}"
                    except Exception:
                        continue
                        
                return False, "調査概要やメタデータが確認できません（.xlsファイルでは詳細検索は制限されます）"
            except Exception as e:
                return False, f".xlsファイルのシート検索でエラー: {e}"
        else:
            return False, ".xlsファイルでないためメタデータチェックをスキップ"
    
    def check_long_format_if_many_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS専用のlong format形式のチェック"""
        if is_likely_long_format(ctx.data):
            return True, "縦型（long format）とみなされます"
        return False, "wide型であり、long型形式ではありません" 