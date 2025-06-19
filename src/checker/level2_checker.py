from typing import Tuple
import pandas as pd
import warnings

from src.processor.context import TableContext
from src.checker.base_checker import BaseLevel2Checker
from src.checker.common import (
    get_excel_column_letter,
    MAX_EXAMPLES,
    is_clean_numeric,
    FREE_TEXT_PATTERN,
    MISSING_VALUE_EXPRESSIONS
)
from src.llm.llm_client import call_llm

# PerformanceWarningを抑制
warnings.simplefilter("ignore", pd.errors.PerformanceWarning)


class Level2Checker(BaseLevel2Checker):
    """
    Level2チェッカー
    CSV、XLS、XLSXの全ファイル形式に対応
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/level2_checker.log", rotation="10 MB", retention="30 days", level="DEBUG")
    
    def get_supported_file_types(self) -> set:
        return {".csv", ".xls", ".xlsx"}
    
    def check_numeric_columns_only(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """数値列の妥当性チェック（全形式共通）"""
        df = ctx.data
        problem_cells = {}
        numeric_columns_checked = 0

        for col_idx, col in enumerate(df.columns, start=1):
            series = df[col].dropna()
            if series.empty:
                continue

            total = len(series)
            ok_count = series.apply(is_clean_numeric).sum()
            if ok_count / total < 0.8:
                continue

            numeric_columns_checked += 1
            if ok_count / total < 0.99:
                for row_idx, val in zip(series.index, series):
                    if not is_clean_numeric(val):
                        coord = f"{get_excel_column_letter(col_idx)}{row_idx + 1}"
                        problem_cells.setdefault(col, []).append(f"{coord}: '{val}'")

        # チェック結果サマリー
        if problem_cells:
            # router.pyでログ出力されるため、ここでの詳細ログは省略
            msgs = [
                f"{col}: {cells[:MAX_EXAMPLES]}" for col, cells in problem_cells.items()
            ]
            return False, "数値列に数値以外が含まれています:\n" + "\n".join(msgs)
        else:
            # router.pyでログ出力されるため、ここでの詳細ログは省略
            return True, "数値列に不正なデータは含まれていません"
    
    def check_separate_other_detail_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """選択肢列と自由記述の分離チェック（全形式共通）"""
        df = ctx.data
        flagged = []
        string_columns_checked = 0

        for col_idx, col in enumerate(df.columns, start=1):
            if not pd.api.types.is_string_dtype(df[col]):
                continue

            series = df[col].dropna().astype(str)
            if series.empty:
                continue

            string_columns_checked += 1
            if series.str.contains(FREE_TEXT_PATTERN).any():
                flagged.append(f"{col}（列: {get_excel_column_letter(col_idx)}）")

        # チェック結果サマリー（router.pyでログ出力）
        if flagged:
            return False, f"選択肢列に自由記述が混在している可能性があります: {flagged}"
        else:
            return True, "選択肢列と自由記述は適切に分離されています"
    
    def check_no_missing_column_headers(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """列ヘッダーの明確性チェック（全形式共通）"""
        df = ctx.data
        suspect = [c for c in df.columns if "Unnamed" in str(c) or str(c).strip() == ""]
        llm_checked_columns = 0

        for col in df.columns:
            if col in suspect:
                continue
            
            llm_checked_columns += 1
            prompt = f"""
            以下の列名について、表の項目として
            - 意味が一義に理解できる → 「明確」  
            - 語義が推測できない、略称すぎる、記号やミスタイプ等 → 「不明瞭」

            回答は「明確」または「不明瞭」のみ：
            列名: {col}
            """
            try:
                res = call_llm(prompt)
            except Exception as e:
                return False, f"ヘッダー名チェックで LLM エラー: {e}"
            if "不明" in res:
                suspect.append(col)

        # チェック結果サマリー（router.pyでログ出力）
        if suspect:
            return False, f"省略・不明な列名が検出されました: {suspect}"
        else:
            return True, "全ての列に意味のあるヘッダーが付いています"
    
    def check_handling_of_missing_values(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """欠損値の統一性チェック（全形式共通）"""
        file_ext = filepath.split('.')[-1].upper() if '.' in filepath else "UNKNOWN"
        self.logger.debug(f"{file_ext} Level2 欠損値チェック開始: {filepath}")
        
        df = ctx.data
        flagged = []
        processed_columns = 0
        skipped_columns = 0

        # MultiIndex対応の改善
        columns_to_check = df.columns
        if isinstance(df.columns, pd.MultiIndex):
            # MultiIndexの場合は各レベルを組み合わせた文字列として扱う
            self.logger.debug(f"MultiIndex列が検出されました: {len(df.columns)}列")
            columns_to_check = [
                '_'.join(str(level) for level in col if str(level).strip() != '')
                for col in df.columns
            ]
        
        for idx, col in enumerate(df.columns, start=1):
            try:
                # 列の取得を改善（MultiIndex対応）
                if isinstance(df.columns, pd.MultiIndex):
                    # MultiIndexの場合は iloc を使用して安全に取得
                    column_series = df.iloc[:, idx-1]
                    col_name = columns_to_check[idx-1]
                else:
                    column_series = df[col]
                    col_name = str(col)
                
                # DataFrameが返される場合の対処
                if isinstance(column_series, pd.DataFrame):
                    # MultiIndexの場合、まれに複数列が返される可能性があるため
                    if column_series.shape[1] == 1:
                        column_series = column_series.iloc[:, 0]
                    else:
                        self.logger.debug(f"列 '{col_name}' が複数列のDataFrameとして取得されました。スキップします。")
                        skipped_columns += 1
                        continue
                
                # ユニーク値の安全な取得
                if not hasattr(column_series, 'dropna'):
                    column_series = pd.Series(column_series)
                
                unique_vals = [
                    str(v).strip() for v in column_series.dropna().unique()
                    if str(v).strip() and str(v).strip() != 'nan'
                ]
                
            except Exception as e:
                self.logger.debug(f"列 '{col}' の処理でエラー: {e}")
                skipped_columns += 1
                continue
            
            if not unique_vals:
                continue

            processed_columns += 1
            sample_vals = unique_vals[:MAX_EXAMPLES * 4]
            prompt = f"""
            以下は列「{col_name}」のユニークな値サンプルです。
            この中に、欠損を意味する語句が含まれ、それらの表現が一貫していないかを判断してください。

            欠損表現の例：
            {MISSING_VALUE_EXPRESSIONS}

            ▼サンプル値：
            {chr(10).join(sample_vals)}

            回答形式：「欠損表現あり」または「なし」
            """
            try:
                res = call_llm(prompt)
            except Exception as e:
                return False, f"列「{col_name}」の欠損表現チェックで LLM 呼び出しエラー: {e}"

            if "欠損表現あり" in res:
                flagged.append(f"{col_name}（列: {get_excel_column_letter(idx)}）")

        # チェック結果サマリー（router.pyでログ出力、スキップ情報は詳細ログのみ）
        if skipped_columns > 0:
            self.logger.debug(f"check_handling_of_missing_values: {skipped_columns}列をスキップしました")
        
        if flagged:
            return False, f"欠損表現が不統一な可能性のある列が見つかりました: {flagged}"
        else:
            return True, "全列の欠損表現は一貫しています" 