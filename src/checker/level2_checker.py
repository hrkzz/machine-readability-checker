from typing import Tuple
import pandas as pd

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

        for col_idx, col in enumerate(df.columns, start=1):
            series = df[col].dropna()
            if series.empty:
                continue

            total = len(series)
            ok_count = series.apply(is_clean_numeric).sum()
            if ok_count / total < 0.8:
                continue

            if ok_count / total < 0.99:
                for row_idx, val in zip(series.index, series):
                    if not is_clean_numeric(val):
                        coord = f"{get_excel_column_letter(col_idx)}{row_idx + 1}"
                        problem_cells.setdefault(col, []).append(f"{coord}: '{val}'")

        if problem_cells:
            msgs = [
                f"{col}: {cells[:MAX_EXAMPLES]}" for col, cells in problem_cells.items()
            ]
            return False, "数値列に数値以外が含まれています:\n" + "\n".join(msgs)

        return True, "数値列に不正なデータは含まれていません"
    
    def check_separate_other_detail_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """選択肢列と自由記述の分離チェック（全形式共通）"""
        df = ctx.data
        flagged = []

        for col_idx, col in enumerate(df.columns, start=1):
            if not pd.api.types.is_string_dtype(df[col]):
                continue

            series = df[col].dropna().astype(str)
            if series.empty:
                continue

            if series.str.contains(FREE_TEXT_PATTERN).any():
                flagged.append(f"{col}（列: {get_excel_column_letter(col_idx)}）")

        if flagged:
            return False, f"選択肢列に自由記述が混在している可能性があります: {flagged}"
        return True, "選択肢列と自由記述は適切に分離されています"
    
    def check_no_missing_column_headers(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """列ヘッダーの明確性チェック（全形式共通）"""
        df = ctx.data
        suspect = [c for c in df.columns if "Unnamed" in str(c) or str(c).strip() == ""]

        for col in df.columns:
            if col in suspect:
                continue
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

        if suspect:
            return False, f"省略・不明な列名が検出されました: {suspect}"
        return True, "全ての列に意味のあるヘッダーが付いています"
    
    def check_handling_of_missing_values(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """欠損値の統一性チェック（全形式共通）"""
        file_ext = filepath.split('.')[-1].upper() if '.' in filepath else "UNKNOWN"
        self.logger.debug(f"{file_ext} Level2 欠損値チェック開始: {filepath}")
        
        df = ctx.data
        flagged = []

        for idx, col in enumerate(df.columns, start=1):
            self.logger.debug(f"列 '{col}' ({idx}) を処理中")
            
            # 明示的にSeriesとして取得してuniqueを呼び出す
            column_series = df[col]
            if isinstance(column_series, pd.DataFrame):
                # 万一DataFrameが返される場合はスキップ
                self.logger.warning(f"列 '{col}' が DataFrame として取得されました。スキップします。")
                continue
            
            self.logger.debug(f"列 '{col}' の型: {type(column_series)}")
            
            try:
                # より安全なユニーク値の取得
                if hasattr(column_series, 'dropna') and hasattr(column_series, 'unique'):
                    unique_vals = [
                        str(v).strip() for v in column_series.dropna().unique()
                        if str(v).strip()
                    ]
                else:
                    # 直接的なアプローチが失敗した場合の代替手段
                    unique_vals = [
                        str(v).strip() for v in pd.Series(column_series).dropna().unique()
                        if str(v).strip()
                    ]
                self.logger.debug(f"列 '{col}' のユニーク値数: {len(unique_vals)}")
            except Exception as e:
                self.logger.error(f"列 '{col}' のユニーク値取得でエラー: {e}")
                continue
            
            if not unique_vals:
                continue

            sample_vals = unique_vals[:MAX_EXAMPLES * 4]
            prompt = f"""
            以下は列「{col}」のユニークな値サンプルです。
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
                return False, f"列「{col}」の欠損表現チェックで LLM 呼び出しエラー: {e}"

            if "欠損表現あり" in res:
                flagged.append(f"{col}（列: {get_excel_column_letter(idx)}）")

        self.logger.debug(f"{file_ext} Level2 欠損値チェック完了: フラグ数={len(flagged)}")
        if flagged:
            return False, f"欠損表現が不統一な可能性のある列が見つかりました: {flagged}"
        return True, "全列の欠損表現は一貫しています" 