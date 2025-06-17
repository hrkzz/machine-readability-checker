from pathlib import Path
import re
from typing import Tuple, cast

from src.processor.context import TableContext
from src.checker.base.base_checker import BaseChecker
from src.checker.utils.common import (
    detect_platform_characters,
    get_excel_column_letter,
    MAX_EXAMPLES,
)
from src.checker.utils.xls_utils import (
    check_xls_merged_cells,
    check_xls_cell_formats,
    check_xls_hidden_rows_columns
)
from src.checker.utils.csv_utils import detect_multiple_tables_csv
from src.llm.llm_client import call_llm


class XLSLevel1Checker(BaseChecker):
    """
    XLS専用のLevel1チェッカー
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/xls_checker1.log", rotation="10 MB", retention="30 days", level="DEBUG")
    
    def get_supported_file_types(self) -> set:
        return {".xls"}
    
    def check_valid_file_format(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有のファイル形式チェック"""
        ext = Path(filepath).suffix.lower()
        if ext != ".xls":
            return False, f"サポート外のファイル形式です: {ext}"
        return True, "旧Excel（.xls）形式のため、一部の自動チェック（書式・図形など）が制限されます。必要に応じて目視での確認を行ってください"
    
    def check_no_images_or_objects(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有のオブジェクトチェック"""
        # .xlsファイルでは図形・オブジェクトの詳細検出は困難だが、
        # 一般的に統計表では図形は使用されないため、警告として扱う
        return False, "xlsファイルでは図形や画像の自動判定ができません。必要に応じて目視でご確認ください"
    
    def check_one_table_per_sheet(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の複数テーブルチェック"""
        # .xlsファイルの場合はDataFrameベースでチェック
        is_multiple, details = detect_multiple_tables_csv(ctx.data, ctx.sheet_name)
        
        if is_multiple:
            return False, f"複数テーブルの疑いがあります: {details}"
        return True, "1つのテーブルのみです"
    
    def check_no_hidden_rows_or_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の非表示行・列チェック"""
        hidden_rows, hidden_cols = check_xls_hidden_rows_columns(Path(filepath))

        row_str = (
            ", ".join(f"{sheet}シートの{r+1}行" for sheet, r in hidden_rows)
            if hidden_rows else "該当なし"
        )
        col_str = (
            ", ".join(f"{sheet}シートの{get_excel_column_letter(c+1)}列" for sheet, c in hidden_cols)
            if hidden_cols else "該当なし"
        )

        if hidden_rows or hidden_cols:
            return False, f"非表示行／列があります（行: {row_str}, 列: {col_str}）"
        return True, "非表示行／列はありません"
    
    def check_no_notes_outside_table(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の表外注釈チェック"""
        if not ctx.upper_annotations.empty or not ctx.lower_annotations.empty:
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "注釈行のチェックに必要な情報が不足しています"

            top = column_rows[0] if isinstance(column_rows, list) else cast(int, column_rows)
            bottom = cast(int, data_end) + 2
            return False, f"注釈行が検出されました（{top}行目より前、または{bottom}行目以降）"
        return True, "表外の注釈や備考はありません"
    
    def check_no_merged_cells(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の結合セルチェック"""
        # テーブル本体の開始・終了行を ctx から取得
        column_rows = ctx.row_indices.get("column_rows")
        data_end    = ctx.row_indices.get("data_end")
        if column_rows is None or data_end is None:
            return False, "結合セルチェックに必要な情報が不足しています"

        start = min(column_rows) if isinstance(column_rows, list) else cast(int, column_rows)
        end   = cast(int, data_end)

        merged = check_xls_merged_cells(Path(filepath), ctx.sheet_name, start, end)
        if merged:
            return False, f"結合セルが検出されました: {merged}"
        else:
            return True, "結合セルはありません"
    
    def check_no_format_based_semantics(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の書式チェック"""
        data_start = ctx.row_indices.get("data_start", 0)
        data_end = ctx.row_indices.get("data_end", len(ctx.data) - 1)
        
        flagged = check_xls_cell_formats(Path(filepath), ctx.sheet_name, data_start, data_end)
        
        if flagged:
            return False, f"視覚的装飾による意味付けが検出されました（例: {flagged[:MAX_EXAMPLES]}）"
        return True, "書式ベースの意味づけは検出されませんでした"
    
    def check_no_whitespace_formatting(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の空白チェック"""
        # DataFrameを使用した簡易版の空白チェック
        sample_cells = []
        for row_idx, row in ctx.data.iterrows():
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and "　" in val:
                    col_letter = get_excel_column_letter(col_idx + 1)
                    cell_ref = f"{col_letter}{row_idx + 1}"
                    sample_cells.append(f"{cell_ref}: '{val.strip()}'")
                    if len(sample_cells) >= 10:
                        break
            if len(sample_cells) >= 10:
                break

        if not sample_cells:
            return True, "体裁調整目的の空白は見つかりませんでした"

        prompt = f"""
            以下はExcel(.xls)のセル値の一部です。これらの中に、見た目を整える目的（位置揃え・スペース調整など）で
            **空白（特に全角スペース）が使われているものがあるか**を判定してください。

            データ:
            {chr(10).join(sample_cells)}

            判断結果を次のいずれか一語で返してください：
            - 調整目的あり
            - 調整目的なし
        """

        result = call_llm(prompt)
        if "調整目的あり" in result:
            return False, f"体裁調整目的の空白が含まれている可能性があります（例: {sample_cells[:MAX_EXAMPLES]}）"
        return True, "体裁調整目的の空白は見つかりませんでした"
    
    def check_single_data_per_cell(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の1セル1データチェック"""
        pattern = re.compile(r"[\n,;/]")
        problems = []

        data_start = ctx.row_indices.get("data_start")
        if data_start is None:
            return False, "データ開始位置が不明です"

        start: int = cast(int, data_start)
        for row_idx_raw, row in ctx.data.iterrows():
            row_idx: int = cast(int, row_idx_raw)
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and pattern.search(val):
                    excel_row = row_idx + 1 + start
                    excel_col_letter = get_excel_column_letter(col_idx + 1)
                    coord = f"{excel_col_letter}{excel_row}"
                    problems.append(f"{coord}: {repr(val)}")

        if problems:
            return False, f"複数データセルが検出されました（例: {problems[:MAX_EXAMPLES]}）"
        return True, "各セルに1データのみです"
    
    def check_no_platform_dependent_characters(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """XLS固有の機種依存文字チェック"""
        # DataFrameを使用した簡易版の機種依存文字チェック
        issues = []
        for row_idx, row in ctx.data.iterrows():
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and detect_platform_characters(val):
                    coord = f"{get_excel_column_letter(col_idx + 1)}{row_idx + 1}"
                    issues.append(f"{coord}: '{val}'")

        if issues:
            return False, f"機種依存文字が含まれています（例: {issues[:MAX_EXAMPLES]}）"
        return True, "機種依存文字は含まれていません" 