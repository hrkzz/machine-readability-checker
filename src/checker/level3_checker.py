from typing import Tuple, Optional
import pandas as pd
from openpyxl.workbook.workbook import Workbook

from src.processor.context import TableContext
from src.checker.utils import is_sheet_likely, is_likely_long_format

def check_code_format_for_choices(
    ctx: TableContext, workbook: Optional[Workbook] = None, filepath: Optional[str] = None
) -> Tuple[bool, str]:
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

def check_codebook_exists(
    ctx: TableContext, workbook: Optional[Workbook], filepath: Optional[str]
) -> Tuple[bool, str]:
    if workbook is None:
        return False, "Workbook が None です"

    for sheet in workbook.worksheets:
        if sheet.title == ctx.sheet_name:
            continue
        if is_sheet_likely(sheet, "コード表"):
            return True, f"コード表とみられるシート: {sheet.title}"
    return False, "コード表が見つかりません"

def check_question_master_exists(
    ctx: TableContext, workbook: Optional[Workbook], filepath: Optional[str]
) -> Tuple[bool, str]:
    if workbook is None:
        return False, "Workbook が None です"

    for sheet in workbook.worksheets:
        if sheet.title == ctx.sheet_name:
            continue
        if is_sheet_likely(sheet, "設問マスター"):
            return True, f"設問マスターとみられるシート: {sheet.title}"
    return False, "設問マスター（変数定義表）が見つかりません"

def check_metadata_presence(
    ctx: TableContext, workbook: Optional[Workbook], filepath: Optional[str]
) -> Tuple[bool, str]:
    if workbook is None:
        return False, "Workbook が None です"

    for sheet in workbook.worksheets:
        if sheet.title == ctx.sheet_name:
            continue
        if is_sheet_likely(sheet, "メタ情報"):
            return True, f"メタ情報とみられるシート: {sheet.title}"
    return False, "調査概要やメタデータが確認できません"

def check_long_format_if_many_columns(
    ctx: TableContext, workbook: Optional[Workbook] = None, filepath: Optional[str] = None
) -> Tuple[bool, str]:
    if is_likely_long_format(ctx.data):
        return True, "縦型（long format）とみなされます"
    return False, "wide型であり、long型形式ではありません"