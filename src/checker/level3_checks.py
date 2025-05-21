import re
from ..llm.llm_client import call_llm
from .utils import is_sheet_likely
def check_code_format_for_choices(df=None, workbook=None, filepath=None):
    candidate_cols = []
    for col in df.columns:
        unique_vals = df[col].dropna().unique()
        if len(unique_vals) < 10:
            if any(not str(val).isdigit() for val in unique_vals):
                candidate_cols.append(col)

    if candidate_cols:
        return False, f"コード表記ではない可能性のある列: {candidate_cols}"
    return True, "選択肢はコード表記されています"

def check_codebook_exists(df=None, workbook=None, filepath=None):
    for sheet in workbook.worksheets:
        if is_sheet_likely(sheet.title, "コード表"):
            return True, f"コード表とみられるシート: {sheet.title}"

        # 内容も見て補強判断（1=○○のような表記）
        for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):
            line = " ".join([str(cell) for cell in row if cell])
            if re.search(r"\b1\s*[=：]\s*\w+", line):
                return True, f"内容からコード表と推定されるシート: {sheet.title}"
    return False, "コード表が見つかりません"

def check_question_master_exists(df=None, workbook=None, filepath=None):
    for sheet in workbook.worksheets:
        if is_sheet_likely(sheet.title, "設問マスターや変数定義"):
            return True, f"設問マスターとみられるシート: {sheet.title}"
    return False, "設問マスター（変数定義表）が見つかりません"

def check_metadata_presence(df=None, workbook=None, filepath=None):
    for sheet in workbook.worksheets:
        if is_sheet_likely(sheet.title, "調査概要やメタデータ"):
            return True, f"メタ情報とみられるシート: {sheet.title}"

        # 内容もチェック
        text_chunks = []
        for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):
            for cell in row:
                if isinstance(cell, str) and len(cell.strip()) > 5:
                    text_chunks.append(cell.strip())
        if text_chunks:
            prompt = f"""
以下はExcelシート内の文章です。これは調査概要やメタデータ（例：出典、単位、調査対象など）に関係していますか？

{text_chunks[:5]}

関係していれば YES、していなければ NO と一語で答えてください。
"""
            response = call_llm(prompt)
            if "YES" in response.upper():
                return True, f"内容からメタデータが見つかりました（例: {text_chunks[0]}）"
    return False, "調査概要やメタデータが確認できません"

def check_long_format_if_many_columns(df=None, workbook=None, filepath=None):
    if len(df.columns) < 10:
        return True, "列数が多くないため縦型要件は不要です"

    required_cols = {"ID", "変数名", "値"}
    if required_cols.issubset(set(df.columns)):
        return True, "縦型（long format）とみなされます"
    else:
        return False, "wide型であり、long型形式ではありません"
