import re
import pandas as pd
from ..llm.llm_client import call_llm
from .utils import is_sheet_likely

def check_code_format_for_choices(df=None, workbook=None, filepath=None):
    """
    選択肢カラムが数値コード（例：1=男性, 2=女性）で表記されているかをチェック。
    値が10種類未満かつ、コードではなくラベル（文字列）が混ざっている列を検出。
    """
    if df is None or not hasattr(df, "columns"):
        return False, "エラー: 有効な DataFrame が渡されていません"

    candidate_cols = []

    for col in df.columns:
        try:
            series = df[col]

            # カラムが重複していて DataFrame になってしまう場合はスキップ
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


def check_codebook_exists(df=None, workbook=None, filepath=None):
    if workbook is None:
        return False, "エラー: 有効な workbook が渡されていません"
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
    if workbook is None:
        return False, "エラー: 有効な workbook が渡されていません"
    for sheet in workbook.worksheets:
        if is_sheet_likely(sheet.title, "設問マスターや変数定義"):
            return True, f"設問マスターとみられるシート: {sheet.title}"
    return False, "設問マスター（変数定義表）が見つかりません"

def check_metadata_presence(df=None, workbook=None, filepath=None):
    if workbook is None:
        return False, "エラー: 有効な workbook が渡されていません"
    for sheet in workbook.worksheets:
        # シート名に基づく推論
        if is_sheet_likely(sheet.title, "調査概要やメタデータ"):
            return True, f"メタ情報とみられるシート: {sheet.title}"

        # 内容ベースの判定
        text_chunks = []
        for row in sheet.iter_rows(min_row=1, max_row=20, values_only=True):
            for cell in row:
                if isinstance(cell, str):
                    text = cell.strip()
                    # 短すぎる・列ラベルっぽいのは除外（「地域 都道府県」など）
                    if len(text) > 10 and not re.match(r"^[\w\s　・･、,]+$", text):
                        text_chunks.append(text)

        if text_chunks:
            sample = "\n".join(text_chunks[:5])
            prompt = f"""
                    次の文章群は、Excelシートから抽出されたセルの内容です。

                    この中に「調査概要・メタデータ」に該当する情報
                    （例：調査時期、出典、対象、単位、備考、調査方法、問合せ先など）は含まれますか？

                    注意：
                    - 単なる列見出し（例：地域、都道府県、性別、年齢など）は **メタデータではありません**
                    - 明らかに表データの一部と思われるものも **メタデータではありません**

                    回答は「YES」または「NO」のどちらか一語で、厳密に判断してください。

                    --- サンプル ---
                    {sample}
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
