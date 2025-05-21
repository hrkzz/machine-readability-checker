import pandas as pd
import re
from typing import List, Dict, Any, Tuple
from ..llm.llm_client import call_llm

def analyze_cell_content(cell_value: Any) -> Dict[str, Any]:
    if pd.isna(cell_value):
        return {
            "is_empty": True,
            "has_multiple_lines": False,
            "has_special_chars": False,
            "word_count": 0
        }

    cell_str = str(cell_value)
    return {
        "is_empty": False,
        "has_multiple_lines": "\n" in cell_str,
        "has_special_chars": bool(re.search(r'[^\w\s\u3000-\u9FFF]', cell_str)),
        "word_count": len(cell_str.split())
    }

def is_likely_comment_column(column_name: str) -> bool:
    comment_keywords = ["備考", "コメント", "注釈", "メモ", "note", "comment", "remarks"]
    return any(keyword in str(column_name).lower() for keyword in comment_keywords)

def get_empty_rows_columns(df: pd.DataFrame) -> Tuple[List[int], List[str]]:
    empty_rows = df.index[df.isna().all(axis=1)].tolist()
    empty_cols = df.columns[df.isna().all(axis=0)].tolist()
    return empty_rows, empty_cols

def count_merged_cells(worksheet) -> int:
    return len(list(worksheet.merged_cells.ranges))

def get_hidden_rows_columns(worksheet) -> Tuple[List[int], List[str]]:
    hidden_rows = [idx + 1 for idx, row in enumerate(worksheet.row_dimensions) 
                   if worksheet.row_dimensions[row].hidden]
    hidden_cols = [chr(65 + idx) for idx, col in enumerate(worksheet.column_dimensions) 
                   if worksheet.column_dimensions[col].hidden]
    return hidden_rows, hidden_cols

def detect_platform_characters(text: str) -> bool:
    pattern = re.compile(r"[①-⑳⓪-⓿Ⅰ-Ⅻ㊤㊥㊦㊧㊨㈱㈲㈹℡〒]")
    return bool(pattern.search(text))

def detect_notes_outside_table(texts: List[str]) -> str:
    """
    表の外側にあるテキストが備考や注釈であるかどうかをLLMに判定させる
    """
    prompt = f"""
以下のテキストは、表の外にある可能性があります。これらが「注釈・備考・説明」のような補助情報かどうかを判定してください。
結果は「注釈あり」または「注釈なし」のいずれかで日本語で答えてください。

{texts[:10]}
"""
    return call_llm(prompt)

def is_sheet_likely(sheet_title: str, category: str, sheet=None) -> bool:
    prompt = f"""
以下のExcelシート名「{sheet_title}」は「{category}」に関係がある名前ですか？

- 例えば、「コード表」「設問一覧」「変数定義」「調査概要」などは関係があります。
- 一方で、「データ」「入力用」「結果」「集計」などは通常、関係ありません。

関係があるなら YES、ないなら NO とだけ答えてください。
"""
    result = call_llm(prompt)
    if "YES" not in result.upper():
        return False

    # 二重確認：シート内容にも "コード", "性別", "1=", などがあるか
    if sheet:
        sample_lines = []
        for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
            line = " ".join(str(cell) for cell in row if cell)
            sample_lines.append(line)
        content = "\n".join(sample_lines)
        if "コード" in content or re.search(r"\b1\s*[=：]\s*\w+", content):
            return True
        else:
            return False

    return True
