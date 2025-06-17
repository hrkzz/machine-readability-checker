import re
from typing import Any
import pandas as pd
import openpyxl

# 共通定数
MAX_EXAMPLES = 10

def get_excel_column_letter(n: int) -> str:
    """Excel列番号をアルファベットに変換"""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def detect_platform_characters(text: str) -> bool:
    """機種依存文字を検出"""
    pattern = re.compile(r"[①-⑳⓪-⓿Ⅰ-Ⅻ㊤㊥㊦㊧㊨㈱㈲㈹℡〒〓※]")
    return bool(pattern.search(text))

def is_clean_numeric(val: Any) -> bool:
    """値が純粋な数値かどうかをチェック"""
    if isinstance(val, (int, float)):
        return True
    if isinstance(val, str):
        s = val.strip()
        if re.search(r"[^\d.\-]", s):
            return False
        try:
            float(s)
            return True
        except ValueError:
            return False
    return False

def is_likely_long_format(df: pd.DataFrame) -> bool:
    """
    ID・変数名・値 を含み、列数が多い DataFrame を縦型とみなす。
    """
    if len(df.columns) < 10:
        return False
    return {"ID", "変数名", "値"}.issubset(set(df.columns))

def is_sheet_likely(sheet: openpyxl.worksheet.worksheet.Worksheet, sheet_type: str) -> bool:
    """
    シートが指定されたタイプ（コード表、設問マスター、メタ情報）である可能性を判定
    """
    sheet_name = sheet.title.lower()
    
    # シート名での判定
    if sheet_type == "コード表":
        keywords = ['code', 'コード', 'master', 'マスタ', 'codebook', 'code_book']
    elif sheet_type == "設問マスター":
        keywords = ['question', '設問', 'master', 'マスタ', 'variable', '変数', 'var', 'item', '項目']
    elif sheet_type == "メタ情報":
        keywords = ['meta', 'メタ', 'info', '情報', '概要', 'readme', 'read_me', 'summary', '要約']
    else:
        return False
    
    # シート名に関連キーワードが含まれているかチェック
    if any(keyword in sheet_name for keyword in keywords):
        return True
    
    # シート内容の簡易チェック（先頭10行程度）
    try:
        values = []
        for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):
            for cell in row:
                if cell is not None:
                    values.append(str(cell).lower())
        
        content = " ".join(values)
        
        # 内容にキーワードが含まれているかチェック
        if sheet_type == "コード表" and any(keyword in content for keyword in ['コード', 'code', '値', 'value']):
            return True
        elif sheet_type == "設問マスター" and any(keyword in content for keyword in ['設問', 'question', '変数', 'variable']):
            return True
        elif sheet_type == "メタ情報" and any(keyword in content for keyword in ['メタ', 'meta', '概要', 'summary']):
            return True
    except Exception:
        # エラーが発生した場合はFalseを返す
        pass
    
    return False

# 自由記述パターン
FREE_TEXT_PATTERN = re.compile(r"""
    ^\s*(?:  
        (?:その他|そのほか)\s*[:：\-\–\/]           |
        (?:その他|そのほか)\s*[\(（].+?[\)）]       |
        コメント\s*[:：]                            |
        自由記述\s*[:：]                            |
        詳細\s*[:：]                                |
        備考\s*[:：]                                |
        補足\s*[:：]                                |
        感想\s*[:：]                                |
        意見\s*[:：]                                |
        メモ\s*[:：]                                |
        特記事項\s*[:：]                            |
        注釈\s*[:：]                                |
        自己PR\s*[:：]                              |
        フリーテキスト\s*[:：]                      |
        フリー回答\s*[:：]
    )
""", re.VERBOSE)

# 欠損値表現
MISSING_VALUE_EXPRESSIONS = [
    "不明", "不詳", "無記入", "無回答", "該当なし", "なし", "無し", "n/a", "na", "nan",
    "未定", "未記入", "未入力", "未回答", "記載なし", "対象外", "空欄", "空白", "不在",
    "特になし", "---", "--", "-", "ー", "―", "？", "?", "わからない", "わかりません",
    "なし（特記なし）", "無し（詳細不明）", "無効", "省略", "null", "none"
] 