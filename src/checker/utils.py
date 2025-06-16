import re
from pathlib import Path
from typing import Any
import zipfile
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from ..llm.llm_client import call_llm
from loguru import logger
import xlrd
MAX_EXAMPLES = 10

def get_excel_column_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def get_xls_workbook_info(file_path: Path) -> dict:
    """xlsファイルの基本情報を取得"""
    try:
        workbook = xlrd.open_workbook(str(file_path))
        
        sheet_info = []
        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)
            sheet_info.append({
                'name': sheet_name,
                'nrows': sheet.nrows,
                'ncols': sheet.ncols
            })
        
        return {
            'file_path': str(file_path),
            'nsheets': workbook.nsheets,
            'sheet_names': workbook.sheet_names(),
            'sheet_info': sheet_info
        }
    except Exception as e:
        logger.error(f"xlsファイルの詳細情報取得でエラー: {e}")
        return {}



def has_any_drawing(path: Path) -> bool:
    """
    Excel ファイルに図形やオブジェクトが含まれているかをチェック
    .xls ファイルは構造上チェックが困難なため常に False を返す
    """
    ext = path.suffix.lower()
    if ext == ".xls":
        # .xls ファイルは構造上図形チェックが困難なため、
        # 図形があるものとして扱う（必要に応じて後で対応）
        return True
    elif ext != ".xlsx":
        return False
    
    try:
        with zipfile.ZipFile(path, 'r') as z:
            for name in z.namelist():
                if name.startswith('xl/drawings/') and name.endswith('.xml'):
                    xml = z.read(name)
                    if b'<xdr:twoCellAnchor' in xml or b'<xdr:oneCellAnchor' in xml:
                        return True
    except Exception:
        return False
    return False

def detect_platform_characters(text: str) -> bool:
    pattern = re.compile(r"[①-⑳⓪-⓿Ⅰ-Ⅻ㊤㊥㊦㊧㊨㈱㈲㈹℡〒〓※]")
    return bool(pattern.search(text))

def is_clean_numeric(val: Any) -> bool:
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

MISSING_VALUE_EXPRESSIONS = [
    "不明", "不詳", "無記入", "無回答", "該当なし", "なし", "無し", "n/a", "na", "nan",
    "未定", "未記入", "未入力", "未回答", "記載なし", "対象外", "空欄", "空白", "不在",
    "特になし", "---", "--", "-", "ー", "―", "？", "?", "わからない", "わかりません",
    "なし（特記なし）", "無し（詳細不明）", "無効", "省略", "null", "none"
]

def is_sheet_likely(sheet: Worksheet, category: str) -> bool:
    text_lines = []
    for row in sheet.iter_rows(min_row=1, max_row=15, values_only=True):
        line = " ".join(str(cell).strip() for cell in row if cell)
        if line:
            text_lines.append(line)

    if not text_lines:
        return False

    sample_text = "\n".join(text_lines[:10])
    prompt = f"""
        以下はExcelシート「{sheet.title}」の冒頭行の内容です：

        {sample_text}

        このシートは「{category}」に該当しますか？

        カテゴリの意味：
        - コード表: 数値コードとラベルの対応表（例: 1=男性, 2=女性 など）
        - 設問マスター: 変数名、設問文、選択肢などの設問一覧表
        - メタ情報: 調査概要、出典、単位、調査時期など、表データ以外の補足情報

        列見出しやデータの一部であっても、該当カテゴリに沿っていれば「YES」としてください。

        回答は必ず「YES」または「NO」のみで返してください。
    """
    result = call_llm(prompt)
    return "YES" in result.upper()

def is_likely_long_format(df: pd.DataFrame) -> bool:
    """
    ID・変数名・値 を含み、列数が多い DataFrame を縦型とみなす。
    """
    if len(df.columns) < 10:
        return False
    return {"ID", "変数名", "値"}.issubset(set(df.columns))

# 後方互換性のために元の関数名も維持
def has_any_drawing_xlsx(path: Path) -> bool:
    """後方互換性のための関数（has_any_drawingを使用することを推奨）"""
    return has_any_drawing(path)


