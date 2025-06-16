import re
from pathlib import Path
from typing import Any, List, Dict, Tuple
import zipfile
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from ..llm.llm_client import call_llm

# xlrdライブラリの読み込み
try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False

MAX_EXAMPLES = 10

def get_excel_column_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def get_xls_workbook_info(file_path: Path) -> Dict[str, Any]:
    """
    .xlsファイルからxlrdを使用して詳細情報を取得
    """
    if not XLRD_AVAILABLE:
        return {}
    
    try:
        wb = xlrd.open_workbook(str(file_path), formatting_info=True)
        return {
            'workbook': wb,
            'sheet_count': wb.nsheets,
            'sheet_names': wb.sheet_names()
        }
    except Exception as e:
        print(f"xlsファイルの詳細情報取得でエラー: {e}")
        return {}

def check_xls_hidden_rows_columns(file_path: Path, sheet_name: str) -> Tuple[List[int], List[int]]:
    """
    .xlsファイルの非表示行・列をチェック
    """
    if not XLRD_AVAILABLE:
        return [], []
    
    try:
        wb = xlrd.open_workbook(str(file_path), formatting_info=True)
        sheet = wb.sheet_by_name(sheet_name)
        
        hidden_rows = []
        hidden_cols = []
        
        # 行の非表示チェック
        if hasattr(sheet, 'rowinfo_map'):
            for row_idx, row_info in sheet.rowinfo_map.items():
                if row_info.hidden:
                    hidden_rows.append(row_idx)
        
        # 列の非表示チェック
        if hasattr(sheet, 'colinfo_map'):
            for col_idx, col_info in sheet.colinfo_map.items():
                if col_info.hidden:
                    hidden_cols.append(col_idx)
        
        return hidden_rows, hidden_cols
    except Exception as e:
        print(f"非表示行・列チェックでエラー: {e}")
        return [], []

def check_xls_merged_cells(file_path: Path, sheet_name: str) -> List[str]:
    """
    .xlsファイルの結合セルをチェック
    """
    if not XLRD_AVAILABLE:
        return []
    
    try:
        wb = xlrd.open_workbook(str(file_path), formatting_info=True)
        sheet = wb.sheet_by_name(sheet_name)
        
        merged_ranges = []
        if hasattr(sheet, 'merged_cells'):
            for rlo, rhi, clo, chi in sheet.merged_cells:
                # Excelのセル範囲表記に変換
                start_cell = f"{get_excel_column_letter(clo + 1)}{rlo + 1}"
                end_cell = f"{get_excel_column_letter(chi)}{rhi}"
                merged_ranges.append(f"{start_cell}:{end_cell}")
        
        return merged_ranges
    except Exception as e:
        print(f"結合セルチェックでエラー: {e}")
        return []

def check_xls_cell_formats(file_path: Path, sheet_name: str, data_start: int, data_end: int) -> List[str]:
    """
    .xlsファイルのセル書式をチェック
    """
    if not XLRD_AVAILABLE:
        return []
    
    try:
        wb = xlrd.open_workbook(str(file_path), formatting_info=True)
        sheet = wb.sheet_by_name(sheet_name)
        
        flagged = []
        
        for row_idx in range(data_start, min(data_end + 1, sheet.nrows)):
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                
                # セルの書式情報を取得
                cell_xf = wb.format_map.get(cell.xf_index)
                if cell_xf:
                    font = wb.font_list[cell_xf.font_index]
                    
                    coord = f"{get_excel_column_letter(col_idx + 1)}{row_idx + 1}"
                    
                    # 太字チェック
                    if font.bold:
                        flagged.append(f"{coord}（太字）")
                    
                    # イタリックチェック
                    if font.italic:
                        flagged.append(f"{coord}（イタリック）")
                    
                    # 下線チェック
                    if font.underline_type:
                        flagged.append(f"{coord}（下線）")
                    
                    # 色チェック（基本色以外）
                    if font.colour_index not in (0, 1, 7, 8):  # 自動、黒、白以外
                        flagged.append(f"{coord}（文字色）")
        
        return flagged
    except Exception as e:
        print(f"書式チェックでエラー: {e}")
        return []

def has_any_drawing(path: Path) -> bool:
    """
    Excel ファイルに図形やオブジェクトが含まれているかをチェック
    .xls ファイルは構造上チェックが困難なため常に False を返す
    """
    ext = path.suffix.lower()
    if ext == ".xls":
        # .xls ファイルは構造上図形チェックが困難なため、
        # 図形がないものとして扱う（必要に応じて後で対応）
        return False
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

def detect_multiple_tables_dataframe(df: pd.DataFrame, column_rows: List[int], data_start: int, data_end: int) -> Tuple[bool, int]:
    """
    DataFrameベースで複数テーブルを検出
    """
    try:
        # データ部分を抽出
        data_section = df.iloc[data_start:data_end + 1]
        
        # 各行にデータが存在するかチェック
        has_data_flags = []
        for _, row in data_section.iterrows():
            # 行にデータが存在するかチェック（空文字、NaN以外）
            has_data = any(
                str(cell).strip() != "" and str(cell).lower() != "nan" 
                for cell in row
            )
            has_data_flags.append(has_data)
        
        # 連続するデータブロックを検出
        in_block = False
        blocks = 0
        
        for has_data in has_data_flags:
            if has_data and not in_block:
                blocks += 1
                in_block = True
            elif not has_data:
                in_block = False
        
        return blocks > 1, blocks
    except Exception as e:
        print(f"複数テーブル検出でエラー: {e}")
        return False, 1
