import re
from pathlib import Path
from typing import Any, List, Dict, Tuple
import zipfile
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from ..llm.llm_client import call_llm
from loguru import logger

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

def check_xls_hidden_rows_columns(file_path: Path) -> tuple:
    """xlsファイルの非表示行・列をチェック"""
    try:
        workbook = xlrd.open_workbook(str(file_path), formatting_info=True)
        hidden_rows = []
        hidden_cols = []
        
        for sheet_idx, sheet_name in enumerate(workbook.sheet_names()):
            sheet = workbook.sheet_by_index(sheet_idx)
            
            # 行の高さをチェック（高さ0の行は非表示）
            for row_idx in range(sheet.nrows):
                row_info = workbook.rowinfo_map.get(sheet_idx, {}).get(row_idx)
                if row_info and row_info.height == 0:
                    hidden_rows.append((sheet_name, row_idx))
            
            # 列の幅をチェック（幅0の列は非表示）
            for col_idx in range(sheet.ncols):
                col_info = workbook.colinfo_map.get(sheet_idx, {}).get(col_idx)
                if col_info and col_info.width == 0:
                    hidden_cols.append((sheet_name, col_idx))
        
        return hidden_rows, hidden_cols
    except Exception as e:
        logger.error(f"非表示行・列チェックでエラー: {e}")
        return [], []

def check_xls_merged_cells(file_path: Path) -> list:
    """xlsファイルの結合セルをチェック"""
    try:
        workbook = xlrd.open_workbook(str(file_path))
        merged_cells = []
        
        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)
            if hasattr(sheet, 'merged_cells'):
                for row_start, row_end, col_start, col_end in sheet.merged_cells:
                    merged_cells.append({
                        'sheet': sheet_name,
                        'range': f"{row_start}:{row_end-1}, {col_start}:{col_end-1}"
                    })
        
        return merged_cells
    except Exception as e:
        logger.error(f"結合セルチェックでエラー: {e}")
        return []

def check_xls_cell_formats(file_path: Path) -> dict:
    """xlsファイルのセル書式をチェック"""
    try:
        workbook = xlrd.open_workbook(str(file_path), formatting_info=True)
        format_info = {}
        
        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)
            sheet_formats = []
            
            # 各セルの書式をチェック
            for row_idx in range(min(sheet.nrows, 100)):  # 最初の100行のみチェック
                for col_idx in range(min(sheet.ncols, 50)):  # 最初の50列のみチェック
                    cell = sheet.cell(row_idx, col_idx)
                    cell_format = workbook.format_map.get(cell.xf_index)
                    
                    if cell_format:
                        font = workbook.font_list[cell_format.font_index]
                        if (font.bold or font.italic or font.underline_type != 0 or 
                            font.colour_index != 0 or cell_format.background.pattern_colour_index != 64):
                            sheet_formats.append({
                                'row': row_idx,
                                'col': col_idx,
                                'bold': font.bold,
                                'italic': font.italic,
                                'underline': font.underline_type != 0,
                                'color_index': font.colour_index,
                                'bg_color_index': cell_format.background.pattern_colour_index
                            })
            
            if sheet_formats:
                format_info[sheet_name] = sheet_formats
        
        return format_info
    except Exception as e:
        logger.error(f"書式チェックでエラー: {e}")
        return {}

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

def detect_multiple_tables_dataframe(df: pd.DataFrame, sheet_name: str = "") -> tuple:
    """
    DataFrameベースで複数テーブルを検出する
    
    Args:
        df: 対象のDataFrame
        sheet_name: シート名（ログ用）
    
    Returns:
        tuple: (has_multiple_tables: bool, details: str)
    """
    try:
        if df.empty or len(df) < 3:
            return False, "データが少ないため複数テーブルの検出をスキップ"
        
        # 完全に空の行を検索
        empty_rows = []
        for idx, row in df.iterrows():
            if row.isna().all() or (row.astype(str).str.strip() == "").all():
                empty_rows.append(idx)
        
        # 連続する空行を検索（テーブル区切りの可能性）
        if len(empty_rows) > 0:
            consecutive_groups = []
            current_group = [empty_rows[0]]
            
            for i in range(1, len(empty_rows)):
                if empty_rows[i] == empty_rows[i-1] + 1:
                    current_group.append(empty_rows[i])
                else:
                    if len(current_group) >= 2:  # 2行以上の連続空行
                        consecutive_groups.append(current_group)
                    current_group = [empty_rows[i]]
            
            if len(current_group) >= 2:
                consecutive_groups.append(current_group)
            
            if consecutive_groups:
                return True, f"複数の連続空行グループが見つかりました: {len(consecutive_groups)}箇所"
        
        # ヘッダー様の行の検出
        header_like_rows = []
        for idx, row in df.iterrows():
            non_na_values = row.dropna().astype(str).str.strip()
            if len(non_na_values) > 0:
                # 数値以外が多い行をヘッダー候補とする
                numeric_count = sum(1 for val in non_na_values if val.replace('.', '').replace('-', '').isdigit())
                if numeric_count / len(non_na_values) < 0.5:
                    header_like_rows.append(idx)
        
        # 複数のヘッダー様行が離れて存在する場合
        if len(header_like_rows) >= 2:
            gaps = [header_like_rows[i+1] - header_like_rows[i] for i in range(len(header_like_rows)-1)]
            if any(gap > 3 for gap in gaps):  # 3行以上離れたヘッダーがある
                return True, f"離れた位置に複数のヘッダー様行が検出されました: {header_like_rows}"
        
        return False, "単一テーブルと判定"
        
    except Exception as e:
        logger.error(f"複数テーブル検出でエラー: {e}")
        return False, f"検出処理でエラーが発生: {str(e)}"
