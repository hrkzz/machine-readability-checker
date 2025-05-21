import pandas as pd
import openpyxl
from pathlib import Path
from typing import Tuple, Dict, Any, List

# ファイル拡張子の設定
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"} 

def load_file(file_path: Path) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    ファイルを読み込み、DataFrameとメタデータを返す
    """
    if not file_path.suffix.lower() in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file format. Allowed formats: {ALLOWED_EXTENSIONS}")
    
    metadata = {}
    
    if file_path.suffix.lower() in {".xlsx", ".xls"}:
        # Excelファイルの場合
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        
        # メタデータの収集
        metadata["worksheet"] = ws
        metadata["merged_cells"] = list(ws.merged_cells.ranges)
        metadata["hidden_rows"], metadata["hidden_cols"] = get_hidden_rows_columns(ws)
        
        # DataFrameの読み込み
        df = pd.read_excel(file_path)
        
    else:  # CSVファイルの場合
        df = pd.read_csv(file_path)
        metadata["worksheet"] = None
        metadata["merged_cells"] = []
        metadata["hidden_rows"] = []
        metadata["hidden_cols"] = []
    
    return df, metadata

def get_hidden_rows_columns(worksheet) -> Tuple[List[int], List[str]]:
    """
    非表示の行と列を特定
    """
    hidden_rows = [idx + 1 for idx, row in enumerate(worksheet.row_dimensions) 
                  if worksheet.row_dimensions[row].hidden]
    hidden_cols = [chr(65 + idx) for idx, col in enumerate(worksheet.column_dimensions) 
                  if worksheet.column_dimensions[col].hidden]
    return hidden_rows, hidden_cols 