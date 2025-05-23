import pandas as pd
import openpyxl
from pathlib import Path
from typing import Tuple, Dict, Any, List, Optional
from openpyxl.worksheet.worksheet import Worksheet

# サポートされる拡張子
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"} 

def load_file(file_path: Path) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    ファイルを読み込み、DataFrameとメタデータを返す
    """
    if file_path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file format. Allowed formats: {ALLOWED_EXTENSIONS}")
    
    metadata: Dict[str, Any] = {}

    if file_path.suffix.lower() in {".xlsx", ".xls"}:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws: Optional[Worksheet] = wb.active if wb.worksheets else None
        
        if ws is not None:
            metadata["worksheet"] = ws
            metadata["merged_cells"] = list(ws.merged_cells.ranges)
            metadata["hidden_rows"], metadata["hidden_cols"] = get_hidden_rows_columns(ws)
        else:
            metadata["worksheet"] = None
            metadata["merged_cells"] = []
            metadata["hidden_rows"] = []
            metadata["hidden_cols"] = []

        df = pd.read_excel(file_path)

    else:  # CSVファイルの場合
        df = pd.read_csv(file_path)
        metadata["worksheet"] = None
        metadata["merged_cells"] = []
        metadata["hidden_rows"] = []
        metadata["hidden_cols"] = []

    return df, metadata

def get_hidden_rows_columns(worksheet: Worksheet) -> Tuple[List[int], List[str]]:
    """
    非表示の行と列を特定
    """
    hidden_rows = [
        idx for idx, dim in worksheet.row_dimensions.items()
        if dim.hidden
    ]
    hidden_cols = [
        key for key, dim in worksheet.column_dimensions.items()
        if dim.hidden
    ]
    return hidden_rows, hidden_cols
