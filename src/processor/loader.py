from pathlib import Path
from typing import Dict, Any, cast
import pandas as pd
import openpyxl
from src.config import PREVIEW_ROW_COUNT

# 対応可能な拡張子（旧形式 .xls は除外）
ALLOWED_EXTENSIONS = {".xlsx", ".csv"}

def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    def is_empty_row(row: pd.Series) -> bool:
        return all(str(cell).strip() == "" or str(cell).lower() == "nan" for cell in row)

    mask = df.apply(is_empty_row, axis=1)
    result = df[~mask].reset_index(drop=True)
    return cast(pd.DataFrame, result)

def load_file(file_path: Path) -> Dict[str, Any]:
    """
    ファイルを読み込み、各シートのデータ（先頭・末尾{PREVIEW_ROW_COUNT}行）を返す。
    loader はあくまで「生のテーブル情報」を提供し、オブジェクト検出などは
    checker 側のユーティリティに委ねます。
    """
    suffix = file_path.suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file format: {suffix}")

    result: Dict[str, Any] = {
        "file_path": file_path,
        "file_type": suffix,
        "sheets": []
    }

    if suffix == ".csv":
        # CSV は単一の「シート」と見なし、ヘッダーなしで読み込む
        df = pd.read_csv(file_path, header=None)
        df = drop_empty_rows(df)
        result["sheets"].append({
            "sheet_name": "CSV",
            "dataframe": df,
            "preview_top": df.head(PREVIEW_ROW_COUNT),
            "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
        })
    else:
        # .xlsx
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for ws in wb.worksheets:
            try:
                # ヘッダーは後で LLM が判定するため、ここでは header=None
                df = pd.read_excel(file_path, sheet_name=ws.title, header=None)
                df = drop_empty_rows(df)
            except Exception:
                # 読み込みに失敗した場合は空 DataFrame
                df = pd.DataFrame()

            result["sheets"].append({
                "sheet_name": ws.title,
                "dataframe": df,
                "preview_top": df.head(10),
                "preview_bottom": df.tail(10),
            })

    return result
