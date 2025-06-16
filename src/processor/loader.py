from pathlib import Path
from typing import Dict, Any, cast
import pandas as pd
import openpyxl
from src.config import PREVIEW_ROW_COUNT

# xlrdライブラリの読み込み（.xls対応のため）
try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False
    print("Warning: xlrd library not available. .xls files may not work properly.")

# 対応可能な拡張子（.xls も対応可能に）
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}

def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    def is_empty_row(row: pd.Series) -> bool:
        return all(str(cell).strip() == "" or str(cell).lower() == "nan" for cell in row)

    mask = df.apply(is_empty_row, axis=1)
    empty_count = mask.sum()
    
    if empty_count > 0:
        print(f"空行として判定された行数: {empty_count}")
        # 空行の例を表示（最初の3行まで）
        empty_indices = df.index[mask].tolist()[:3]
        for idx in empty_indices:
            print(f"空行 {idx}: {df.iloc[idx].tolist()}")
    
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
    elif suffix == ".xls":
        # .xls形式の場合はxlrdエンジンを使用
        print(f"=== .xls ファイル読み込み開始: {file_path} ===")
        try:
            # xlrdライブラリを直接使用してファイル情報を確認
            wb = xlrd.open_workbook(str(file_path))
            print(f"xlrdで開いたワークブック: {wb.nsheets} シート")
            
            # まずはシート名を取得
            xl_file = pd.ExcelFile(file_path, engine='xlrd')
            print(f"pandasで取得したシート名: {xl_file.sheet_names}")
            
            for sheet_name in xl_file.sheet_names:
                print(f"=== シート '{sheet_name}' を処理中 ===")
                try:
                    # xlrdで直接シート情報を確認
                    sheet = wb.sheet_by_name(sheet_name)
                    print(f"xlrd情報: {sheet.nrows} 行, {sheet.ncols} 列")
                    
                    # ヘッダーは後で LLM が判定するため、ここでは header=None
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
                    print(f"pandas読み込み完了: shape={df.shape}")
                    
                    if not df.empty:
                        print("データサンプル (先頭3行):")
                        print(df.head(3))
                        print("\nデータサンプル (データ型):")
                        print(df.dtypes.head(10))
                    
                    # 空行除去前後の比較
                    original_rows = len(df)
                    df = drop_empty_rows(df)
                    print(f"空行除去: {original_rows} -> {len(df)} 行")
                    
                except Exception as e:
                    print(f"シート '{sheet_name}' の読み込みでエラー: {e}")
                    import traceback
                    traceback.print_exc()
                    # 読み込みに失敗した場合は空 DataFrame
                    df = pd.DataFrame()

                result["sheets"].append({
                    "sheet_name": sheet_name,
                    "dataframe": df,
                    "preview_top": df.head(PREVIEW_ROW_COUNT),
                    "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
                })
                print(f"最終的なDataFrame shape: {df.shape}")
                
        except Exception as e:
            print(f"xlsファイル全体の読み込みでエラー: {e}")
            import traceback
            traceback.print_exc()
            # xlsファイル全体の読み込みに失敗した場合
            result["sheets"].append({
                "sheet_name": "Sheet1",
                "dataframe": pd.DataFrame(),
                "preview_top": pd.DataFrame(),
                "preview_bottom": pd.DataFrame(),
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
                "preview_top": df.head(PREVIEW_ROW_COUNT),
                "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
            })

    return result
