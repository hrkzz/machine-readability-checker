from pathlib import Path
from typing import Dict, Any, cast
import pandas as pd
import openpyxl
from src.config import PREVIEW_ROW_COUNT
import xlrd
from loguru import logger

# ログファイルの設定
logger.add("logs/file_loader.log", rotation="10 MB", retention="30 days", level="DEBUG")

# 対応可能な拡張子
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}

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
        logger.info(f"CSV読み込み完了: shape={df.shape}")
        
        result["sheets"].append({
            "sheet_name": "CSV",
            "dataframe": df,
            "preview_top": df.head(PREVIEW_ROW_COUNT),
            "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
        })
    elif suffix == ".xls":
        # .xls形式はxlrd 1.2.0で処理する（pandasを使わない）
        logger.info(f"=== .xls ファイル読み込み開始: {file_path} ===")
        try:
            wb = xlrd.open_workbook(str(file_path), formatting_info=True)
            logger.info(f"xlrdで開いたワークブック: {wb.nsheets} シート")

            for sheet in wb.sheets():
                logger.info(f"=== シート '{sheet.name}' を処理中 ===")
                rows = []
                for row_idx in range(sheet.nrows):
                    row = sheet.row_values(row_idx)
                    rows.append(row)

                df = pd.DataFrame(rows)
                logger.info(f"xlrdで構築したDataFrame: shape={df.shape}")

                result["sheets"].append({
                    "sheet_name": sheet.name,
                    "dataframe": df,
                    "preview_top": df.head(PREVIEW_ROW_COUNT),
                    "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
                })

        except Exception as e:
            logger.error(f".xls読み込みでエラー: {e}")
            logger.exception("詳細なエラー情報:")
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
                logger.info(f"xlsx シート '{ws.title}' 読み込み完了: shape={df.shape}")
            except Exception as e:
                logger.error(f"xlsx シート '{ws.title}' の読み込みでエラー: {e}")
                # 読み込みに失敗した場合は空 DataFrame
                df = pd.DataFrame()

            result["sheets"].append({
                "sheet_name": ws.title,
                "dataframe": df,
                "preview_top": df.head(PREVIEW_ROW_COUNT),
                "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
            })

    return result
