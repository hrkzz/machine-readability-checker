from pathlib import Path
from typing import Dict, Any
import pandas as pd
import xlrd
import openpyxl
from loguru import logger

from src.config import PREVIEW_ROW_COUNT


def load_file(file_path: Path) -> Dict[str, Any]:
    """
    ファイルを読み込み、各シートのデータ（先頭・末尾{PREVIEW_ROW_COUNT}行）を返す。
    loader はあくまで「生のテーブル情報」を提供し、オブジェクト検出などは
    checker 側のユーティリティに委ねます。
    """
    extension = file_path.suffix.lower()
    
    if extension == ".csv":
        return _load_csv(file_path)
    elif extension == ".xls":
        return _load_xls(file_path)
    elif extension == ".xlsx":
        return _load_xlsx(file_path)
    else:
        raise ValueError(f"サポートされていないファイル形式: {extension}")


def _load_csv(file_path: Path) -> Dict[str, Any]:
    """CSV形式の読み込み処理（エンコーディング対応）"""
    try:
        try:
            df = pd.read_csv(file_path, header=None, encoding="utf-8")
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, header=None, encoding="cp932")
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, header=None, encoding="shift_jis")
        logger.info(f"CSV読み込み完了: shape={df.shape}")
    except UnicodeDecodeError as e:
        logger.error(f"CSVのエンコーディングエラー: {e}")
        raise ValueError("CSVファイルのエンコーディングに問題があります。UTF-8, Shift_JIS, CP932のいずれでも読み込めませんでした。")
    
    sheet_info = {
        "sheet_name": "CSV",
        "dataframe": df,
        "preview_top": df.head(PREVIEW_ROW_COUNT),
        "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
    }
    
    return {
        "file_path": file_path,
        "file_type": ".csv",
        "sheets": [sheet_info]
    }


def _load_xls(file_path: Path) -> Dict[str, Any]:
    """XLS形式の読み込み処理"""
    logger.info(f"=== .xls ファイル読み込み開始: {file_path} ===")
    
    sheets = []
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

            sheet_info = {
                "sheet_name": sheet.name,
                "dataframe": df,
                "preview_top": df.head(PREVIEW_ROW_COUNT),
                "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
            }
            sheets.append(sheet_info)

    except Exception as e:
        logger.error(f".xls読み込みでエラー: {e}")
        logger.exception("詳細なエラー情報:")
        # エラー時は空のデータフレームを返す
        sheet_info = {
            "sheet_name": "Sheet1",
            "dataframe": pd.DataFrame(),
            "preview_top": pd.DataFrame(),
            "preview_bottom": pd.DataFrame(),
        }
        sheets.append(sheet_info)
    
    return {
        "file_path": file_path,
        "file_type": ".xls",
        "sheets": sheets
    }


def _load_xlsx(file_path: Path) -> Dict[str, Any]:
    """XLSX形式の読み込み処理"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheets = []
    
    for ws in wb.worksheets:
        try:
            # ヘッダーは後で LLM が判定するため、ここでは header=None
            df = pd.read_excel(file_path, sheet_name=ws.title, header=None)
            logger.info(f"xlsx シート '{ws.title}' 読み込み完了: shape={df.shape}")
        except Exception as e:
            logger.error(f"xlsx シート '{ws.title}' の読み込みでエラー: {e}")
            # 読み込みに失敗した場合は空 DataFrame
            df = pd.DataFrame()

        sheet_info = {
            "sheet_name": ws.title,
            "dataframe": df,
            "preview_top": df.head(PREVIEW_ROW_COUNT),
            "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
        }
        sheets.append(sheet_info)

    return {
        "file_path": file_path,
        "file_type": ".xlsx",
        "sheets": sheets
    }


# 対応可能な拡張子
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"} 