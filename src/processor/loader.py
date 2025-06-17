from pathlib import Path
from typing import Dict, Any
import pandas as pd
import xlrd
import openpyxl
from loguru import logger

from src.config import PREVIEW_ROW_COUNT

def detect_encoding(file_path: Path) -> str:
    """ファイルのエンコーディングを自動検出"""
    try:
        import chardet
        with open(file_path, 'rb') as f:
            raw_data = f.read()
        result = chardet.detect(raw_data)
        detected_encoding = result['encoding']
        confidence = result['confidence']
        logger.info(f"エンコーディング検出: {detected_encoding} (信頼度: {confidence})")
        return detected_encoding if confidence > 0.7 else None
    except ImportError:
        logger.warning("chardetライブラリが利用できません")
        return None
    except Exception as e:
        logger.error(f"エンコーディング検出でエラー: {e}")
        return None


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
    df = None
    
    # Streamlit Cloud環境での特別な処理
    logger.info(f"CSV読み込み開始: {file_path}")
    logger.info(f"ファイル存在確認: {file_path.exists()}")
    if file_path.exists():
        logger.info(f"ファイルサイズ: {file_path.stat().st_size} bytes")
    
    # 自動検出を最初に試行
    detected_encoding = detect_encoding(file_path)
    encodings = ["utf-8", "cp932", "shift_jis", "iso-2022-jp", "euc-jp", "latin1"]
    
    # 検出されたエンコーディングを最初に試行
    if detected_encoding and detected_encoding not in encodings:
        encodings.insert(0, detected_encoding)
    elif detected_encoding and detected_encoding in encodings:
        encodings.remove(detected_encoding)
        encodings.insert(0, detected_encoding)
    
    # Streamlit Cloud環境での追加オプション
    pandas_options = {
        'header': None,
        'on_bad_lines': 'skip',  # 問題のある行をスキップ
        'engine': 'python',     # Pythonエンジンを使用（より寛容）
    }
    
    for encoding in encodings:
        try:
            logger.info(f"エンコーディング {encoding} で読み込み試行中...")
            df = pd.read_csv(file_path, encoding=encoding, **pandas_options)
            logger.info(f"CSV読み込み完了（{encoding}）: shape={df.shape}")
            break
        except UnicodeDecodeError as e:
            logger.debug(f"エンコーディング {encoding} で読み込み失敗: {e}")
            continue
        except Exception as e:
            logger.error(f"CSV読み込みでエラー（{encoding}）: {e}")
            continue
    
    if df is None:
        logger.error("全てのエンコーディングで読み込みに失敗")
        # 最後の手段：バイナリ読み込みでデバッグ情報を出力
        try:
            with open(file_path, 'rb') as f:
                first_bytes = f.read(100)
            logger.error(f"ファイルの最初の100バイト: {first_bytes}")
        except Exception as debug_e:
            logger.error(f"デバッグ情報取得失敗: {debug_e}")
        
        raise ValueError(f"CSVファイル {file_path} のエンコーディングを特定できませんでした。対応エンコーディング: {encodings}")
    
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