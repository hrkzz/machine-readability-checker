from pathlib import Path
from typing import Dict, Any, List, cast
import pandas as pd
import openpyxl
from loguru import logger
import xlrd

# TableContext のインポートが必要 (src/processor/context.py から)
from .context import TableContext

# ログファイルの設定
logger.add("logs/file_loader.log", rotation="10 MB", retention="30 days", level="DEBUG")

# 対応可能な拡張子
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


# --- 構造解析・抽出関数 ---
def extract_structured_table(
    info: Dict[str, Any],
    header_start_row: int = 1,  # 1-based: ヘッダー開始行
    header_end_row: int = 1,  # 1-based: ヘッダー終了行
) -> TableContext:
    """
    ユーザー定義の構造（ヘッダー範囲）に基づいてTableContextを構築する。
    ヘッダー行数は header_end_row - header_start_row + 1 として計算される。
    """
    df = info["dataframe"]
    total_rows = len(df)

    # 1. ヘッダー行のインデックス (0-based) を計算
    header_start_idx = header_start_row - 1
    header_end_idx = header_end_row - 1

    # データ開始行のインデックス (0-based)
    data_start = header_end_idx + 1
    data_end = total_rows - 1  # 最終行

    # 2. 基本的な妥当性チェック
    if (
        total_rows == 0
        or header_start_idx < 0
        or header_end_idx < header_start_idx
        or header_end_idx >= total_rows
    ):
        logger.warning(
            f"データセットが小さすぎるか、ヘッダー指定が無効です。Total rows: {total_rows}"
        )
        # ヘッダー指定が無効な場合は、データ開始/終了も無効に
        invalid_data = pd.DataFrame()
        if total_rows > 0:
            # エラー時の表示用に最初の行をカラム名とする
            cols = df.iloc[0].tolist()
        else:
            cols = []

        return TableContext(
            sheet_name=info["sheet_name"],
            data=invalid_data,
            columns=cols,
            upper_annotations=pd.DataFrame(),
            lower_annotations=pd.DataFrame(),
            row_indices={
                "column_rows": [],
                "data_start": 0,
                "data_end": -1,
                "annotations": [],
            },
        )

    # 3. 構造情報の定義
    # column_rows: ヘッダーとして指定された行のインデックスリスト (0-based)
    column_rows = list(range(header_start_idx, header_end_idx + 1))
    header_row_count = len(column_rows)

    # 4. 注釈行の抽出
    # 上部注釈（ヘッダー行より上）
    upper = (
        df.iloc[:header_start_idx].copy() if header_start_idx > 0 else pd.DataFrame()
    )
    lower = pd.DataFrame()

    # 5. カラム名（ヘッダー）の構築
    col_df = df.iloc[column_rows].fillna("").astype(str)

    if header_row_count > 1:
        # マルチインデックスの構築ロジックを再利用
        arrays = col_df.values
        arrays_fixed = []
        for level in arrays:
            fixed_level = []
            last_val = ""
            for val in level:
                # 空白セルを上の行または左のセルの値で補完する簡易ロジック
                if val == "":
                    val = last_val or "(空白)"
                else:
                    last_val = val
                fixed_level.append(val)
            arrays_fixed.append(fixed_level)
        cols = pd.MultiIndex.from_arrays(arrays_fixed)
    else:
        # 単一行ヘッダー
        cols = col_df.iloc[0].tolist()

    # 6. 実データ部分の切り出し
    if data_start <= data_end and data_start < total_rows:
        data = df.iloc[data_start : data_end + 1].copy()
    else:
        data = pd.DataFrame()  # データなし

    if not data.empty:
        # カラム数がデータフレームの列数と合わない場合のエラー回避
        if len(cols) != data.shape[1]:
            logger.warning(
                f"ヘッダー列数({len(cols)})とデータ列数({data.shape[1]})が一致しません。"
            )
            cols = [f"Col{i + 1}" for i in range(data.shape[1])]  # 仮の列名を割り当て

        data.columns = cols
        data.reset_index(drop=True, inplace=True)

    row_indices: Dict[str, int] = cast(
        Dict[str, int],
        {
            "column_rows": column_rows,
            "data_start": data_start,
            "data_end": data_end,
            "annotations": [],
        },
    )

    return TableContext(
        sheet_name=info["sheet_name"],
        data=data,
        columns=cols,
        upper_annotations=upper,
        lower_annotations=lower,
        row_indices=row_indices,
    )


def load_file_and_extract_context(
    file_path: Path,
    sheet_name: str,  # 新規引数: シート名
    header_start_row: int = 1,  # 1-based: ヘッダー開始行
    header_end_row: int = 1,  # 1-based: ヘッダー終了行
) -> TableContext:
    """
    ファイルを読み込み、指定されたシートを選択し、ユーザー定義の設定でTableContextを抽出する。
    """
    suffix = file_path.suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file format: {suffix}")

    # 1. ファイル読み込み (全シート/CSVデータをロード)
    sheets_data: List[Dict[str, Any]] = []

    if suffix == ".csv":
        # CSV は単一の「シート」と見なし、シート名は無視して "CSV" で固定
        try:
            df = pd.read_csv(file_path, header=None, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, header=None, encoding="shift_jis")
        except Exception as e:
            logger.error(f"CSV読み込みエラー: {e}")
            raise ValueError(f"CSVファイルの読み込みに失敗しました: {e}")

        sheets_data.append(
            {
                "sheet_name": "CSV",
                "dataframe": df,
            }
        )
    elif suffix == ".xls":
        # .xls形式
        try:
            wb = xlrd.open_workbook(str(file_path), formatting_info=True)
            for sheet in wb.sheets():
                rows = []
                for row_idx in range(sheet.nrows):
                    rows.append(sheet.row_values(row_idx))
                df = pd.DataFrame(rows)
                sheets_data.append({"sheet_name": sheet.name, "dataframe": df})
        except Exception as e:
            logger.error(f".xls読み込みでエラー: {e}")
            raise ValueError(f"XLSファイルの読み込みに失敗しました: {e}")
    else:
        # .xlsx
        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None, header=None)
            for name, df in all_sheets.items():
                sheets_data.append({"sheet_name": name, "dataframe": df})
        except Exception as e:
            logger.error(f".xlsx読み込みでエラー: {e}")
            raise ValueError(f"XLSXファイルの読み込みに失敗しました: {e}")

    # 2. 指定シートの選択
    main_sheet = None
    if suffix == ".csv":
        main_sheet = sheets_data[0]
    else:
        for sheet in sheets_data:
            if sheet["sheet_name"] == sheet_name:
                main_sheet = sheet
                break

    if main_sheet is None:
        if suffix != ".csv":
            raise ValueError(f"指定されたシート名 '{sheet_name}' が見つかりません。")

    if main_sheet["dataframe"].empty:
        raise ValueError("選択されたシート/ファイルに有効なデータが含まれていません。")

    # 3. 構造解析とデータ抽出 (ユーザー定義の引数を渡す)
    try:
        ctx = extract_structured_table(
            main_sheet, header_start_row=header_start_row, header_end_row=header_end_row
        )
        return ctx
    except Exception as e:
        logger.error(f"TableContext 抽出エラー: {e}")
        raise ValueError(f"データ構造の抽出に失敗しました: {e}")


# load_file を load_file_and_extract_context のエイリアスとして残す
load_file = load_file_and_extract_context

# --------------------------------------------------------------------------------
# 補足: シート名リスト取得関数 (app.pyで必要)
# --------------------------------------------------------------------------------


def get_sheet_names(file_path: Path) -> List[str]:
    """Excelファイルからシート名のリストを取得する（CSVの場合は["CSV"]を返す）"""
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return ["CSV"]

    sheet_names = []
    try:
        if suffix == ".xlsx":
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
        elif suffix == ".xls":
            wb = xlrd.open_workbook(str(file_path))
            sheet_names = wb.sheet_names()
    except Exception as e:
        logger.error(f"シート名取得エラー: {e}")
        # エラー発生時はダミーのシート名を返す
        return ["Sheet1 (読み込みエラー)"]

    return sheet_names
