from pathlib import Path
from typing import Dict, Any, cast
import pandas as pd
import openpyxl
import asyncio
from pyppeteer import launch

from src.config import PREVIEW_ROW_COUNT

# 対応可能な拡張子（旧形式 .xls は除外）
ALLOWED_EXTENSIONS = {".xlsx", ".csv"}

# 保存先ディレクトリ（画像）
IMAGE_DIR = Path("images")
IMAGE_DIR.mkdir(exist_ok=True)


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    def is_empty_row(row: pd.Series) -> bool:
        return all(str(cell).strip() == "" or str(cell).lower() == "nan" for cell in row)

    mask = df.apply(is_empty_row, axis=1)
    result = df[~mask].reset_index(drop=True)
    return cast(pd.DataFrame, result)


def dataframe_to_html(df: pd.DataFrame) -> str:
    table_html = df.to_html(index=False, border=1, escape=False)
    html_template = f"""
    <html>
        <head>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    padding: 20px;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    font-size: 14px;
                }}
                th, td {{
                    border: 1px solid #333;
                    padding: 6px;
                    text-align: center;
                }}
            </style>
        </head>
        <body>{table_html}</body>
    </html>
    """
    return html_template


async def html_to_image(html_str: str, output_path: Path):
    browser = await launch(headless=True, args=['--no-sandbox'])
    page = await browser.newPage()
    await page.setContent(html_str)
    await page.screenshot({'path': str(output_path), 'fullPage': True})
    await browser.close()


def save_dataframe_image(df: pd.DataFrame, file_stem: str, sheet_name: str) -> Path:
    image_path = IMAGE_DIR / f"{file_stem}_{sheet_name}.png"
    html = dataframe_to_html(df.head(30))  # 表が長すぎると見切れるので先頭30行程度
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(html_to_image(html, image_path))
    return image_path


def load_file(file_path: Path) -> Dict[str, Any]:
    """
    ファイルを読み込み、各シートのデータ（先頭・末尾{PREVIEW_ROW_COUNT}行）と画像パスを返す。
    loader は「生のテーブル情報」を提供。構造解析などは table_parser 側で実施。
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
        # CSV は単一シートとして扱う（ヘッダーなし）
        df = pd.read_csv(file_path, header=None)
        df = drop_empty_rows(df)
        image_path = save_dataframe_image(df, file_path.stem, "CSV")

        result["sheets"].append({
            "sheet_name": "CSV",
            "dataframe": df,
            "preview_top": df.head(PREVIEW_ROW_COUNT),
            "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
            "image_path": image_path
        })

    else:
        # Excel (.xlsx) の場合は全シートを対象
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for ws in wb.worksheets:
            try:
                df = pd.read_excel(file_path, sheet_name=ws.title, header=None)
                df = drop_empty_rows(df)
            except Exception:
                df = pd.DataFrame()

            image_path = save_dataframe_image(df, file_path.stem, ws.title)

            result["sheets"].append({
                "sheet_name": ws.title,
                "dataframe": df,
                "preview_top": df.head(PREVIEW_ROW_COUNT),
                "preview_bottom": df.tail(PREVIEW_ROW_COUNT),
                "image_path": image_path
            })

    return result
