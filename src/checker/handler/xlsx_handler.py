import zipfile
from pathlib import Path
from openpyxl.worksheet.worksheet import Worksheet
from ...llm.llm_client import call_llm
from src.checker.common import get_excel_column_letter


def has_any_drawing_xlsx(path: Path) -> bool:
    """
    Excel ファイル（.xlsx）に図形やオブジェクトが含まれているかをチェック
    """
    if path.suffix.lower() != ".xlsx":
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


def is_sheet_likely_xlsx(sheet: Worksheet, category: str) -> bool:
    """
    XLSXシートが特定のカテゴリに該当するかを判定
    """
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


def check_xlsx_format_semantics(worksheet, column_rows, data_end):
    """
    XLSX用の書式チェック
    """
    flagged = []
    start = min(column_rows) + 1 if isinstance(column_rows, list) else column_rows + 1
    end = data_end + 1

    for row in worksheet.iter_rows(min_row=start, max_row=end):
        for cell in row:
            coord = cell.coordinate
            fill = cell.fill
            if fill and fill.fgColor:
                fg = fill.fgColor
                if hasattr(fg, "rgb") and isinstance(fg.rgb, str):
                    rgb = fg.rgb.upper()
                    if rgb not in ("00000000", "FFFFFFFF", "FF000000"):
                        flagged.append(f"{coord}（塗りつぶし）")

            font = cell.font
            if font:
                if font.color and hasattr(font.color, "rgb") and isinstance(font.color.rgb, str):
                    rgb = font.color.rgb.upper()
                    if rgb not in ("00000000", "FF000000"):
                        flagged.append(f"{coord}（文字色）")

                if font.bold:
                    flagged.append(f"{coord}（太字）")
                if font.italic:
                    flagged.append(f"{coord}（イタリック）")
                if font.underline:
                    flagged.append(f"{coord}（下線）")
                if font.sz and (font.sz < 9 or font.sz > 13):
                    flagged.append(f"{coord}（フォントサイズ {font.sz}）")

    return flagged 