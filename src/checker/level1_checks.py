import os
import re
import pandas as pd
from openpyxl import load_workbook
from .utils import (
    detect_platform_characters,
    detect_notes_outside_table
)

def check_no_images_or_objects(df=None, workbook=None, filepath=None):
    for sheet in workbook.worksheets:
        if sheet._images or getattr(sheet, 'drawings', None):
            return False, f"シート '{sheet.title}' に画像またはオブジェクトが含まれています"
    return True, "画像やオブジェクトは含まれていません"

def check_no_format_based_semantics(df=None, workbook=None, filepath=None):
    styled_cells = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb != '00000000':
                    styled_cells.append(cell.coordinate)
    if styled_cells:
        return False, f"書式による強調セルがあります（例: {styled_cells[:3]})"
    return True, "書式ベースの意味づけは検出されませんでした"

def check_no_merged_cells(df=None, workbook=None, filepath=None):
    merged = []
    for sheet in workbook.worksheets:
        merged += [str(rng) for rng in sheet.merged_cells.ranges]
    if merged:
        return False, f"結合セルが検出されました: {merged}"
    return True, "結合セルはありません"

def check_valid_file_format(df=None, workbook=None, filepath=None):
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in ['.csv', '.xlsx']:
        return False, f"サポート外のファイル形式です: {ext}"
    return True, "ファイル形式はCSVまたはExcelです"

def check_single_data_per_cell(df=None, workbook=None, filepath=None):
    problem_cells = []
    pattern = re.compile(r"[\n,;/]")
    sheet = workbook.worksheets[0]
    for row in sheet.iter_rows(values_only=True):
        for idx, cell in enumerate(row):
            if isinstance(cell, str) and pattern.search(cell):
                problem_cells.append(cell)
    if problem_cells:
        return False, f"複数データセルが検出されました（例: {problem_cells[:3]})"
    return True, "各セルに1データのみです"

def check_no_whitespace_formatting(df=None, workbook=None, filepath=None):
    problem = []
    sheet = workbook.worksheets[0]
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str) and (cell.startswith(' ') or cell.endswith(' ') or '\n' in cell):
                problem.append(repr(cell))
    if problem:
        return False, f"余分な空白/改行が検出されました（例: {problem[:3]})"
    return True, "スペースや改行による整形はありません"

def check_one_table_per_sheet(df=None, workbook=None, filepath=None):
    issues = []
    for sheet in workbook.worksheets:
        empty_rows = sum(1 for row in sheet.iter_rows(values_only=True) if all(cell in (None, '') for cell in row))
        if empty_rows > 1:
            issues.append(sheet.title)
    if issues:
        return False, f"複数表が含まれている可能性のあるシート: {issues}"
    return True, "各シートに1つの表のみです"

def check_no_hidden_rows_or_columns(df=None, workbook=None, filepath=None):
    issues = []
    for sheet in workbook.worksheets:
        hidden_rows = [dim.index for dim in sheet.row_dimensions.values() if dim.hidden]
        hidden_cols = [dim.index for dim in sheet.column_dimensions.values() if dim.hidden]
        if hidden_rows or hidden_cols:
            issues.append({
                'sheet': sheet.title,
                'hidden_rows': hidden_rows,
                'hidden_cols': hidden_cols
            })
    if issues:
        return False, f"非表示行/列があります: {issues}"
    return True, "行や列の非表示はありません"

def check_no_notes_outside_table(df=None, workbook=None, filepath=None):
    texts = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
            for cell in row:
                if isinstance(cell, str) and cell.strip():
                    texts.append(cell.strip())
    if not texts:
        return True, "表外のメモや備考は検出されませんでした"
    result = detect_notes_outside_table(texts)
    if "注釈あり" in result:
        return False, "表外に備考/注釈が含まれている可能性があります"
    return True, "表外の注釈や備考は含まれていません"

def check_no_platform_dependent_characters(df=None, workbook=None, filepath=None):
    issues = []
    sheet = workbook.worksheets[0]
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str) and detect_platform_characters(cell):
                issues.append(cell)
    if issues:
        return False, f"機種依存文字が含まれています（例: {issues[:3]})"
    return True, "機種依存文字は含まれていません"
