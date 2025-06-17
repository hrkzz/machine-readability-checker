from pathlib import Path
from typing import Tuple, Optional, cast

from src.checker.common import get_excel_column_letter, detect_platform_characters, MAX_EXAMPLES, is_sheet_likely
from src.checker.handler.csv_handler import detect_multiple_tables_csv
from src.checker.handler.xls_handler import check_xls_merged_cells, check_xls_cell_formats, check_xls_hidden_rows_columns
from src.checker.handler.xlsx_handler import has_any_drawing_xlsx, check_xlsx_format_semantics
from src.processor.context import TableContext
from src.llm.llm_client import call_llm


class FormatHandler:
    """
    ファイル形式固有の処理を統合するハンドラークラス
    """
    
    @staticmethod
    def check_file_format(filepath: str) -> Tuple[bool, str]:
        """ファイル形式の妥当性チェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return True, "CSVファイル形式です"
        elif ext == ".xls":
            return True, "旧Excel（.xls）形式のため、一部の自動チェック（書式・図形など）が制限されます。必要に応じて目視での確認を行ってください"
        elif ext == ".xlsx":
            return True, "ファイル形式はExcel（.xlsx）です"
        else:
            return False, f"サポート外のファイル形式です: {ext}"
    
    @staticmethod
    def check_images_objects(filepath: str) -> Tuple[bool, str]:
        """画像・オブジェクトの存在チェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return True, "CSVファイルのため、画像やオブジェクトに関する問題はありません"
        elif ext == ".xls":
            return False, "xlsファイルでは図形や画像の自動判定ができません。必要に応じて目視でご確認ください"
        elif ext == ".xlsx":
            if has_any_drawing_xlsx(Path(filepath)):
                return False, "図形・テキストボックスが検出されました"
            return True, "図形・テキストボックスは見つかりませんでした"
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_multiple_tables(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """複数テーブルのチェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext in [".csv", ".xls"]:
            # CSV/XLSはDataFrameベースでチェック
            is_multiple, details = detect_multiple_tables_csv(ctx.data, ctx.sheet_name)
            if is_multiple:
                return False, f"複数テーブルの疑いがあります: {details}"
            return True, "1つのテーブルのみです"
        
        elif ext == ".xlsx":
            # XLSXはワークブック直接チェック
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            ws = workbook[ctx.sheet_name]
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "シート範囲情報が不足しているためチェックできません"

            start = min(column_rows) if isinstance(column_rows, list) else cast(int, column_rows)
            end = cast(int, data_end)

            flags = [
                any(cell not in (None, "") for cell in row)
                for row in ws.iter_rows(min_row=start + 1, max_row=end + 1, values_only=True)
            ]

            in_block = False
            blocks = 0
            for has_data in flags:
                if has_data and not in_block:
                    blocks += 1
                    in_block = True
                elif not has_data:
                    in_block = False

            if blocks > 1:
                return False, f"複数テーブルの疑いがあります（検出ブロック数: {blocks}）"
            return True, "1つのテーブルのみです"
        
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_hidden_rows_columns(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """非表示行・列のチェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return True, "CSVファイルのため、非表示行・列に関する問題はありません"
        
        elif ext == ".xls":
            hidden_rows, hidden_cols = check_xls_hidden_rows_columns(Path(filepath))

            row_str = (
                ", ".join(f"{sheet}シートの{r+1}行" for sheet, r in hidden_rows)
                if hidden_rows else "該当なし"
            )
            col_str = (
                ", ".join(f"{sheet}シートの{get_excel_column_letter(c+1)}列" for sheet, c in hidden_cols)
                if hidden_cols else "該当なし"
            )

            if hidden_rows or hidden_cols:
                return False, f"非表示行／列があります（行: {row_str}, 列: {col_str}）"
            return True, "非表示行／列はありません"
        
        elif ext == ".xlsx":
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            ws = workbook[ctx.sheet_name]
            hidden_rows = [d.index for d in ws.row_dimensions.values() if d.hidden]
            hidden_cols = [d.index for d in ws.column_dimensions.values() if d.hidden]

            row_str = (
                ", ".join(f"{r}行" for r in hidden_rows) if hidden_rows else "該当なし"
            )
            col_str = (
                ", ".join(f"{get_excel_column_letter(c)}列" for c in hidden_cols) if hidden_cols else "該当なし"
            )

            if hidden_rows or hidden_cols:
                return False, f"非表示行／列があります（行: {row_str}, 列: {col_str}）"
            return True, "非表示行／列はありません"
        
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_merged_cells(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """結合セルのチェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return True, "CSVファイルのため、結合セルに関する問題はありません"
        
        elif ext == ".xls":
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")
            if column_rows is None or data_end is None:
                return False, "結合セルチェックに必要な情報が不足しています"

            start = min(column_rows) if isinstance(column_rows, list) else cast(int, column_rows)
            end = cast(int, data_end)

            merged = check_xls_merged_cells(Path(filepath), ctx.sheet_name, start, end)
            if merged:
                return False, f"結合セルが検出されました: {merged}"
            else:
                return True, "結合セルはありません"
        
        elif ext == ".xlsx":
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            ws = workbook[ctx.sheet_name]
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "結合セルチェックに必要な情報が不足しています"

            start = min(column_rows) + 1 if isinstance(column_rows, list) else cast(int, column_rows) + 1
            end = cast(int, data_end) + 1

            relevant_merges = [
                str(rng)
                for rng in ws.merged_cells.ranges
                if rng.min_row >= start and rng.max_row <= end
            ]

            if relevant_merges:
                return False, f"結合セルが検出されました: {relevant_merges}"
            return True, "結合セルはありません"
        
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_format_semantics(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """書式による意味付けのチェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return True, "CSVファイルのため、書式による意味付けに関する問題はありません"
        
        elif ext == ".xls":
            data_start = ctx.row_indices.get("data_start", 0)
            data_end = ctx.row_indices.get("data_end", len(ctx.data) - 1)
            
            flagged = check_xls_cell_formats(Path(filepath), ctx.sheet_name, data_start, data_end)
            
            if flagged:
                return False, f"視覚的装飾による意味付けが検出されました（例: {flagged[:MAX_EXAMPLES]}）"
            return True, "書式ベースの意味づけは検出されませんでした"
        
        elif ext == ".xlsx":
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            ws = workbook[ctx.sheet_name]
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "書式チェックに必要な情報が不足しています"

            flagged = check_xlsx_format_semantics(ws, column_rows, data_end)

            if flagged:
                return False, f"視覚的装飾による意味付けが検出されました（例: {flagged[:MAX_EXAMPLES]}）"
            return True, "書式ベースの意味づけは検出されませんでした"
        
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_whitespace_formatting(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """空白による体裁調整のチェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            format_label = "CSV"
        elif ext == ".xls":
            format_label = "Excel(.xls)"
        elif ext == ".xlsx":
            format_label = "Excel(.xlsx)"
            # XLSXの場合はワークブック情報をチェック
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
        else:
            return False, "サポート外のファイル形式です"
        
        # 共通のロジック：DataFrame または worksheet から空白文字を検索
        sample_cells = []
        
        if ext == ".xlsx":
            # XLSX専用のworksheet読み取り
            ws = workbook[ctx.sheet_name]
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "空白チェックに必要な情報が不足しています"

            start = min(column_rows) + 1 if isinstance(column_rows, list) else cast(int, column_rows) + 1
            end = cast(int, data_end) + 1

            for r_idx, row in enumerate(ws.iter_rows(min_row=start, max_row=end, values_only=True), start=start):
                for c_idx, val in enumerate(row, start=1):
                    if isinstance(val, str) and "　" in val:
                        col_letter = get_excel_column_letter(c_idx)
                        cell_ref = f"{col_letter}{r_idx}"
                        sample_cells.append(f"{cell_ref}: '{val.strip()}'")
                        if len(sample_cells) >= 10:
                            break
                if len(sample_cells) >= 10:
                    break
        else:
            # CSV/XLS共通のDataFrame読み取り
            for row_idx, row in ctx.data.iterrows():
                for col_idx, val in enumerate(row):
                    if isinstance(val, str) and "　" in val:
                        col_letter = get_excel_column_letter(col_idx + 1)
                        cell_ref = f"{col_letter}{row_idx + 1}"
                        sample_cells.append(f"{cell_ref}: '{val.strip()}'")
                        if len(sample_cells) >= 10:
                            break
                if len(sample_cells) >= 10:
                    break

        if not sample_cells:
            return True, "体裁調整目的の空白は見つかりませんでした"

        prompt = f"""
            以下は{format_label}のセル値の一部です。これらの中に、見た目を整える目的（位置揃え・スペース調整など）で
            **空白（特に全角スペース）が使われているものがあるか**を判定してください。

            データ:
            {chr(10).join(sample_cells)}

            判断結果を次のいずれか一語で返してください：
            - 調整目的あり
            - 調整目的なし
        """

        result = call_llm(prompt)
        if "調整目的あり" in result:
            return False, f"体裁調整目的の空白が含まれている可能性があります（例: {sample_cells[:MAX_EXAMPLES]}）"
        return True, "体裁調整目的の空白は見つかりませんでした"
    
    @staticmethod
    def check_platform_dependent_characters(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """機種依存文字のチェック"""
        ext = Path(filepath).suffix.lower()
        issues = []
        
        if ext in [".csv", ".xls"]:
            # CSV/XLS共通：DataFrame使用の簡易版
            for row_idx, row in ctx.data.iterrows():
                for col_idx, val in enumerate(row):
                    if isinstance(val, str) and detect_platform_characters(val):
                        coord = f"{get_excel_column_letter(col_idx + 1)}{row_idx + 1}"
                        issues.append(f"{coord}: '{val}'")
        
        elif ext == ".xlsx":
            # XLSX専用：worksheet直接読み取り
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            ws = workbook[ctx.sheet_name]
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "機種依存文字チェックに必要な情報が不足しています"

            if isinstance(column_rows, list):
                start = min(column_rows) + 1
            else:
                start = cast(int, column_rows) + 1
            end = cast(int, data_end) + 1

            for r, row in enumerate(ws.iter_rows(min_row=start, max_row=end, values_only=True), start=start):
                for c, val in enumerate(row, start=1):
                    if isinstance(val, str) and detect_platform_characters(val):
                        coord = f"{get_excel_column_letter(c)}{r}"
                        issues.append(f"{coord}: '{val}'")
        
        else:
            return False, "サポート外のファイル形式です"

        if issues:
            return False, f"機種依存文字が含まれています（例: {issues[:MAX_EXAMPLES]}）"
        return True, "機種依存文字は含まれていません"
    
    # Level3専用の形式固有処理
    
    @staticmethod
    def check_codebook_exists(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """コード表の存在チェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return False, "CSVファイルのためコード表チェックをスキップします。コード表は別途提供してください"
        
        elif ext == ".xls":
            if filepath and filepath.lower().endswith('.xls'):
                try:
                    import xlrd
                    xl_workbook = xlrd.open_workbook(filepath)
                    for sheet_name in xl_workbook.sheet_names():
                        if sheet_name == ctx.sheet_name:
                            continue
                        
                        # シート名からコード表らしさを判定
                        if any(keyword in sheet_name.lower() for keyword in ['code', 'コード', 'master', 'マスタ']):
                            return True, f"コード表とみられるシート: {sheet_name}"
                            
                    return False, "コード表が見つかりません（.xlsファイルでは詳細検索は制限されます）"
                except Exception as e:
                    return False, f".xlsファイルのシート検索でエラー: {e}"
            else:
                return False, ".xlsファイルでないためコード表チェックをスキップ"
        
        elif ext == ".xlsx":
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            for sheet in workbook.worksheets:
                if sheet.title == ctx.sheet_name:
                    continue
                if is_sheet_likely(sheet, "コード表"):
                    return True, f"コード表とみられるシート: {sheet.title}"
            return False, "コード表が見つかりません"
        
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_question_master_exists(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """設問マスターの存在チェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return False, "CSVファイルのため設問マスターチェックをスキップします。設問マスター（変数定義表）は別途提供してください"
        
        elif ext == ".xls":
            if filepath and filepath.lower().endswith('.xls'):
                try:
                    import xlrd
                    xl_workbook = xlrd.open_workbook(filepath)
                    for sheet_name in xl_workbook.sheet_names():
                        if sheet_name == ctx.sheet_name:
                            continue
                        
                        # シート名から設問マスターらしさを判定
                        if any(keyword in sheet_name.lower() for keyword in ['question', '設問', 'master', 'マスタ', 'variable', '変数']):
                            return True, f"設問マスターとみられるシート: {sheet_name}"
                            
                    return False, "設問マスター（変数定義表）が見つかりません（.xlsファイルでは詳細検索は制限されます）"
                except Exception as e:
                    return False, f".xlsファイルのシート検索でエラー: {e}"
            else:
                return False, ".xlsファイルでないため設問マスターチェックをスキップ"
        
        elif ext == ".xlsx":
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            for sheet in workbook.worksheets:
                if sheet.title == ctx.sheet_name:
                    continue
                if is_sheet_likely(sheet, "設問マスター"):
                    return True, f"設問マスターとみられるシート: {sheet.title}"
            return False, "設問マスター（変数定義表）が見つかりません"
        
        else:
            return False, "サポート外のファイル形式です"
    
    @staticmethod
    def check_metadata_presence(ctx: TableContext, workbook: Optional[object], filepath: str) -> Tuple[bool, str]:
        """メタデータの存在チェック"""
        ext = Path(filepath).suffix.lower()
        
        if ext == ".csv":
            return False, "CSVファイルのためメタデータチェックをスキップします。調査概要やメタデータは別途提供してください"
        
        elif ext == ".xls":
            if filepath and filepath.lower().endswith('.xls'):
                try:
                    import xlrd
                    xl_workbook = xlrd.open_workbook(filepath)
                    for sheet_name in xl_workbook.sheet_names():
                        if sheet_name == ctx.sheet_name:
                            continue
                        
                        # シート名からメタ情報らしさを判定
                        if any(keyword in sheet_name.lower() for keyword in ['meta', 'メタ', 'info', '情報', '概要', 'readme']):
                            return True, f"メタ情報とみられるシート: {sheet_name}"
                            
                    return False, "調査概要やメタデータが確認できません（.xlsファイルでは詳細検索は制限されます）"
                except Exception as e:
                    return False, f".xlsファイルのシート検索でエラー: {e}"
            else:
                return False, ".xlsファイルでないためメタデータチェックをスキップ"
        
        elif ext == ".xlsx":
            if workbook is None:
                return False, "ワークブック情報が不足しているためチェックできません"
            
            for sheet in workbook.worksheets:
                if sheet.title == ctx.sheet_name:
                    continue
                if is_sheet_likely(sheet, "メタ情報"):
                    return True, f"メタ情報とみられるシート: {sheet.title}"
            return False, "調査概要やメタデータが確認できません"
        
        else:
            return False, "サポート外のファイル形式です"