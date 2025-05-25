from pathlib import Path
import json

from src.processor.loader import load_file
from src.processor.table_parser import analyze_table_structure, select_main_sheet

# データファイルを指定（data/ に置いた Excel または CSV ファイル名を指定）
file_path = Path("data/b2020_gsm1j.xlsx")  # ← ← ← 適宜ファイル名を変更すること

# ファイル読み込み（各シートの DataFrame＋プレビュー取得）
file_result = load_file(file_path)

# メインシート選定
main_sheet = select_main_sheet(file_result["sheets"])

# 構造解析を実行
struct_info = analyze_table_structure(main_sheet)

# 結果表示
print("=== メインシート名 ===")
print(struct_info["sheet_name"])

print("=== メインシートの末尾10行（df.tail(10)）===")
print(main_sheet["preview_bottom"])

print("\n=== 構造解析レスポンス（LLMのJSON出力）===")
print(json.dumps(struct_info["structure_response"], indent=2, ensure_ascii=False))

