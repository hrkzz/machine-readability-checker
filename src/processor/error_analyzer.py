# src/processor/error_analyzer.py
from src.llm.llm_client import call_llm

def analyze_parsing_error(
    sheet_name: str,
    error_message: str,
    header_preview: str
) -> str:
    prompt = f"""
シート「{sheet_name}」のテーブル構造解析中に以下のエラーが発生しました：
{error_message}

ヘッダープレビュー（先頭25行）：
{header_preview}

このエラーを解決するために、
1. プロンプトのどこをどう改善すればよいか  
2. 前処理コード（loader や extract_structured_table）のどこをどう変えればよいか  

を日本語で具体的にアドバイスしてください。
"""
    return call_llm(prompt)