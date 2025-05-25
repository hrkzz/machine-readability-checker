from typing import Dict, Any, List, cast
import pandas as pd
import json
import re

from src.llm.llm_client import call_llm
from .context import TableContext  # TableContext を返すように

def select_main_sheet(sheet_infos: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    複数シートの先頭10行をLLMに渡し、メインテーブルと判断されたシートの情報を返す。
    """
    prompt_parts = []
    for sheet in sheet_infos:
        preview = sheet["preview_top"].fillna("").astype(str).values.tolist()
        text = "\n".join(",".join(row) for row in preview)
        prompt_parts.append(f"[{sheet['sheet_name']}]\n{text}")

    prompt = (
        "以下は複数のExcelシートの先頭10行です。\n"
        "業務上・分析上のメインテーブルに該当するシート名のみを、正確に1つだけ出力してください。\n\n"
        + "\n\n".join(prompt_parts)
    )

    response = call_llm(prompt)
    for sheet in sheet_infos:
        if sheet["sheet_name"] in response:
            return sheet

    # マッチしなければ最初のシートを返す
    return sheet_infos[0]


def analyze_table_structure(sheet: Dict[str, Any]) -> Dict[str, Any]:
    """
    メインシートの先頭・末尾10行をLLMに渡し、構造情報（カラム行、データ行、注釈行）を取得する。
    """
    df = sheet["dataframe"]
    top = sheet["preview_top"].fillna("").astype(str).values.tolist()
    bot = sheet["preview_bottom"].fillna("").astype(str).values.tolist()
    content = "\n".join(",".join(row) for row in (top + [["..."]] + bot))

    prompt = (
        f"以下はシート「{sheet['sheet_name']}」の先頭10行と末尾10行です。\n"
        "セル結合や複数行ヘッダーを含む可能性のある構造化データです。\n\n"
        "次の情報を必ず **JSON形式**（1-indexed）で返してください：\n"
        "- column_rows: カラム行番号のリスト（例: [3], [3,4]）\n"
        "- data_start: データ開始行番号（例: 5）\n"
        "- data_end: データ終了行番号（例: 29）\n"
        "- annotation_rows: 注釈行（省略可, 複数可）\n\n"
        "出力例：\n"
        "```json\n"
        "{\n"
        "  \"column_rows\": [3, 4],\n"
        "  \"data_start\": 5,\n"
        "  \"data_end\": 29,\n"
        "  \"annotation_rows\": [1, 30]\n"
        "}\n"
        "```\n\n"
        f"プレビュー:\n{content}"
    )

    response = call_llm(prompt)
    return {
        "sheet_name": sheet["sheet_name"],
        "dataframe": df,
        "structure_response": response
    }


def extract_structured_table(info: Dict[str, Any]) -> TableContext:
    """
    LLMのJSON応答をパースし、TableContextを構築して返す。
    マルチ行ヘッダーも pandas.MultiIndex で扱う。
    """
    raw = info["structure_response"]
    df = info["dataframe"]

    # ```json ... ``` で返ってきた場合は中身だけを取り出し
    m = re.search(r"```json\s*(\{.*?\})\s*```", raw, re.DOTALL)
    js = m.group(1) if m else raw.strip()

    try:
        parsed = json.loads(js)
        column_rows = [i - 1 for i in parsed["column_rows"]]
        data_start = parsed["data_start"] - 1
        data_end = parsed["data_end"] - 1
        annotations = [i - 1 for i in parsed.get("annotation_rows", [])]
    except Exception:
        print("=== JSONパース失敗 ===")
        print(raw)
        raise ValueError("構造出力のパースに失敗しました")

    # 上部注釈（ヘッダー行の上）
    upper = df.iloc[:min(column_rows)].copy()
    # 下部注釈（データ行の下）
    lower = df.iloc[data_end + 1:].copy()

    # カラム行の抽出（単一行 or 複数行 → MultiIndex）
    col_df = df.iloc[column_rows].fillna("").astype(str)
    if len(column_rows) > 1:
        cols = pd.MultiIndex.from_arrays(col_df.values)
    else:
        cols = col_df.iloc[0].tolist()

    # 実データ部分
    data = df.iloc[data_start : data_end + 1].copy()
    data.columns = cols
    data.reset_index(drop=True, inplace=True)

    # row_indices の型を強制キャストして pyright エラーを抑制
    row_indices: Dict[str, int] = cast(Dict[str, int], {
        "column_rows": column_rows,
        "data_start": data_start,
        "data_end": data_end,
        "annotations": annotations
    })

    return TableContext(
        sheet_name=info["sheet_name"],
        data=data,
        columns=cols,
        upper_annotations=upper,
        lower_annotations=lower,
        row_indices=row_indices
    )
