from typing import Dict, Any, List, cast
import pandas as pd
import json
import re
import textwrap
from pathlib import Path
import base64
import openai

from src.config import PREVIEW_ROW_COUNT
from src.llm.llm_client import call_llm
from src.llm.llm_client import call_llm_with_image
from .context import TableContext


def select_main_sheet(sheet_infos: List[Dict[str, Any]]) -> Dict[str, Any]:
    prompt_parts = []
    for sheet in sheet_infos:
        preview = sheet["preview_top"].fillna("").astype(str).values.tolist()
        text = "\n".join(",".join(row) for row in preview)
        prompt_parts.append(f"[{sheet['sheet_name']}]\n{text}")

    prompt = (
        f"以下は複数のExcelシートの先頭{PREVIEW_ROW_COUNT}行です。\n"
        "業務上・分析上のメインテーブルに該当するシート名のみを、正確に1つだけ出力してください。\n\n"
        + "\n\n".join(prompt_parts)
    )

    response = call_llm(prompt)
    for sheet in sheet_infos:
        if sheet["sheet_name"] in response:
            return sheet

    return sheet_infos[0]


def analyze_table_structure(sheet: Dict[str, Any]) -> Dict[str, Any]:
    df = sheet["dataframe"]
    total_rows = df.shape[0]
    top = sheet["preview_top"].fillna("").astype(str).values.tolist()
    bot = sheet["preview_bottom"].fillna("").astype(str).values.tolist()
    content = "\n".join(",".join(row) for row in (top + [["..."]] + bot))

    prompt = textwrap.dedent(f"""\
        以下はシート「{sheet['sheet_name']}」の先頭{PREVIEW_ROW_COUNT}行と末尾{PREVIEW_ROW_COUNT}行です。
        表の構造を解析して、以下のJSON形式（1-indexed）で返してください：
        {{
          "column_rows": [],
          "data_start": ,
          "data_end": ,
          "annotation_rows": []
        }}
        プレビュー:
        {content}
    """)

    raw_response = call_llm(prompt)

    match = re.search(r"{[\s\S]+}", raw_response)
    if match:
        try:
            parsed = json.loads(match.group())
        except json.JSONDecodeError as e:
            raise ValueError(f"構造解析のレスポンスがJSONとして解析できません: {e}")
    else:
        raise ValueError("構造解析レスポンスからJSONを抽出できませんでした")

    return {
        "sheet_name": sheet["sheet_name"],
        "dataframe": df,
        "structure_response": parsed
    }


def analyze_table_structure_by_image(sheet: Dict[str, Any], image_path: Path) -> Dict[str, Any]:
    prompt = """
これは政府統計のExcel表の画像です。以下の構造をJSON形式（1-indexed）で返してください：

{
  "column_rows": [],
  "data_start": ,
  "data_end": ,
  "annotation_rows": []
}
"""
    response = call_llm_with_image(prompt, image_path)

    match = re.search(r"{[\s\S]+}", response)
    if match:
        parsed = json.loads(match.group())
    else:
        raise ValueError("画像構造解析レスポンスからJSONを抽出できませんでした")

    return {
        "sheet_name": sheet["sheet_name"],
        "dataframe": sheet["dataframe"],
        "structure_response": parsed
    }

def parse_sheet_structure(sheet: Dict[str, Any]) -> Dict[str, Any]:
    try:
        return analyze_table_structure(sheet)
    except Exception as e:
        print(f"⚠️ テキストベース解析に失敗: {e}")
        if "image_path" not in sheet:
            raise ValueError("画像パスが指定されていないため、画像ベース構造解析に切り替えできません")
        image_path = Path(sheet["image_path"])
        return analyze_table_structure_by_image(sheet, image_path)


def extract_structured_table(info: Dict[str, Any]) -> TableContext:
    raw_resp = info["structure_response"]

    if isinstance(raw_resp, dict):
        parsed = raw_resp
    else:
        m = re.search(r"json\s*(\{.*?\})\s*", raw_resp, re.DOTALL)
        js = m.group(1) if m else raw_resp.strip()
        try:
            parsed = json.loads(js)
        except json.JSONDecodeError as e:
            raise ValueError(f"構造出力のパースに失敗しました: {e}")

    def to_int_list(values):
        return [int(v) for v in values if isinstance(v, (int, str)) and str(v).strip().isdigit()]
    
    column_rows = [i - 1 for i in to_int_list(parsed.get("column_rows", []))]
    data_start  = int(parsed["data_start"]) - 1
    data_end    = int(parsed["data_end"]) - 1
    annotations = [i - 1 for i in to_int_list(parsed.get("annotation_rows", []))]

    df = info["dataframe"]

    upper = df.iloc[:min(column_rows)].copy()
    lower = df.iloc[data_end + 1:].copy()

    col_df = df.iloc[column_rows].fillna("").astype(str)
    if len(column_rows) > 1:
        arrays_fixed = []
        for level in col_df.values:
            fixed = []
            last_val = ""
            for val in level:
                if val == "":
                    val = last_val or "(空白)"
                else:
                    last_val = val
                fixed.append(val)
            arrays_fixed.append(fixed)
        cols = pd.MultiIndex.from_arrays(arrays_fixed)
    else:
        cols = col_df.iloc[0].tolist()

    data = df.iloc[data_start : data_end + 1].copy()
    data.columns = cols
    data.reset_index(drop=True, inplace=True)

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
