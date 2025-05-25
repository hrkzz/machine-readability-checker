from typing import Dict, Any, List, cast
import pandas as pd
import json
import re
import textwrap

from src.llm.llm_client import call_llm
from .context import TableContext

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
    total_rows = df.shape[0]
    top = sheet["preview_top"].fillna("").astype(str).values.tolist()
    bot = sheet["preview_bottom"].fillna("").astype(str).values.tolist()
    content = "\n".join(",".join(row) for row in (top + [["..."]] + bot))

    prompt = textwrap.dedent(f"""\
        以下はシート「{sheet['sheet_name']}」の先頭10行と末尾10行です。
        このシートは以下の特徴があります：
        - セル結合（merged cells）が含まれる可能性があります。
        - ヘッダーが複数行（マルチインデックス）にまたがっている可能性があります。
        - 行の一部が注釈やタイトルなど、データではないメタ情報を含むことがあります。
        - 合計 {total_rows} 行存在します。

        この表の構造を解析して、以下のJSON形式（1-indexed）で返してください：
        {{
          "column_rows": [],
          "data_start": ,
          "data_end": ,
          "annotation_rows": []
        }}

        出力例：
        json
        {{
          "column_rows": [3, 4],
          "data_start": 5,
          "data_end": 29,
          "annotation_rows": [1, 30]
        }}

        判定ルールのヒント：
        1. column_rows（列見出し、データの列名に相当）
        - データを列方向に整理するための名前（例えば「品目番号」「項目名」「年月日」など）。
        - **表タイトルや注釈（例：「項目：総合季節調整指数」など）は含めないこと。**
        - 必要に応じて、複数行にわたるマルチインデックスの可能性あり。
        - セル結合や「単位」などが含まれることもある。

        2. annotation_rows（タイトル・出典・備考などの説明行）
        - 表の上部や下部に位置し、「表タイトル」「注記」「出典」などを含む行。
        - 多くの列で文字列のみで構成されている。
        - 表の構造に含めたくない場合、すべてこちらに分類する。

        3. data_start / data_end（データ行の範囲）
        - 実データが含まれており、多くのセルが数値型。
        - column_rows の直後から始まり、連続する構造。

        プレビュー:
        {content}

    """)

    raw_response = call_llm(prompt)

    # JSON部分だけ抽出（```json ... ``` 形式のガードを削除）
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
        "structure_response": parsed  # ← JSONそのものを格納
    }


def extract_structured_table(info: Dict[str, Any]) -> TableContext:
    """
    LLMのJSON応答をパースし、TableContextを構築して返す。
    マルチ行ヘッダーも pandas.MultiIndex で扱う。
    """
    raw_resp = info["structure_response"]

    # raw_resp が dict か str かで処理を分岐
    if isinstance(raw_resp, dict):
        parsed = raw_resp
    else:
        # ```json {...} ``` 形式が含まれる場合は中身を抽出
        m = re.search(r"```json\s*(\{.*?\})\s*```", raw_resp, re.DOTALL)
        js = m.group(1) if m else raw_resp.strip()
        try:
            parsed = json.loads(js)
        except json.JSONDecodeError as e:
            print("=== JSONパース失敗 ===")
            print(raw_resp)
            raise ValueError(f"構造出力のパースに失敗しました: {e}")

    # インデックス（0-based）に調整
    column_rows = [i - 1 for i in parsed["column_rows"]]
    data_start   = parsed["data_start"] - 1
    data_end     = parsed["data_end"]   - 1
    annotations  = [i - 1 for i in parsed.get("annotation_rows", [])]

    df = info["dataframe"]

    # 上部注釈（ヘッダー行より上）
    upper = df.iloc[:min(column_rows)].copy()
    # 下部注釈（データ行の下）
    lower = df.iloc[data_end + 1:].copy()

    # カラム行の抽出
    col_df = df.iloc[column_rows].fillna("").astype(str)
    if len(column_rows) > 1:
        arrays = col_df.values
        arrays_fixed = []
        for level in arrays:
            fixed_level = []
            last_val = ""
            for val in level:
                if val == "":
                    val = last_val or "(空白)"
                else:
                    last_val = val
                fixed_level.append(val)
            arrays_fixed.append(fixed_level)
        cols = pd.MultiIndex.from_arrays(arrays_fixed)
    else:
        cols = col_df.iloc[0].tolist()

    # 実データ部分を切り出し
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
