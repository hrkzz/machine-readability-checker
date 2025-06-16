from typing import Dict, Any, List, cast
import pandas as pd
import json
import re
import textwrap
from loguru import logger

from src.config import PREVIEW_ROW_COUNT
from src.llm.llm_client import call_llm
from .context import TableContext

def select_main_sheet(sheet_infos: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    複数シートの先頭をLLMに渡し、メインテーブルと判断されたシートの情報を返す。
    """
    prompt_parts = []
    for sheet in sheet_infos:
        preview = sheet["preview_top"].fillna("").astype(str).values.tolist()
        text = "\n".join(",".join(row) for row in preview)
        prompt_parts.append(f"[{sheet['sheet_name']}]\n{text}")

    prompt = (
        "以下は複数のExcelシートの先頭{PREVIEW_ROW_COUNT}行です。\n"
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
    メインシートの先頭・末尾をLLMに渡し、構造情報（カラム行、データ行、注釈行）を取得する。
    """
    df = sheet["dataframe"]
    total_rows = df.shape[0]
    
    # データが空の場合のデフォルト処理
    if total_rows == 0:
        return {
            "sheet_name": sheet["sheet_name"],
            "dataframe": df,
            "structure_response": {
                "column_rows": [1],
                "data_start": 2,
                "data_end": 1,
                "annotation_rows": []
            }
        }
    
    top = sheet["preview_top"].fillna("").astype(str).values.tolist()
    bot = sheet["preview_bottom"].fillna("").astype(str).values.tolist()
    content = "\n".join(",".join(row) for row in (top + [["..."]] + bot))

    prompt = textwrap.dedent(f"""\
        以下はシート「{sheet['sheet_name']}」の先頭{PREVIEW_ROW_COUNT}行と末尾{PREVIEW_ROW_COUNT}行です。
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
        - 表タイトルや注釈（例：「項目：総合季節調整指数」など）は含めないこと。
        - 必要に応じて、複数行にわたるマルチインデックスの可能性あり。
        - セル結合や「単位」などが含まれることもある。

        2. annotation_rows（タイトル・出典・備考などの説明行）
        - 表の上部や下部に位置し、「表タイトル」「注記」「出典」などを含む行。
        - 多くの列で文字列のみで構成されている。
        - 表の構造に含めたくない場合、すべてこちらに分類する。

        3. data_start / data_end（データ行の範囲）
        - 実データが含まれており、多くのセルが数値型。
        - column_rows の直後から始まり、連続する構造。

        注意点
        column_rows, data_start, data_end, annotation_rowsが重なることはありません。
        この表は Wide形式（列数が非常に多い）または統計表形式である可能性があります。
        - 列名が複数行にまたがっている場合（例：男女別×指標別）、column_rows は 2〜4 行を含むことがあります。
        - 表タイトルや補足文（例：「第2表 就業状態別15歳以上人口」など）は annotation_rows に分類し、column_rows に含めないでください。
        - LLMがトークン制限で一部の列しか見えない場合も、構造の一貫性から column_rows を類推して判断してください。

        プレビュー:
        {content}


    """)

    try:
        raw_response = call_llm(prompt)

        # JSON部分だけ抽出（```json ... ``` 形式のガードを削除）
        match = re.search(r"{[\s\S]+}", raw_response)
        if match:
            try:
                parsed = json.loads(match.group())
            except json.JSONDecodeError as e:
                logger.error(f"JSON解析エラー: {e}")
                logger.debug(f"レスポンス: {raw_response}")
                raise ValueError(f"構造解析のレスポンスがJSONとして解析できません: {e}")
        else:
            logger.error(f"JSON抽出失敗。レスポンス: {raw_response}")
            raise ValueError("構造解析レスポンスからJSONを抽出できませんでした")

        # 必要なフィールドの検証
        required_fields = ["column_rows", "data_start", "data_end"]
        for field in required_fields:
            if field not in parsed or parsed[field] is None:
                logger.error(f"必要フィールド不足: {field}")
                raise ValueError(f"必要なフィールドが不足: {field}")

        return {
            "sheet_name": sheet["sheet_name"],
            "dataframe": df,
            "structure_response": parsed  # ← JSONそのものを格納
        }
        
    except Exception as e:
        logger.error(f"構造解析でエラーが発生しました: {e}")
        logger.info("デフォルトの構造を適用します")
        
        # フォールバック: シンプルなデフォルト構造を適用
        if total_rows <= 2:
            # 非常に小さなデータセット
            default_structure = {
                "column_rows": [1],  # 1行目をヘッダーとする
                "data_start": 2 if total_rows >= 2 else 1,     # 2行目からデータ開始（または1行目）
                "data_end": total_rows,  # 最終行まで
                "annotation_rows": []
            }
        else:
            # 通常のデータセット
            default_structure = {
                "column_rows": [1],  # 1行目をヘッダーとする
                "data_start": 2,     # 2行目からデータ開始
                "data_end": min(total_rows, 100),  # 最大100行または総行数まで
                "annotation_rows": []
            }
        
        return {
            "sheet_name": sheet["sheet_name"],
            "dataframe": df,
            "structure_response": default_structure
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
            logger.error("=== JSONパース失敗 ===")
            logger.debug(raw_resp)
            raise ValueError(f"構造出力のパースに失敗しました: {e}")

    # 必要なフィールドの存在確認とデバッグ情報出力
    required_fields = ["column_rows", "data_start", "data_end"]
    missing_fields = []
    for field in required_fields:
        if field not in parsed or parsed[field] is None:
            missing_fields.append(field)
    
    if missing_fields:
        logger.debug("=== 構造解析レスポンスのデバッグ情報 ===")
        logger.debug(f"パース結果: {parsed}")
        logger.debug(f"不足フィールド: {missing_fields}")
        raise ValueError(f"構造解析で必要なフィールドが不足しています: {missing_fields}")

    # インデックス（0-based）に調整
    column_rows = [i - 1 for i in parsed["column_rows"]] if isinstance(parsed["column_rows"], list) else [parsed["column_rows"] - 1]
    data_start   = parsed["data_start"] - 1
    data_end     = parsed["data_end"]   - 1
    annotations  = [i - 1 for i in parsed.get("annotation_rows", [])]

    df = info["dataframe"]

    # データが非常に少ない場合の特別処理
    if len(df) <= 1:
        logger.info(f"=== 非常に小さなデータセット (行数: {len(df)}) ===")
        # 1行以下の場合はすべてをヘッダーとして処理
        return TableContext(
            sheet_name=info["sheet_name"],
            data=pd.DataFrame(),  # 空のデータ
            columns=df.iloc[0].tolist() if len(df) == 1 else [],
            upper_annotations=pd.DataFrame(),
            lower_annotations=pd.DataFrame(),
            row_indices={
                "column_rows": [0] if len(df) == 1 else [],
                "data_start": 0,
                "data_end": -1,  # データなし
                "annotations": []
            }
        )

    # データ範囲の妥当性チェック
    if data_start < 0 or data_end < 0 or data_start > data_end:
        logger.warning("=== データ範囲エラー（修正前） ===")
        logger.debug(f"data_start: {data_start}, data_end: {data_end}, df.shape: {df.shape}")
        
        # 自動修正を試行
        if data_start > data_end:
            # data_startがdata_endより大きい場合、data_endを調整
            data_end = max(data_start, len(df) - 1)
            logger.debug(f"data_endを修正: {data_end}")
        
        if data_start < 0:
            data_start = 0
            logger.debug(f"data_startを修正: {data_start}")
        
        if data_end < 0:
            data_end = len(df) - 1
            logger.debug(f"data_endを修正: {data_end}")
        
        # まだ不正な場合はデフォルト値を設定
        if data_start > data_end or data_end >= len(df):
            logger.warning("デフォルト値を適用")
            if len(column_rows) > 0:
                data_start = max(column_rows) + 1
                data_end = len(df) - 1
            else:
                data_start = 1
                data_end = len(df) - 1
            
            # 境界チェック
            data_start = min(data_start, len(df) - 1)
            data_end = min(data_end, len(df) - 1)
            data_start = max(data_start, 0)
            data_end = max(data_end, data_start)
        
        logger.info(f"修正後: data_start={data_start}, data_end={data_end}")
        
        # まだ不正な場合はエラー
        if data_start > data_end:
            raise ValueError(f"データ範囲の修正に失敗しました: data_start={data_start}, data_end={data_end}")

    # column_rowsの妥当性チェック
    if not column_rows or min(column_rows) < 0 or max(column_rows) >= len(df):
        logger.warning("=== カラム行エラー ===")
        logger.debug(f"column_rows: {column_rows}, df.shape: {df.shape}")
        
        # 自動修正
        if not column_rows:
            column_rows = [0]  # デフォルトで1行目をヘッダー
        else:
            # 範囲外の値を修正
            column_rows = [max(0, min(row, len(df) - 1)) for row in column_rows]
        
        logger.debug(f"修正後のcolumn_rows: {column_rows}")
        
        if not column_rows or min(column_rows) < 0 or max(column_rows) >= len(df):
            raise ValueError(f"カラム行の範囲修正に失敗しました: column_rows={column_rows}")

    # 上部注釈（ヘッダー行より上）
    upper = df.iloc[:min(column_rows)].copy() if column_rows else pd.DataFrame()
    # 下部注釈（データ行の下）
    lower = df.iloc[data_end + 1:].copy() if data_end + 1 < len(df) else pd.DataFrame()

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
    if data_start <= data_end and data_start < len(df) and data_end < len(df):
        data = df.iloc[data_start : data_end + 1].copy()
    else:
        logger.warning("=== データ切り出しでデフォルト処理 ===")
        logger.debug(f"data_start: {data_start}, data_end: {data_end}, df.shape: {df.shape}")
        # デフォルト: ヘッダー行以降すべてをデータとする
        if column_rows:
            start_idx = max(column_rows) + 1
            if start_idx < len(df):
                data = df.iloc[start_idx:].copy()
            else:
                data = pd.DataFrame()  # 空のデータフレーム
        else:
            data = df.copy()  # 全体をデータとする
    
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
