import re
import random
import pandas as pd
from typing import Optional, Tuple
from ..llm.llm_client import call_llm

def check_numeric_columns_only(
    df: Optional[pd.DataFrame] = None,
    workbook=None,
    filepath: Optional[str] = None
) -> Tuple[bool, str]:
    if df is None:
        return False, "DataFrameが指定されていません"

    likely_numeric_cols = []
    problem_details = {}

    def is_clean_numeric(value):
        if isinstance(value, (int, float)):
            return True
        if isinstance(value, str):
            value = value.strip()
            if re.search(r"[^\d.\-]", value):
                return False
            try:
                float(value)
                return True
            except ValueError:
                return False
        return False

    for col in df.columns:
        values = df[col].dropna()
        total = len(values)
        numeric_like_count = values.apply(is_clean_numeric).sum()

        if total == 0:
            continue

        ratio = numeric_like_count / total
        if ratio >= 0.8:
            likely_numeric_cols.append(col)
            problem_values = values[~values.apply(is_clean_numeric)].unique().tolist()
            if problem_values:
                problem_details[col] = problem_values

    if problem_details:
        messages = [f"{col}: {vals[:3]}" for col, vals in problem_details.items()]
        return False, "数値列に数値以外が含まれています:\n" + "\n".join(messages)

    return True, "数値列に不正なデータは含まれていません"

def check_separate_other_detail_columns(
    df: Optional[pd.DataFrame] = None,
    workbook=None,
    filepath: Optional[str] = None
) -> Tuple[bool, str]:
    if df is None:
        return False, "DataFrameが指定されていません"

    flagged_columns = []

    for col in df.columns:
        try:
            if not pd.api.types.is_string_dtype(df[col]):
                continue

            values = df[col].dropna().astype(str).unique().tolist()
            sample = values[:20]

            prompt = f"""
以下は、列名「{col}」に含まれるデータの一部です。
この列には「選択肢（例：読書、ゲーム、料理）」のような値と、
「その他（〜）」のように自由記述が含まれた値が**混在しているか**を判定してください。

このような混在は、本来別カラムに分けるべき構造とされます。
混在が確認できた場合は「混在している」、そうでない場合は「混在していない」とだけ答えてください。

▼列データサンプル：
{chr(10).join(sample)}
"""
            result = call_llm(prompt)
            if "混在している" in result:
                flagged_columns.append(col)

        except Exception as e:
            return False, f"列「{col}」のチェック中にエラーが発生しました: {e}"

    if flagged_columns:
        return False, f"選択肢と「その他」の詳細が混在している可能性のある列: {flagged_columns}"

    return True, "選択肢と「その他」の詳細は適切に分離されています"

def check_no_missing_column_headers(
    df: Optional[pd.DataFrame] = None,
    workbook=None,
    filepath: Optional[str] = None
) -> Tuple[bool, str]:
    if df is None:
        return False, "DataFrameが指定されていません"

    unnamed = [col for col in df.columns if "Unnamed" in str(col) or str(col).strip() == ""]
    suspect = []

    for col in df.columns:
        if col in unnamed:
            continue

        prompt = f"""
以下の列名が表のヘッダーとして自然かどうかを判定してください。

### 判断基準
- 明確：意味が明瞭で、何のデータを表しているかが読み手にとって直感的にわかる
- 不明：意味が曖昧、略語すぎる、または内容が推測しづらい。明らかに日本語として間違えている

### 対象
列名: {col}

### 回答形式（必ず以下のいずれか1語で）
明確 または 不明
"""
        result = call_llm(prompt)
        if "不明" in result:
            suspect.append(col)

    if unnamed or suspect:
        issues = unnamed + suspect
        return False, f"省略・不明な列名が検出されました: {issues}"
    return True, "すべての列に意味のある項目名が付けられています"

def check_handling_of_missing_values(
    df: Optional[pd.DataFrame] = None,
    workbook=None,
    filepath: Optional[str] = None
) -> Tuple[bool, str]:
    if df is None:
        return False, "DataFrameが指定されていません"

    missing_like = [
        "不明", "不詳", "無記入", "無回答", "該当なし", "なし", "無し", "n/a", "na", "nan",
        "未定", "未記入", "未入力", "未回答", "記載なし", "対象外", "空欄", "空白", "不在",
        "特になし", "---", "--", "-", "ー", "―", "？", "?", "わからない", "わかりません",
        "なし（特記なし）", "無し（詳細不明）", "無効", "省略", "null", "none"
    ]

    def normalize_text(text: str) -> str:
        return str(text).strip().lower().replace("（", "(").replace("）", ")")

    suspect_cols = []

    for idx, col in enumerate(df.columns):
        try:
            values = df[col].dropna().astype(str)
            normalized = [normalize_text(v) for v in values]
            count_like_missing = sum(1 for v in normalized if any(m in v for m in missing_like))
            ratio = count_like_missing / len(normalized) if normalized else 0

            if ratio > 0.2:
                sample = random.sample(values.tolist(), min(10, len(values)))

                prompt = f"""
次の列「{col}」（{idx + 1}列目）には、欠損値や未記入を意味する表現が含まれている可能性があります。
これらの表現が「一貫していない」または「曖昧・ばらつきがある」かを判定してください。

データサンプル：
{sample}

回答形式：「欠損表現あり」または「なし」のどちらか一語で答えてください。
"""
                result = call_llm(prompt)
                if "欠損表現あり" in result:
                    suspect_cols.append(f"{idx + 1}列目（{col}）")

        except Exception as e:
            return False, f"列「{col}」の処理中にエラーが発生しました: {e}"

    if suspect_cols:
        return False, f"実質的な欠損表現が検出されました（表現の統一が必要）: {suspect_cols}"
    return True, "欠損の扱いに一貫性があり、問題ありません"
