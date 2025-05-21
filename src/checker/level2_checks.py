import pandas as pd
from ..llm.llm_client import call_llm

def check_numeric_columns_only(df=None, workbook=None, filepath=None):
    """
    データ型から数値列を推定し、その列だけ数値混入チェックを行う
    """
    likely_numeric_cols = []
    problem_cols = []

    for col in df.columns:
        # 値が数値っぽいものの割合が高い列を「数値列候補」とする
        values = df[col].dropna()
        numeric_like_count = values.apply(lambda x: isinstance(x, (int, float)) or str(x).replace('.', '', 1).isdigit()).sum()
        ratio = numeric_like_count / len(values) if len(values) > 0 else 0

        if ratio > 0.8:  # 閾値：80%以上が数値っぽい
            likely_numeric_cols.append(col)
            # 実際に変換できない値が含まれるかをチェック
            try:
                pd.to_numeric(values)
            except ValueError:
                problem_cols.append(col)

    if problem_cols:
        return False, f"数値列に数値以外が含まれています: {problem_cols}"
    return True, "数値列に不正なデータは含まれていません"

def check_separate_other_detail_columns(df=None, workbook=None, filepath=None):
    """
    「その他」の自由記述が選択肢列と混ざっていないかチェック（列名から推測）
    """
    problem_columns = [col for col in df.columns if "その他" in str(col) and "自由" not in str(col)]
    if problem_columns:
        return False, f"その他の記述が分離されていない可能性のある列: {problem_columns}"
    return True, "その他の詳細記述は適切に分離されています"

def check_no_missing_column_headers(df=None, workbook=None, filepath=None):
    unnamed = [col for col in df.columns if "Unnamed" in str(col) or str(col).strip() == ""]
    suspect = []

    for col in df.columns:
        if col in unnamed:
            continue
        prompt = f"""
以下の列名が表のヘッダーとして自然かどうかを判定してください。

### 判断基準
- 明確：意味が明瞭で、何のデータを表しているかが読み手にとって直感的にわかる
- 不明：意味が曖昧、略語すぎる、または内容が推測しづらい

### 例
- 「氏名」 → 明確
- 「備考」 → 明確
- 「名」 → 明確
- 「ABC」 → 不明
- 「a1」 → 不明
- 「」 → 不明
- 「Unnamed: 0」 → 不明

### 対象
列名: {col}

### 回答形式（必ず以下のいずれか1語で）
明確 または 不明
"""
        result = call_llm(prompt)
        if "不明" in result:
            suspect.append(col)

    if unnamed or suspect:
        all_issues = unnamed + suspect
        return False, f"省略・不明な列名が検出されました: {all_issues}"
    return True, "すべての列に意味のある項目名が付けられています"

def check_handling_of_missing_values(df=None, workbook=None, filepath=None):
    """
    欠損値の割合が高い列や、扱いに一貫性がない列がないか（暫定）
    """
    cols_with_na = df.columns[df.isna().any()].tolist()
    if not cols_with_na:
        return True, "欠損値はありません"
    else:
        return False, f"欠損値が存在する列があります: {cols_with_na}（明示的な扱いが必要）"
