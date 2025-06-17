import pandas as pd
from pathlib import Path
from loguru import logger


def detect_multiple_tables_csv(df: pd.DataFrame, sheet_name: str = "CSV") -> tuple:
    """
    CSV用の複数テーブル検出（DataFrame専用）
    
    Args:
        df: 対象のDataFrame
        sheet_name: シート名（ログ用）
    
    Returns:
        tuple: (has_multiple_tables: bool, details: str)
    """
    try:
        if df.empty or len(df) < 3:
            return False, "データが少ないため複数テーブルの検出をスキップ"
        
        # 完全に空の行を検索
        empty_rows = []
        for idx, row in df.iterrows():
            if row.isna().all() or (row.astype(str).str.strip() == "").all():
                empty_rows.append(idx)
        
        # 連続する空行を検索（テーブル区切りの可能性）
        if len(empty_rows) > 0:
            consecutive_groups = []
            current_group = [empty_rows[0]]
            
            for i in range(1, len(empty_rows)):
                if empty_rows[i] == empty_rows[i-1] + 1:
                    current_group.append(empty_rows[i])
                else:
                    if len(current_group) >= 2:  # 2行以上の連続空行
                        consecutive_groups.append(current_group)
                    current_group = [empty_rows[i]]
            
            if len(current_group) >= 2:
                consecutive_groups.append(current_group)
            
            if consecutive_groups:
                return True, f"複数の連続空行グループが見つかりました: {len(consecutive_groups)}箇所"
        
        # ヘッダー様の行の検出
        header_like_rows = []
        for idx, row in df.iterrows():
            non_na_values = row.dropna().astype(str).str.strip()
            if len(non_na_values) > 0:
                # 数値以外が多い行をヘッダー候補とする
                numeric_count = sum(1 for val in non_na_values if val.replace('.', '').replace('-', '').isdigit())
                if numeric_count / len(non_na_values) < 0.5:
                    header_like_rows.append(idx)
        
        # 複数のヘッダー様行が離れて存在する場合
        if len(header_like_rows) >= 2:
            gaps = [header_like_rows[i+1] - header_like_rows[i] for i in range(len(header_like_rows)-1)]
            if any(gap > 3 for gap in gaps):  # 3行以上離れたヘッダーがある
                return True, f"離れた位置に複数のヘッダー様行が検出されました: {header_like_rows}"
        
        return False, "単一テーブルと判定"
        
    except Exception as e:
        logger.error(f"CSV複数テーブル検出でエラー: {e}")
        return False, f"検出処理でエラーが発生: {str(e)}"


def csv_specific_check(file_path: Path) -> dict:
    """
    CSV固有のチェック処理
    """
    return {
        "encoding_check": True,  # CSV は読み込み時点でエンコーディングチェック済み
        "delimiter_check": True,  # パンダスが自動検出
        "quote_check": True,      # パンダスが自動処理
    } 