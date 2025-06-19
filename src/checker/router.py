import json
from pathlib import Path
from typing import Any
from loguru import logger

from src.processor.context import TableContext
from .factory import checker_factory


def run_checks_from_rules(
    rule_file: str,
    ctx: TableContext,
    workbook: Any,
    filepath: str,
    level: str
) -> list[dict[str, Any]]:
    """
    ルールファイルに従ってチェックを実行し、結果を返す
    ファクトリーパターンを使用してファイル形式に応じたチェッカーを選択
    """
    level_upper = level.upper()
    sheet_name = ctx.sheet_name
    
    # レベル開始ログ
    logger.info(f"===== {level_upper} チェック開始 =====")
    logger.info(f"▼ シート: {sheet_name} の{level_upper}チェック開始")
    
    with open(rule_file, encoding="utf-8") as f:
        rules = json.load(f)
    
    # ファイルパスから適切なチェッカーを取得
    file_path = Path(filepath)
    checker = checker_factory.get_checker(file_path, level)
    
    results = []
    success_count = 0
    warning_count = 0
    error_count = 0
    
    for rule in rules:
        function_name = rule["function"]
        
        # チェッカーから関数を取得
        if not hasattr(checker, function_name):
            error_message = f"関数 '{function_name}' がチェッカーに実装されていません"
            logger.error(f"[{level_upper}] {function_name}: エラー - {error_message}")
            results.append({
                "id": rule["id"],
                "description": rule["description"],
                "result": "✗",
                "message": error_message
            })
            error_count += 1
            continue
            
        fn = getattr(checker, function_name)
        try:
            passed, msg = fn(ctx, workbook, filepath)
            
            # 結果に応じたログ出力
            if passed:
                logger.info(f"[{level_upper}] {function_name}: OK")
                success_count += 1
            else:
                logger.warning(f"[{level_upper}] {function_name}: 問題検出 - {msg[:100]}...")
                warning_count += 1
                
        except Exception as e:
            passed, msg = False, f"エラー発生: {e}"
            logger.error(f"[{level_upper}] {function_name}: エラー - {str(e)}")
            error_count += 1
            
        results.append({
            "id": rule["id"],
            "description": rule["description"],
            "result": "✓" if passed else "✗",
            "message": msg
        })
    
    # レベル終了ログとサマリー
    total_checks = len(rules)
    logger.info(f"▲ シート: {sheet_name} の{level_upper}チェック終了")
    logger.info(f"[{level_upper}] 結果サマリー - 成功:{success_count}, 警告:{warning_count}, エラー:{error_count} (全{total_checks}件)")
    logger.info(f"===== {level_upper} チェック終了 =====")
    
    return results
