import json
from pathlib import Path
from typing import Any

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
    新しいファクトリーパターンを使用してファイル形式に応じたチェッカーを選択
    """
    with open(rule_file, encoding="utf-8") as f:
        rules = json.load(f)
    
    # ファイルパスから適切なチェッカーを取得
    file_path = Path(filepath)
    checker = checker_factory.get_checker(file_path, level)
    
    results = []
    for rule in rules:
        function_name = rule["function"]
        
        # チェッカーから関数を取得
        if not hasattr(checker, function_name):
            results.append({
                "id": rule["id"],
                "description": rule["description"],
                "result": "✗",
                "message": f"関数 '{function_name}' がチェッカーに実装されていません"
            })
            continue
            
        fn = getattr(checker, function_name)
        try:
            passed, msg = fn(ctx, workbook, filepath)
        except Exception as e:
            passed, msg = False, f"エラー発生: {e}"
            
        results.append({
            "id": rule["id"],
            "description": rule["description"],
            "result": "✓" if passed else "✗",
            "message": msg
        })
    
    return results
