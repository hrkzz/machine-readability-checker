import json
import importlib
from typing import List, Dict, Any

def run_checks_from_rules(rule_file: str, df, workbook, filepath: str, level: str = "level1") -> List[Dict[str, Any]]:
    """
    ルールファイルを読み込み、各チェック関数を実行し、結果を返す。
    """
    with open(rule_file, "r", encoding="utf-8") as f:
        rules = json.load(f)

    # level1_checks モジュールを動的に import
    check_module = importlib.import_module(f"src.checker.{level}_checks")

    results = []
    for rule in rules:
        func_name = rule["function"]
        func = getattr(check_module, func_name)

        try:
            passed, message = func(df=df, workbook=workbook, filepath=filepath)
        except Exception as e:
            passed, message = False, f"エラー発生: {e}"

        results.append({
            "id": rule["id"],
            "description": rule["description"],
            "result": "✓" if passed else "✗",
            "message": message
        })

    return results
