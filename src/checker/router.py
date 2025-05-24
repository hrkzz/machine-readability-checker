import json, importlib
from typing import Any
from src.processor.context import TableContext

MODULE_MAP = {
    "level1": "src.checker.level1_checker",
    "level2": "src.checker.level2_checker",
    "level3": "src.checker.level3_checker",
}

def run_checks_from_rules(
    rule_file: str,
    ctx: TableContext,
    workbook: Any,
    filepath: str,
    level: str
) -> list[dict[str, Any]]:
    with open(rule_file, encoding="utf-8") as f:
        rules = json.load(f)
    mod = importlib.import_module(MODULE_MAP[level])
    results = []
    for rule in rules:
        fn = getattr(mod, rule["function"])
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
