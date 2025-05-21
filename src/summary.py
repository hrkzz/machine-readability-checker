from src.llm.llm_client import call_llm

def summarize_results(results_per_level):
    summary = {}
    lines_for_prompt = []

    for level, checks in results_per_level:
        total = len(checks)
        passed = sum(1 for item in checks if item["result"] == "✓")
        summary[level] = (passed, total)

        lines_for_prompt.append(f"【{level.upper()}】")
        for item in checks:
            status = "OK" if item["result"] == "✓" else "NG"
            lines_for_prompt.append(f"{item['id']} ({item['description']}): {status}")

    # LLMプロンプト（所見＋改善ステップ）
    prompt = (
        "以下は機械可読性診断の各レベルにおける評価結果です。\n"
        "それぞれのレベルごとに見られた主な問題点を日本語で簡潔に要約してください。\n"
        "そのうえで、データ改善のための優先手順を提案してください。\n"
        "出力形式：\n"
        "1. レベルごとの分析\n"
        "2. ステップバイステップの改善手順\n\n"
        + "\n".join(lines_for_prompt)
    )

    llm_comment = call_llm(prompt)
    return summary, llm_comment
