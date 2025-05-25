from src.llm.llm_client import call_llm

def summarize_results(results_per_level):
    summary = {}
    lines_for_prompt = []
    table_lines = []

    table_lines.append("| チェックレベル | 合格数 | 全体数 | 合格率 |")
    table_lines.append("|----------------|--------|--------|--------|")

    for level, checks in results_per_level:
        total = len(checks)
        passed = sum(1 for item in checks if item["result"] == "✓")
        rate = f"{(passed / total * 100):.0f}%" if total > 0 else "N/A"
        summary[level] = (passed, total)

        table_lines.append(f"| {level.upper()} | {passed} | {total} | {rate} |")

        lines_for_prompt.append(f"【{level.upper()}】")
        for item in checks:
            status = "OK" if item["result"] == "✓" else "NG"
            lines_for_prompt.append(f"{item['id']} ({item['description']}): {status}")

    prompt = (
        "以下は機械可読性診断の評価結果です。\n"
        "2〜3文の要約から構成されるリード文と、それぞれのレベルにおける主な問題点と、改善手順を簡潔に日本語でまとめてください。\n"
        + "\n".join(lines_for_prompt)
    )

    llm_comment = call_llm(prompt)

    summary_md = "### チェック結果サマリー\n\n" + "\n".join(table_lines)
    return summary, summary_md, llm_comment
