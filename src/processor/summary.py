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

        for item in checks:
            status = "OK" if item["result"] == "✓" else "NG"
            lines_for_prompt.append(f"{item['id']} ({item['description']}): {status}")

    overall_comment = (
        "診断は完了しました。詳細は各チェック項目の結果を確認してください。"
    )

    summary_md = "### チェック結果サマリー\n\n" + "\n".join(table_lines)
    return summary, summary_md, overall_comment
