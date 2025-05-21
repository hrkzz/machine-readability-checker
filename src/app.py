import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

from src.checker.router import run_checks_from_rules
from src.summary import summarize_results

st.set_page_config(page_title="機械可読性チェック（レベル1〜3）", layout="wide")
st.title("📊 機械可読性診断ツール（レベル1 → 2 → 3）")

uploaded_file = st.file_uploader("CSV または Excel ファイルをアップロードしてください", type=["csv", "xlsx"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=uploaded_file.name) as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        tmp_path = tmp_file.name

    ext = os.path.splitext(tmp_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(tmp_path)
        workbook = None
    else:
        df = pd.read_excel(tmp_path)
        workbook = load_workbook(tmp_path)

    # 判定ボタンの表示
    if st.button("📋 判定開始"):
        with st.spinner("レベル1〜3 の可読性チェックを実行中..."):
            results = []
            for level in ["level1", "level2", "level3"]:
                rule_file = f"rules/{level}.json"
                level_results = run_checks_from_rules(
                    rule_file=rule_file,
                    df=df,
                    workbook=workbook,
                    filepath=tmp_path,
                    level=level
                )
                results.append((level, level_results))

        # ✅ サマリー＋LLM所見の表示
        st.subheader("📌 判定結果サマリー")
        summary, llm_comment = summarize_results(results)

        for level in ["level1", "level2", "level3"]:
            passed, total = summary[level]
            st.markdown(f"**{level.upper()}：{passed}/{total} 項目 合格**")

        st.markdown("### 🧠 所見（LLMによる要約）")
        st.info(llm_comment)

        # ✅ 詳細表示（折りたたみ式）
        for level, checks in results:
            with st.expander(f"🔍 {level.upper()} チェックの詳細を表示"):
                for item in checks:
                    st.markdown(f"**{item['id']} - {item['description']}**")
                    st.markdown(f"- 判定: {'✅ OK' if item['result'] == '✓' else '❌ NG'}")
                    st.markdown(f"- 詳細: {item['message']}")
                    st.markdown("---")

        # ✅ Markdown形式のレポート生成
        report_lines = [
            "# 機械可読性チェックレポート（レベル1〜3）\n",
            f"ファイル名: {uploaded_file.name}\n",
            "---\n",
            "## 所見（LLM生成）\n",
            llm_comment + "\n",
            "---\n"
        ]

        for level in ["level1", "level2", "level3"]:
            passed, total = summary[level]
            report_lines.append(f"## {level.upper()}：{passed}/{total} 合格\n")

        for level, checks in results:
            report_lines.append(f"## {level.upper()} チェック詳細\n")
            for item in checks:
                report_lines.append(f"### {item['id']} - {item['description']}\n")
                report_lines.append(f"- 判定: {item['result']}\n")
                report_lines.append(f"- 詳細: {item['message']}\n")
                report_lines.append("\n")
            report_lines.append("---\n")

        report_str = "\n".join(report_lines)
        st.download_button(
            label="📥 レポートをダウンロード",
            data=report_str,
            file_name="readability_report_level1_2_3.md",
            mime="text/markdown"
        )
