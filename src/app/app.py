import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

from src.checker.router import run_checks_from_rules
from src.summary import summarize_results

st.set_page_config(page_title="機械可読性チェック", layout="wide")

# CSS の読み込みと適用（存在しない場合に備えて例外処理）
css_path = os.path.join("src", "app", "styles", "style.css")
if os.path.exists(css_path):
    with open(css_path) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title("機械可読性チェックツール")
st.markdown("ファイルの内容に基づき、レベル1〜3の可読性チェックを実行します。")

# ファイルアップロード
uploaded_file = st.file_uploader("CSV または Excel ファイルをアップロード", type=["csv", "xlsx"])

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
        workbook = load_workbook(tmp_path, data_only=True)

    # 判定実行ボタン
    if st.button("チェックを実行"):
        with st.spinner("チェック中..."):
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

            summary, llm_comment = summarize_results(results)

        # サマリー表示
        st.subheader("チェック結果サマリー")
        for level in ["level1", "level2", "level3"]:
            passed, total = summary[level]
            st.markdown(f"**{level.upper()}**：{passed} / {total} 項目 合格")

        st.markdown("### LLMによる総評")
        st.write(llm_comment)

        # 詳細結果（折りたたみ）
        for level, checks in results:
            with st.expander(f"{level.upper()} チェックの詳細"):
                for item in checks:
                    st.markdown(f"**{item['id']} – {item['description']}**")
                    st.markdown(f"- 判定: {'合格' if item['result'] == '✓' else '不合格'}")
                    st.markdown(f"- 詳細: {item['message']}")
                    st.markdown("---")

        # Markdownレポートの作成
        report_lines = [
            "# 機械可読性チェックレポート（レベル1〜3）",
            f"ファイル名: {uploaded_file.name}",
            "",
            "## 総評",
            llm_comment,
            ""
        ]

        for level in ["level1", "level2", "level3"]:
            passed, total = summary[level]
            report_lines.append(f"## {level.upper()}：{passed}/{total} 合格")

        for level, checks in results:
            report_lines.append(f"\n### {level.upper()} チェック詳細")
            for item in checks:
                report_lines.append(f"#### {item['id']} – {item['description']}")
                report_lines.append(f"- 判定: {item['result']}")
                report_lines.append(f"- 詳細: {item['message']}\n")

        report_str = "\n".join(report_lines)

        st.download_button(
            label="レポートをダウンロード",
            data=report_str,
            file_name="readability_report.md",
            mime="text/markdown"
        )
