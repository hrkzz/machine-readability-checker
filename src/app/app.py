import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import streamlit as st
from openpyxl import load_workbook
import tempfile
from pathlib import Path
from datetime import datetime

from src.processor.loader import load_file
from src.processor.table_parser import (
    select_main_sheet,
    analyze_table_structure,
    extract_structured_table,
)

from src.checker.router import run_checks_from_rules
from src.summary import summarize_results

# レポートディレクトリの初期化
REPORT_DIR = Path("reports")
if REPORT_DIR.exists():
    for f in REPORT_DIR.iterdir():
        try:
            if f.is_file():
                f.unlink()
        except Exception as e:
            print(f"ファイル {f} の削除に失敗しました: {e}")
else:
    REPORT_DIR.mkdir(parents=True)

st.set_page_config(page_title="機械可読性チェック", layout="wide")

# スタイル適用（任意）
css_path = os.path.join("src", "app", "styles", "style.css")
if os.path.exists(css_path):
    with open(css_path) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title("機械可読性チェックツール")
#st.markdown("CSV / Excel ファイルをアップロードして、レベル1〜3の自動チェックを実行できます。")

uploaded_file = st.file_uploader("CSV または Excel ファイルをアップロード", type=["csv", "xlsx"])

if uploaded_file is not None:
    st.session_state["uploaded_file"] = uploaded_file
    st.session_state["uploaded_path"] = None

    with tempfile.NamedTemporaryFile(delete=False, suffix=uploaded_file.name) as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        st.session_state["uploaded_path"] = tmp_file.name

    st.info(f"{uploaded_file.name}がアップロードされました。下のボタンを押して構造解析を開始してください。")

# 構造解析の実行ボタン
if uploaded_file is not None and "structure_done" not in st.session_state:
    if st.button("構造解析を実行"):
        with st.spinner("構造解析中..."):
            file_result = load_file(Path(st.session_state["uploaded_path"]))
            main_sheet = select_main_sheet(file_result["sheets"])
            struct_info = analyze_table_structure(main_sheet)
            ctx = extract_structured_table(struct_info)
            wb = load_workbook(st.session_state["uploaded_path"], data_only=True)

            st.session_state["ctx"] = ctx
            st.session_state["workbook"] = wb
            st.session_state["structure_done"] = True

        st.success(f"メインシート「{ctx.sheet_name}」を選択し、構造を解析しました。")

# ctx / wb の初期化と安全な取得
ctx = None
wb = None
if "ctx" in st.session_state:
    ctx = st.session_state["ctx"]
    wb = st.session_state["workbook"]

# テーブル構造の表示
if ctx is not None:
    with st.expander("テーブル構造解析結果"):
        st.markdown("カラム構造")
        st.write(ctx.columns)

        st.markdown("データ（先頭5行）")
        try:
            st.dataframe(ctx.data.head())
        except Exception:
            st.warning("⚠️ 表示中にエラーが発生したため、テキスト表示に切り替えます。")
            st.code(ctx.data.head().to_string(), language="text")

        if not ctx.upper_annotations.empty:
            st.markdown("上部注釈")
            st.dataframe(ctx.upper_annotations)

        if not ctx.lower_annotations.empty:
            st.markdown("下部注釈")
            st.dataframe(ctx.lower_annotations)

    st.info("下のボタンを押して機械可読性のチェックを開始してください。")

# チェック実行ボタン
if ctx is not None and "check_done" not in st.session_state:
    if st.button("チェックを実行"):
        with st.spinner("チェック中..."):
            results = []
            progress = st.progress(0, text="LEVEL1 チェック中...")

            levels = ["level1", "level2", "level3"]
            progress_percentages = [0.0, 0.3, 0.6]

            for i, level in enumerate(levels):
                progress.progress(progress_percentages[i], text=f"{level.upper()} チェック中...")

                rule_file = f"rules/{level}.json"
                level_results = run_checks_from_rules(
                    rule_file=rule_file,
                    ctx=ctx,
                    workbook=wb,
                    filepath=st.session_state["uploaded_path"],
                    level=level
                )
                results.append((level, level_results))

            progress.progress(0.9, text="チェック結果の整理と要約生成...")

            summary, summary_md, llm_comment = summarize_results(results)

            progress.progress(1.0, text="全てのチェックが完了しました")

            st.session_state["results"] = results
            st.session_state["summary"] = summary
            st.session_state["summary_md"] = summary_md
            st.session_state["llm_comment"] = llm_comment
            st.session_state["check_done"] = True

# チェック結果の表示とレポート生成
if "results" in st.session_state and "summary" in st.session_state:
    results = st.session_state["results"]
    summary = st.session_state["summary"]
    summary_md = st.session_state.get("summary_md", "")
    llm_comment = st.session_state["llm_comment"]
    uploaded_file = st.session_state.get("uploaded_file", None)
    file_name = uploaded_file.name if uploaded_file is not None else "不明"

    st.markdown(summary_md)
    st.markdown("### 結果概要")
    st.write(llm_comment)

    for level, checks in results:
        with st.expander(f"{level.upper()} チェックの詳細"):
            for item in checks:
                st.markdown(f"**{item['id']} – {item['description']}**")
                st.markdown(f"- 判定: {'合格' if item['result'] == '✓' else '不合格'}")
                st.markdown(f"- 詳細: {item['message']}")
                st.markdown("---")

    report_lines = [
        "# 機械可読性チェックレポート（レベル1〜3）",
        f"ファイル名: {file_name}",
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

    report_filename = f"{Path(file_name).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    report_path = REPORT_DIR / report_filename
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(report_str)
    except Exception as e:
        st.error(f"レポート保存中にエラーが発生しました: {e}")
