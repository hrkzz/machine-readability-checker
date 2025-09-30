import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import streamlit as st
from openpyxl import load_workbook
import tempfile
from pathlib import Path
from datetime import datetime
from loguru import logger
import pandas as pd # pandas の import を追加

# load_file は load_file_and_extract_context のエイリアス (loader.py の修正に基づく)
# get_sheet_names を import するために loader.py から明示的に import
from src.processor.loader import load_file, get_sheet_names

# table_parser の import はすべて削除

from src.processor.summary import summarize_results
from src.checker.level1_checker import CHECK_FUNCTIONS
import json

# ログファイルの設定
logger.add("logs/app.log", rotation="10 MB", retention="30 days", level="INFO")

# レポートディレクトリの初期化
REPORT_DIR = Path("reports")
if REPORT_DIR.exists():
    for f in REPORT_DIR.iterdir():
        try:
            if f.is_file():
                f.unlink()
        except Exception as e:
            logger.error(f"ファイル {f} の削除に失敗しました: {e}")
else:
    REPORT_DIR.mkdir(parents=True)

st.set_page_config(page_title="機械可読性チェック", layout="wide")

# スタイル適用（任意）
css_path = os.path.join("src", "app", "styles", "style.css")
if os.path.exists(css_path):
    with open(css_path) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title("機械可読性チェックツール")
st.markdown("⚠️ **構造解析はLLMに依存せず、ユーザー定義のシート名とヘッダー範囲に基づいて実行されます。**")

# --- 1. ファイルアップロード ---
uploaded_file = st.file_uploader("CSV または Excel ファイルをアップロード", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    # ファイルが変更されたらセッション状態をリセット
    if st.session_state.get("last_upload_name") != uploaded_file.name:
        st.session_state["uploaded_file"] = uploaded_file
        st.session_state["uploaded_path"] = None
        st.session_state["structure_done"] = False
        st.session_state["check_done"] = False
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=uploaded_file.name) as tmp_file:
            tmp_file.write(uploaded_file.getbuffer())
            st.session_state["uploaded_path"] = tmp_file.name
        
        st.session_state["last_upload_name"] = uploaded_file.name
        st.session_state["sheet_names"] = get_sheet_names(Path(st.session_state["uploaded_path"]))
        st.session_state["selected_sheet"] = st.session_state["sheet_names"][0]
        
    st.info(f"アップロードファイル: {st.session_state['uploaded_file'].name}")
    
    # --- 2. 構造定義の入力 UI ---
    st.markdown("### 📊 データ構造定義")
    col1, col2, col3 = st.columns(3)
    
    # シート名選択
    with col1:
        selected_sheet = st.selectbox(
            "対象シートの選択",
            st.session_state["sheet_names"],
            key="selected_sheet"
        )
    
    # ヘッダー開始行
    with col2:
        header_start_row = st.number_input(
            "表頭（ヘッダー）の**開始行**（1から数える）", 
            min_value=1, 
            value=1, 
            key="header_start_row"
        )
    
    # ヘッダー終了行
    with col3:
        # 終了行は開始行以上であることを保証
        min_end_row = header_start_row if header_start_row else 1
        
        # セッションから前回の値を安全に取得
        previous_end_row = st.session_state.get("header_end_row_default", min_end_row)
        
        # 新しいデフォルト値を決定: 以前の値がmin_end_rowより小さければ、min_end_rowを強制的に使用
        safe_end_row_value = max(previous_end_row, min_end_row)

        header_end_row = st.number_input(
            "表頭（ヘッダー）の**終了行**（1から数える）", 
            min_value=min_end_row, 
            value=safe_end_row_value, # 修正後の安全な値を使用
            key="header_end_row"
        )
        # 終了行の値をセッションに保存（次回再描画時用）
        st.session_state["header_end_row_default"] = header_end_row
    
    # --- 3. 構造解析とチェック実行ボタン ---
    if st.button("構造解析とチェックを実行", key="run_analysis_check"):
        st.session_state["structure_done"] = False
        st.session_state["check_done"] = False
        
        try:
            with st.spinner("構造解析中..."):
                file_path_obj = Path(st.session_state["uploaded_path"])
                file_suffix = file_path_obj.suffix.lower()
                
                # 統合された load_file (load_file_and_extract_context) を呼び出す
                ctx = load_file(
                    file_path_obj, 
                    sheet_name=selected_sheet,
                    header_start_row=header_start_row,
                    header_end_row=header_end_row
                ) 
                
                # ファイル形式に応じてワークブックを読み込み (エラー修正後のロジック)
                if file_suffix == ".xls" or file_suffix == ".csv":
                    # .xlsファイルまたは.csvファイルの場合はワークブックをNoneとして扱う
                    wb = None
                else:
                    # .xlsxなど openpyxl がサポートする形式の場合のみ読み込む
                    wb = load_workbook(st.session_state["uploaded_path"], data_only=True)

                st.session_state["ctx"] = ctx
                st.session_state["workbook"] = wb
                st.session_state["structure_done"] = True
                
            st.success(f"シート「{ctx.sheet_name}」の構造を解析しました。")
            
        except ValueError as ve:
            st.error(f"❌ 構造解析エラー: {ve}")
            st.session_state["structure_done"] = False
            ctx = None
        except Exception as e:
             st.error(f"❌ 予期せぬエラー: {e}")
             logger.exception("予期せぬエラーの詳細:")
             st.session_state["structure_done"] = False
             ctx = None

# ctx / wb の初期化と安全な取得
ctx = None
wb = None
if "ctx" in st.session_state and st.session_state.get("structure_done"):
    ctx = st.session_state["ctx"]
    wb = st.session_state["workbook"]

# --- 4. テーブル構造の表示と自動チェック実行 ---
if ctx is not None and st.session_state.get("structure_done"):
    
    # 構造解析結果の表示
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
            
    # --- チェックの実行 ---
    # `structure_done` が True になった直後、または `check_done` が False の場合に実行
    if st.session_state.get("check_done") is not True and st.session_state.get("structure_done"):
        st.info("機械可読性のチェックを開始します...")
        
        with st.spinner("チェック中..."):
            results = []
            progress = st.progress(0, text="LEVEL1 チェック中...")

            # レベルを level1 のみに限定
            level = "level1"
            rule_file = f"rules/{level}.json"

            try:
                with open(rule_file, encoding="utf-8") as f:
                    rules = json.load(f)
            except FileNotFoundError:
                st.error(f"ルールファイル {rule_file} が見つかりません。")
                st.session_state["check_done"] = True
                st.rerun()

            level_results = []
            total_checks = len(rules)
            for i, rule in enumerate(rules):
                fn_name = rule.get("function")
                fn = CHECK_FUNCTIONS.get(fn_name)

                progress_val = 0.1 + 0.8 * (i / total_checks if total_checks else 1)
                progress.progress(progress_val, text=f"LEVEL1 チェック中: {rule.get('id', '')} - {rule.get('description', '')}...")

                if fn is None:
                    passed, msg = False, f"エラー発生: 対応する関数 '{fn_name}' が level1_checker に見つかりません。"
                else:
                    try:
                        passed, msg = fn(ctx, wb, st.session_state["uploaded_path"])
                    except Exception as e:
                        passed, msg = False, f"実行エラー: {e}"
                        logger.error(f"チェック {rule.get('id', '')} でエラー: {e}")

                level_results.append({
                    "id": rule.get("id", "unknown"),
                    "description": rule.get("description", ""),
                    "result": "✓" if passed else "✗",
                    "message": msg
                })

            results.append((level, level_results))

            # LLMを使用しない簡潔なサマリー生成
            progress.progress(0.9, text="チェック結果の整理...")
            
            summary = {}
            summary_md = "### チェック結果サマリー\n\n"
            table_lines = ["| チェックレベル | 合格数 | 全体数 | 合格率 |", "|----------------|--------|--------|--------|"]

            for level, checks in results:
                total = len(checks)
                passed = sum(1 for item in checks if item["result"] == "✓")
                rate = f"{(passed / total * 100):.0f}%" if total > 0 else "N/A"
                summary[level] = (passed, total)
                table_lines.append(f"| {level.upper()} | {passed} | {total} | {rate} |")
                
            summary_md += "\n".join(table_lines)
            llm_comment = "レベル1のチェックが完了しました。詳細は下の「詳細」セクションを参照してください。（LLMによる総評はスキップ）"

            progress.progress(1.0, text="全てのチェックが完了しました")

            st.session_state["results"] = results
            st.session_state["summary"] = summary
            st.session_state["summary_md"] = summary_md
            st.session_state["llm_comment"] = llm_comment
            st.session_state["check_done"] = True
            st.rerun() # 結果表示のため再実行

# --- 5. チェック結果の表示とレポート生成 ---
if "results" in st.session_state and "summary" in st.session_state and st.session_state.get("check_done"):
    
    st.markdown("---") 
    st.header("✅ 診断結果")
    
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
        if not checks:
            continue
            
        with st.expander(f"**{level.upper()} チェックの詳細**"):
            for item in checks:
                st.markdown(f"**{item['id']} – {item['description']}**")
                st.markdown(f"- 判定: {'**合格**' if item['result'] == '✓' else '**不合格**'}")
                st.markdown(f"- 詳細: {item['message']}")
                st.markdown("---")

    # レポート生成ロジック
    report_lines = [
        "# 機械可読性チェックレポート（レベル1）",
        f"ファイル名: {file_name}",
        "",
        "## 総評",
        llm_comment,
        ""
    ]
    
    if "level1" in summary:
        passed, total = summary["level1"]
        report_lines.append(f"## LEVEL1：{passed}/{total} 合格")

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

def cleanup_files():
    """一時ファイルを削除"""
    for f in ["uploaded_file.xlsx", "uploaded_file.xls", "uploaded_file.csv"]:
        if os.path.exists(f):
            try:
                os.remove(f)
            except Exception as e:
                logger.error(f"ファイル {f} の削除に失敗しました: {e}")