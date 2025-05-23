import streamlit as st
import json
import datetime
from processor.loader import load_file
from checker.level1_checks import (
    check_no_merged_cells,
    check_single_data_per_cell,
    check_no_hidden_rows_columns,
    check_clear_column_names,
    check_no_empty_rows_columns
)
from config import DATA_DIR, REPORTS_DIR, RULES_DIR

# ページ設定
st.set_page_config(
    page_title="機械可読性診断ツール",
    page_icon="📊",
    layout="wide"
)

# タイトル
st.title("機械可読性診断ツール")
st.markdown("ExcelまたはCSVファイルの機械可読性を診断します。")

# ルールの読み込み
with open(RULES_DIR / "level1.json", "r", encoding="utf-8") as f:
    rules = json.load(f)

# ファイルアップロード
uploaded_file = st.file_uploader(
    "ファイルをアップロードしてください",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    # ファイルの保存
    file_path = DATA_DIR / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    try:
        # ファイルの読み込み
        df, metadata = load_file(file_path)
        
        # 診断結果の初期化
        results = []
        
        # 各ルールのチェック
        for rule in rules["rules"]:
            check_result = None
            
            if rule["id"] == "no_merged_cells":
                check_result = check_no_merged_cells(metadata.get("worksheet"))
            elif rule["id"] == "single_data_per_cell":
                check_result = check_single_data_per_cell(df)
            elif rule["id"] == "no_hidden_rows_columns":
                check_result = check_no_hidden_rows_columns(metadata.get("worksheet"))
            elif rule["id"] == "clear_column_names":
                check_result = check_clear_column_names(df)
            elif rule["id"] == "no_empty_rows_columns":
                check_result = check_no_empty_rows_columns(df)
            
            if check_result:
                results.append({
                    "rule": rule,
                    "result": check_result
                })
        
        # 結果の表示
        st.header("診断結果")
        
        # サマリー
        passed_count = sum(1 for r in results if r["result"]["passed"])
        total_count = len(results)
        
        st.metric(
            "合格項目",
            f"{passed_count}/{total_count}",
            f"{passed_count/total_count*100:.1f}%"
        )
        
        # 詳細結果
        for result in results:
            rule = result["rule"]
            check_result = result["result"]
            
            with st.expander(f"{'✅' if check_result['passed'] else '❌'} {rule['name']}"):
                st.markdown(f"**説明**: {rule['description']}")
                st.markdown(f"**重要度**: {rule['severity']}")
                
                if not check_result["passed"]:
                    st.markdown(f"**推奨対応**: {rule['recommendation']}")
                    
                    if "details" in check_result:
                        st.markdown("**詳細**:")
                        for key, value in check_result["details"].items():
                            st.markdown(f"- {key}: {value}")
        
        # レポートの保存
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = REPORTS_DIR / f"report_{timestamp}.md"
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("# 機械可読性診断レポート\n\n")
            f.write(f"診断日時: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"ファイル名: {uploaded_file.name}\n\n")
            f.write("## サマリー\n")
            f.write(f"- 合格項目: {passed_count}/{total_count} ({passed_count/total_count*100:.1f}%)\n\n")
            f.write("## 詳細結果\n")
            
            for result in results:
                rule = result["rule"]
                check_result = result["result"]
                
                f.write(f"### {rule['name']}\n")
                f.write(f"- 結果: {'✅ 合格' if check_result['passed'] else '❌ 不合格'}\n")
                f.write(f"- 説明: {rule['description']}\n")
                f.write(f"- 重要度: {rule['severity']}\n")
                
                if not check_result["passed"]:
                    f.write(f"- 推奨対応: {rule['recommendation']}\n")
                    
                    if "details" in check_result:
                        f.write("- 詳細:\n")
                        for key, value in check_result["details"].items():
                            f.write(f"  - {key}: {value}\n")
                
                f.write("\n")
        
        st.success(f"レポートを保存しました: {report_path}")
        
    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
    
    finally:
        # 一時ファイルの削除
        if file_path.exists():
            file_path.unlink() 