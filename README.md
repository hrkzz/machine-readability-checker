# 機械可読性診断ツール

Excel/CSVファイルの機械可読性を診断するためのWebアプリケーションです。データの品質と再利用性を向上させるためのチェックリストに基づいて、ファイルの診断を行います。

## 機能

- Excel/CSVファイルのアップロード
- テーブルデータの構造解析
- 機械可読性診断
- レポートの生成

### 診断項目
以下の2つのルールの順守度合いを診断
- [統計表における機械判読可能なデータの表記方法の統一ルール](https://www.soumu.go.jp/menu_news/s-news/01toukatsu01_02000186.html)
- [官民におけるデータの利活用について（内閣官房）](https://www.cas.go.jp/jp/seisaku/digital_gyozaikaikaku/data8/data8_siryou1.pdf)の7ページにあるデータの機械可読性を高めるための新しいルールのイメージ

## 必要条件

- Python 3.8以上
- 必要なパッケージ（requirements.txtに記載）

## インストール方法

1. リポジトリのクローン
```bash
git clone [リポジトリURL]
cd machine_readability_checker
```

2. 必要なパッケージのインストール
```bash
pip install -r requirements.txt
```



## 使用方法

1. アプリケーションの起動
```bash
make run
```

2. 診断したいExcel/CSVファイルをアップロード
　※xlsx、xls、csvファイルに対応

3. 診断結果の確認

## プロジェクト構造

```
machine_readability_checker/
├── Makefile                   
├── requirements.txt           
├── README.md                  
├── data/                      # アップロードデータや一時ファイル保存先
├── reports/                   # レポート出力ディレクトリ
├── rules/                     # 各チェックレベルのルール定義
│   ├── level1.json
│   ├── level2.json
│   └── level3.json
├── src/
│   ├── app/
│   │   ├── styles/            # スタイル（CSS等）を格納
│   │   └── app.py             # Streamlit アプリケーションUI
│   ├── checker/               # 機械可読性チェック機能のロジック群
│   │   ├── level1_checker.py  # レベル1チェックの実装
│   │   ├── level2_checker.py  # レベル2チェックの実装
│   │   ├── level3_checker.py  # レベル3チェックの実装
│   │   ├── router.py          # チェック処理のルーティング管理
│   │   └── utils.py           # 共通ユーティリティ関数
│   ├── processor/             # データ処理・パース用モジュール
│   │   ├── context.py         # 文脈情報の処理
│   │   ├── loader.py          # データ読み込み処理
│   │   └── table_parser.py    # テーブル構造の解析処理
│   ├── config.py              # アプリケーション設定
│   └── summary.py             # チェック結果の要約処理
```

