# 機械可読性診断ツール

Excel/CSVファイルの機械可読性を診断するためのWebアプリケーションです。データの品質と再利用性を向上させるためのチェックリストに基づいて、ファイルの診断を行います。

## 機能

- Excel/CSVファイルのアップロードと診断
- レベル1の診断項目（10項目）のチェック
- 診断結果の詳細表示
- レポートの生成と保存

### 診断項目
(データの機械可読性を高めるための新しいルールのイメージ)<https://www.cas.go.jp/jp/seisaku/digital_gyozaikaikaku/data8/data8_siryou1.pdf>の7ページ

## 必要条件

- Python 3.8以上
- 必要なパッケージ（requirements.txtに記載）
- OpenAI APIキー（.envファイルで設定）

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

3. 環境変数の設定
`.env`ファイルを作成し、以下の内容を設定：
```
OPENAI_API_KEY=your_api_key_here
```

## 使用方法

1. アプリケーションの起動
```bash
make run
```

2. 診断したいExcel/CSVファイルをアップロード

3. 診断結果の確認
   - サマリー表示
   - 詳細な診断結果
   - レポートのダウンロード

## プロジェクト構造

```
machine_readability_checker/
├── src/
│   ├── app.py              # Streamlitアプリケーション
│   ├── config.py           # 設定ファイル
│   ├── checker/            # チェック機能
│   │   ├── level1_checks.py
│   │   └── utils.py
│   ├── processor/          # データ処理
│   │   └── loader.py
│   └── llm/               # LLM関連
│       └── column_meaning.py
├── rules/                 # ルール定義
│   └── level1.json
├── data/                  # 一時ファイル保存
├── reports/              # レポート出力
├── requirements.txt      # 依存パッケージ
└── README.md            # 本ファイル
```

