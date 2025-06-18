# 機械可読性診断ツール

Excel/CSVファイルの機械可読性を診断するためのWebアプリケーションです。データの品質と再利用性を向上させるためのチェックリストに基づいて、3段階のレベルでファイルの診断を行います。

## 特徴

- **マルチフォーマット対応**: Excel (.xlsx/.xls) および CSV ファイルの診断
- **3段階チェック**: レベル1〜3での段階的な機械可読性診断
- **統一されたアーキテクチャ**: レベル別チェッカーとフォーマット固有ハンドラーの組み合わせ
- **包括的エラーハンドリング**: 堅牢なエラー処理とログ機能
- **モダンなアーキテクチャ**: ファクトリーパターンによる拡張性の高い設計

## 機能概要

### 診断レベル

1. **レベル1**: 基本的なフォーマット整合性チェック
   - ファイル形式の妥当性、画像・オブジェクトの存在、結合セル、書式による意味付けなど
2. **レベル2**: データ構造と値の妥当性チェック
   - 数値列の妥当性、選択肢列の分離、列ヘッダーの明確性、欠損値の統一性など
3. **レベル3**: メタデータとドキュメント品質チェック
   - コード化、コード表の存在、設問マスター、メタデータの完備性など

### 対応ファイル形式

- **CSV**: カンマ区切りファイル（エンコーディング自動検出対応）
- **XLS**: Excel 97-2003 形式（一部機能制限あり）
- **XLSX**: Excel 2007以降の形式（全機能対応）

### 診断基準

以下の公的ガイドラインに基づいた診断を実施：
- [統計表における機械判読可能なデータの表記方法の統一ルール](https://www.soumu.go.jp/menu_news/s-news/01toukatsu01_02000186.html)
- [官民におけるデータの利活用について（内閣官房）](https://www.cas.go.jp/jp/seisaku/digital_gyozaikaikaku/data8/data8_siryou1.pdf)の7ページ：データの機械可読性を高めるための新しいルール

## 技術仕様

### 必要条件

- Python 3.8以上
- 必要なパッケージ（requirements.txtに記載）
- OpenAI APIキー（.streamlit/secrets.tomlファイルで設定）

### 主要依存関係

- `streamlit==1.32.0`: Webアプリケーションフレームワーク
- `pandas==2.2.1`: データ分析・処理
- `openpyxl==3.1.2`: Excel .xlsx ファイル処理
- `xlrd==1.2.0`: Excel .xls ファイル処理
- `loguru==0.7.2`: 高機能ログ管理
- `litellm==1.61.8`: LLM統合ライブラリ
- `chardet==5.2.0`: エンコーディング自動検出

## インストール・セットアップ

### 1. リポジトリのクローン

```bash
git clone [リポジトリURL]
cd machine_readability_checker
```

### 2. 依存関係のインストール

```bash
make install
```

または

```bash
pip install -r requirements.txt
```

### 3. 環境設定

`.streamlit/secrets.toml` ファイルを作成し、以下を記述：

```toml
OPENAI_API_KEY = "your_api_key_here"
```

## 使用方法

### アプリケーションの起動

```bash
make run
```

または

```bash
streamlit run src/app/app.py
```

### 診断の実行

1. Webブラウザでアプリケーションにアクセス
2. 診断したいファイル（.xlsx/.xls/.csv）をアップロード
3. 構造解析ボタンをクリックして、テーブル構造を解析
4. チェック実行ボタンをクリックして、3段階の診断を実行
5. 生成されたレポートを確認・ダウンロード

## プロジェクト構造

```
machine_readability_checker/
├── Makefile                   # ビルド・実行コマンド
├── requirements.txt           # Python依存関係
├── README.md                  # プロジェクト説明書
├── data/                      # アップロードファイル保存先
├── reports/                   # 診断レポート出力先
├── logs/                      # アプリケーションログ
├── rules/                     # 診断ルール定義（JSON）
│   ├── level1.json           # レベル1診断ルール
│   ├── level2.json           # レベル2診断ルール
│   └── level3.json           # レベル3診断ルール
├── .streamlit/               # Streamlit設定
│   └── secrets.toml          # API キー等の機密情報
└── src/                      # ソースコード
    ├── config.py             # アプリケーション設定
    ├── app/                  # Webアプリケーション
    │   ├── app.py           # メインアプリケーション
    │   └── styles/          # UI スタイル（CSS等）
    ├── checker/              # 機械可読性診断エンジン
    │   ├── factory.py       # チェッカーファクトリー
    │   ├── router.py        # 診断ルーティング
    │   ├── base_checker.py  # 抽象基底クラス
    │   ├── common.py        # 共通ユーティリティ
    │   ├── level1_checker.py # レベル1チェッカー
    │   ├── level2_checker.py # レベル2チェッカー
    │   ├── level3_checker.py # レベル3チェッカー
    │   └── handler/         # フォーマット固有ハンドラー
    │       ├── format_handler.py # 統合フォーマットハンドラー
    │       ├── csv_handler.py    # CSV専用処理
    │       ├── xls_handler.py    # XLS専用処理
    │       └── xlsx_handler.py   # XLSX専用処理
    ├── processor/            # データ処理エンジン
    │   ├── loader/          # ファイル読み込み
    │   │   ├── base.py      # 基底ローダー
    │   │   └── file_loader.py # ファイル読み込み実装
    │   ├── loader.py        # メインローダー
    │   ├── table_parser.py  # テーブル構造解析
    │   ├── context.py       # コンテキスト管理
    │   └── summary.py       # 結果要約処理
    └── llm/                  # LLM（大規模言語モデル）連携
        └── llm_client.py    # LLM クライアント
```

## アーキテクチャ

### 設計思想

本プロジェクトは以下の設計原則に基づいて構築されています：

1. **レイヤー分離**: チェッカー層、ハンドラー層、プロセッサー層の明確な分離
2. **ファクトリーパターン**: レベルに応じた適切なチェッカーの動的生成
3. **委譲パターン**: フォーマット固有処理の専用ハンドラーへの委譲
4. **拡張性**: 新しいレベルやファイル形式の追加が容易
5. **堅牢性**: 包括的なエラーハンドリングとログ管理

### 主要コンポーネント

- **Factory**: `src/checker/factory.py` - レベル別チェッカーの生成と管理
- **Router**: `src/checker/router.py` - 診断処理のルーティング
- **Level Checkers**: `level1_checker.py`, `level2_checker.py`, `level3_checker.py` - レベル別診断ロジック
- **Format Handlers**: `src/checker/handler/` - ファイル形式固有の処理を実装
- **Processor**: `src/processor/` - ファイル読み込みとテーブル構造解析
- **Logging**: Loguru による統一されたログ管理

### 処理フロー

1. **ファイル読み込み**: `processor/loader.py` でファイル形式を判定し、適切なローダーで読み込み
2. **構造解析**: `processor/table_parser.py` でテーブル構造を解析
3. **チェッカー選択**: `factory.py` でレベルに応じたチェッカーを選択
4. **診断実行**: 各チェッカーがフォーマットハンドラーに処理を委譲して診断実行
5. **結果統合**: `processor/summary.py` で結果を統合し、LLMによる総評を生成

## 開発・貢献

### コマンド

```bash
# 依存関係のインストール
make install

# アプリケーションの実行
make run

# コードのリンター実行
make lint

# コードのフォーマット
make format

# コードチェック（リンター + フォーマット確認）
make check
```

### ブランチ戦略

- `main`: 安定版
- `feature/*`: 新機能開発
- `fix/*`: バグ修正

## トラブルシューティング

### よくある問題

1. **CSVファイルの文字化け**: 複数エンコーディングでの自動検出を実装済み（UTF-8, CP932, Shift_JIS等）
2. **XLSファイルの機能制限**: 図形・書式の一部チェックが制限される旨を明示
3. **メモリ不足**: プレビュー行数を制限し、大ファイルに対応

### ログの確認

詳細なエラー情報は `logs/` ディレクトリのログファイルで確認できます：
- `logs/app.log`: アプリケーション全体のログ
- `logs/level1_checker.log`: レベル1診断の詳細ログ

## ライセンス

[ライセンス情報を記載]

## 更新履歴

- **v2.0.0**: アーキテクチャの全面リファクタリング
  - レベル別チェッカー + フォーマット固有ハンドラーの実装
  - エンコーディング自動検出機能の追加
  - 処理フローの最適化とエラーハンドリングの強化

