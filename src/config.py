import os
from pathlib import Path

# ベースディレクトリの設定
BASE_DIR = Path(__file__).parent.parent

# 各ディレクトリのパス設定
DATA_DIR = BASE_DIR / "data"
REPORTS_DIR = BASE_DIR / "reports"
RULES_DIR = BASE_DIR / "rules"

# ディレクトリが存在しない場合は作成
for directory in [DATA_DIR, REPORTS_DIR, RULES_DIR]:
    directory.mkdir(exist_ok=True)
