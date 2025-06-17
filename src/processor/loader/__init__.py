from pathlib import Path
from typing import Dict, Any

from .unified_loader import UnifiedLoader


class LoaderFactory:
    """
    ファイル形式に応じて適切なローダーを生成するファクトリークラス
    """
    
    def __init__(self):
        self.loaders = [
            UnifiedLoader(),
        ]
    
    def get_loader(self, file_path: Path):
        """
        ファイルパスに応じて適切なローダーを返す
        """
        for loader in self.loaders:
            if loader.can_handle(file_path):
                return loader
        raise ValueError(f"Unsupported file format: {file_path.suffix}")


# 従来の関数インターフェースを維持
def load_file(file_path: Path) -> Dict[str, Any]:
    """
    ファイルを読み込み、各シートのデータ（先頭・末尾{PREVIEW_ROW_COUNT}行）を返す。
    loader はあくまで「生のテーブル情報」を提供し、オブジェクト検出などは
    checker 側のユーティリティに委ねます。
    """
    factory = LoaderFactory()
    loader = factory.get_loader(file_path)
    return loader.load_file(file_path)


# 対応可能な拡張子
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"} 