from pathlib import Path

from .level1_checker import Level1Checker
from .level2_checker import Level2Checker
from .level3_checker import Level3Checker


class CheckerFactory:
    """
    ファイル形式とレベルに応じて適切なチェッカーを生成するファクトリークラス
    """
    
    def __init__(self):
        self.level1_checkers = [
            Level1Checker(),
        ]
        
        self.level2_checkers = [
            Level2Checker(),
        ]
        
        self.level3_checkers = [
            Level3Checker(),
        ]
    
    def get_level1_checker(self, file_path: Path):
        """
        ファイルパスに応じて適切なLevel1チェッカーを返す
        """
        for checker in self.level1_checkers:
            if checker.can_handle(file_path):
                return checker
        raise ValueError(f"Unsupported file format for Level1 checker: {file_path.suffix}")
    
    def get_level2_checker(self, file_path: Path):
        """
        ファイルパスに応じて適切なLevel2チェッカーを返す
        """
        for checker in self.level2_checkers:
            if checker.can_handle(file_path):
                return checker
        raise ValueError(f"Unsupported file format for Level2 checker: {file_path.suffix}")
    
    def get_level3_checker(self, file_path: Path):
        """
        ファイルパスに応じて適切なLevel3チェッカーを返す
        """
        for checker in self.level3_checkers:
            if checker.can_handle(file_path):
                return checker
        raise ValueError(f"Unsupported file format for Level3 checker: {file_path.suffix}")
    
    def get_checker(self, file_path: Path, level: str):
        """
        ファイルパスとレベルに応じて適切なチェッカーを返す
        """
        level = level.lower()
        if level == "level1":
            return self.get_level1_checker(file_path)
        elif level == "level2":
            return self.get_level2_checker(file_path)
        elif level == "level3":
            return self.get_level3_checker(file_path)
        else:
            raise ValueError(f"Unsupported level: {level}")


# ファクトリーインスタンス
checker_factory = CheckerFactory() 