# Common utilities
from .common import (
    MAX_EXAMPLES,
    get_excel_column_letter,
    detect_platform_characters,
    is_clean_numeric,
    is_likely_long_format,
    FREE_TEXT_PATTERN,
    MISSING_VALUE_EXPRESSIONS
)

# CSV specific utilities
from .csv_utils import (
    detect_multiple_tables_csv,
    csv_specific_check
)

# XLS specific utilities
from .xls_utils import (
    get_xls_workbook_info,
    check_xls_merged_cells,
    check_xls_cell_formats,
    check_xls_hidden_rows_columns
)

# XLSX specific utilities
from .xlsx_utils import (
    has_any_drawing_xlsx,
    is_sheet_likely_xlsx,
    check_xlsx_format_semantics
)

# Backward compatibility - deprecated, use specific modules instead
def has_any_drawing(path):
    """後方互換性のための関数（推奨されません）"""
    if path.suffix.lower() == ".xlsx":
        return has_any_drawing_xlsx(path)
    elif path.suffix.lower() == ".xls":
        return True  # .xls ファイルは構造上図形チェックが困難
    return False

# Deprecated function for backward compatibility
has_any_drawing_xlsx = has_any_drawing_xlsx 