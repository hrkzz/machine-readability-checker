from dataclasses import dataclass
import pandas as pd
from typing import Any, Dict


@dataclass
class TableContext:
    sheet_name: str
    data: pd.DataFrame
    columns: Any  # List[str] or pandas.MultiIndex
    upper_annotations: pd.DataFrame
    lower_annotations: pd.DataFrame
    row_indices: Dict[str, int]
