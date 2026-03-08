"""
ProVA -- excel_module/core.py
Column classification with ID exclusion, string date detection, bool categories.
"""
from __future__ import annotations
from typing import List
import logging
import pandas as pd

log = logging.getLogger("ProVA.Excel.Core")

class SchemaError(ValueError):
    pass

_ID_EXACT = frozenset([
    "id", "zip", "postal_code", "postal code", "postcode",
    "phone", "fax", "ssn", "number",
    "row_id", "row id", "order id", "order_id", "emp id", "emp_id",
])
_ID_SUFFIXES = ("id", "_id", " id", "hipe", "code", "num", "no", "key")
_ID_PREFIXES = ("row_", "row ")

def _is_id_column(name: str, series: pd.Series) -> bool:
    cl = name.lower().strip()
    if cl in _ID_EXACT:
        return True
    if any(cl.endswith(s) for s in _ID_SUFFIXES):
        return True
    if any(cl.startswith(s) for s in _ID_PREFIXES):
        return True
    if (pd.api.types.is_integer_dtype(series)
            and series.nunique() == len(series.dropna())):
        return True
    return False

def assert_columns(df: pd.DataFrame, required: List[str]) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SchemaError(f"Missing required column(s): {missing}.")

def numeric_columns(df: pd.DataFrame) -> List[str]:
    return [
        c for c in df.select_dtypes(include=["number"]).columns
        if not _is_id_column(c, df[c])
    ]

def categorical_columns(df: pd.DataFrame, max_unique: int = 20) -> List[str]:
    cols = []
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]):
            n = df[c].nunique(dropna=True)
            if 2 <= n <= max_unique:
                cols.append(c)
        elif df[c].dtype == bool:
            cols.append(c)
        elif pd.api.types.is_integer_dtype(df[c]):
            n = df[c].nunique(dropna=True)
            if 2 <= n <= 15 and not _is_id_column(c, df[c]):
                cols.append(c)
    return cols

def _try_parse_date(series: pd.Series, dayfirst: bool) -> float:
    try:
        parsed = pd.to_datetime(series.dropna().head(40), errors="coerce", dayfirst=dayfirst)
        return float(parsed.notna().mean())
    except Exception:
        return 0.0

def date_columns(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            cols.append(c)
            continue
        if pd.api.types.is_string_dtype(df[c]):
            score = max(_try_parse_date(df[c], True), _try_parse_date(df[c], False))
            if score >= 0.80:
                cols.append(c)
    return cols