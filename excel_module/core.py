# excel_module/core.py
"""
Data loader, cleaner, and lightweight analysis helpers.

This file is intentionally small and readable:
- load_data: loads a CSV or the first sheet from an Excel workbook
- clean_data: normalizes headers, trims strings, coerces numeric-like text to numbers, fills numeric NA
- detect_column_types: finds date / numeric / categorical columns
- correlation_insights, detect_outliers_zscore, fuzzy_column_match: useful helpers for dashboards
"""
from typing import Optional, Dict, List
import pandas as pd
import numpy as np
import logging
from difflib import get_close_matches
import warnings

logger = logging.getLogger("excel_module.core")
logger.addHandler(logging.NullHandler())


def load_data(path: str, sheet_name: Optional[str] = None, nrows: Optional[int] = None) -> pd.DataFrame:
    """Load CSV or Excel (first sheet by default). Always returns a DataFrame."""
    path = str(path)
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, nrows=nrows)
    if sheet_name is None:
        xls = pd.ExcelFile(path)
        first = xls.sheet_names[0]
        return pd.read_excel(path, sheet_name=first, nrows=nrows)
    return pd.read_excel(path, sheet_name=sheet_name, nrows=nrows)


def _try_parse_dates(series: pd.Series) -> pd.Series:
    """
    Try a few common date formats first (fast). If none parses a majority,
    fall back to pandas' general parser while suppressing the 'could not infer' warning.
    """
    formats = ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%b/%Y", "%Y/%m/%d"]
    for fmt in formats:
        try:
            parsed = pd.to_datetime(series, format=fmt, errors="coerce")
            if parsed.notnull().mean() > 0.6:
                return parsed
        except Exception:
            pass
    # fallback - silence the repetitive warning about inference
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Could not infer format,*")
        return pd.to_datetime(series, errors="coerce")


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Trim header whitespace and replace newlines in column names."""
    df = df.copy()
    df.columns = [str(c).strip().replace("\n", " ").replace("\r", " ") for c in df.columns]
    return df


def clean_data(df: pd.DataFrame, fill_numeric: str = "median") -> pd.DataFrame:
    """
    Clean pipeline:
    - drop empty rows/columns
    - normalize headers
    - trim strings
    - coerce numeric-like strings to numbers
    - fill numeric NA with median/mean/zero
    """
    df = df.copy()
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = normalize_headers(df)

    # trim strings
    obj_cols = df.select_dtypes(include=["object"]).columns.tolist()
    for c in obj_cols:
        df[c] = df[c].apply(lambda v: v.strip() if isinstance(v, str) else v)

    # try to coerce numeric-like columns
    for c in df.columns:
        if df[c].dtype == object:
            cleaned = pd.to_numeric(df[c].astype(str).str.replace(r"[,$]", "", regex=True), errors="coerce")
            if cleaned.notnull().mean() > 0.5:
                logger.debug("Coercing column '%s' to numeric", c)
                df[c] = cleaned

    # fill numeric missing values
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    for c in num_cols:
        if fill_numeric == "median":
            fill_val = df[c].median()
        elif fill_numeric == "mean":
            fill_val = df[c].mean()
        else:
            fill_val = 0
        df[c] = df[c].fillna(fill_val)

    return df


def detect_column_types(df: pd.DataFrame) -> Dict[str, List[str]]:
    """
    Detects date, numeric, and categorical columns.
    Returns: {"date_cols": [...], "numeric_cols": [...], "categorical_cols": [...]}
    """
    date_cols, numeric_cols, categorical_cols = [], [], []
    for col in df.columns:
        ser = df[col]
        if pd.api.types.is_numeric_dtype(ser):
            numeric_cols.append(col)
            continue
        if pd.api.types.is_datetime64_any_dtype(ser) or pd.api.types.is_period_dtype(ser):
            date_cols.append(col)
            continue
        # try parse as date with helper
        parsed = _try_parse_dates(ser)
        if parsed.notnull().mean() >= 0.6:
            date_cols.append(col)
            continue
        categorical_cols.append(col)
    logger.debug("columns detected: dates=%s numeric=%s cats=%s", date_cols, numeric_cols, categorical_cols)
    return {"date_cols": date_cols, "numeric_cols": numeric_cols, "categorical_cols": categorical_cols}


def correlation_insights(df: pd.DataFrame, top_n: int = 5):
    """Return correlation matrix and top correlated pairs among numeric columns."""
    num = df.select_dtypes(include=["number"])
    if num.shape[1] < 2:
        return {"matrix": pd.DataFrame(), "top_pairs": []}
    corr = num.corr()
    pairs = []
    cols = corr.columns.tolist()
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            pairs.append((cols[i], cols[j], float(corr.iloc[i, j])))
    pairs_sorted = sorted(pairs, key=lambda x: abs(x[2]), reverse=True)
    return {"matrix": corr, "top_pairs": pairs_sorted[:top_n]}


def detect_outliers_zscore(df: pd.DataFrame, column: str, z_thresh: float = 3.0):
    """Boolean Series marking outliers by z-score on a numeric column."""
    ser = pd.to_numeric(df[column], errors="coerce")
    mean = ser.mean()
    std = ser.std(ddof=0) if ser.std(ddof=0) != 0 else 1.0
    z = (ser - mean) / std
    return z.abs() > z_thresh


def fuzzy_column_match(query: str, columns: List[str], n: int = 1, cutoff: float = 0.6) -> List[str]:
    """Find best fuzzy matches for a query among column names."""
    if not query or not columns:
        return []
    return get_close_matches(query, columns, n=n, cutoff=cutoff)
