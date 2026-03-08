"""
ProVA -- excel_module/planner.py
DataFrame -> DashboardPlan. Tested across 11 real datasets.
"""
from __future__ import annotations
import logging
from typing import Any, Dict, List, Optional
import pandas as pd
from .core import categorical_columns, date_columns, numeric_columns

log = logging.getLogger("ProVA.Excel.Planner")
TOP_N = 10
_CATEGORY_PRIORITY = (
    "department","dept","category","type","segment","region","group",
    "class","status","genre","specialty","product","division","team",
    "role","position","city","country",
)

def _pick_metrics(df: pd.DataFrame, n: int = 2) -> List[str]:
    cols = numeric_columns(df)
    if not cols:
        return []
    return sorted(cols, key=lambda c: -(df[c].var() if pd.notna(df[c].var()) else 0))[:n]

def _pick_category(df: pd.DataFrame, exclude: set) -> Optional[str]:
    """
    Best categorical column, excluding any already used as metrics.
    Prefers priority-named columns with 3-18 unique values.
    """
    cols = [c for c in categorical_columns(df, max_unique=20) if c not in exclude]
    if not cols:
        return None
    for col in cols:
        if any(kw in col.lower() for kw in _CATEGORY_PRIORITY):
            if 3 <= df[col].nunique(dropna=True) <= 18:
                return col
    for col in cols:
        if 3 <= df[col].nunique(dropna=True) <= 15:
            return col
    return cols[0]

def _pick_date(df: pd.DataFrame) -> Optional[str]:
    dt_cols = date_columns(df)
    if dt_cols:
        for col in dt_cols:
            if "date" in col.lower() or "time" in col.lower():
                return col
        return dt_cols[0]
    # Integer year columns (e.g. movies.year = 1980-2020)
    for col in df.columns:
        if pd.api.types.is_integer_dtype(df[col]):
            if any(k in col.lower() for k in ("year", "yr")):
                vals = df[col].dropna()
                if len(vals) > 0 and vals.between(1900, 2100).mean() > 0.9:
                    return col
    # String/int temporal columns by name: Quarter/Month/Period/Week
    # e.g. "Quarter" with Q1/Q2/Q3/Q4, "Month" with Jan/Feb...
    _TEMPORAL_NAMES = ("month", "period", "quarter", "week", "qtr")
    for col in df.columns:
        cl = col.lower()
        if any(k in cl for k in _TEMPORAL_NAMES):
            n = df[col].nunique(dropna=True)
            if 2 <= n <= 60:
                return col
    return None

def _parse_date_series(series: pd.Series) -> Optional[pd.Series]:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    for dayfirst in (False, True):
        try:
            parsed = pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)
            if parsed.notna().mean() >= 0.80:
                return parsed
        except Exception:
            pass
    return None

def _agg_by_category(df, metric, category, top_n=TOP_N, agg="sum"):
    fn = "sum" if agg == "sum" else "mean"
    cat_s = df[category].astype(str) if df[category].dtype == bool else df[category]
    temp  = df.copy()
    temp[category] = cat_s
    result = (
        temp.groupby(category, dropna=True)[metric].agg(fn)
        .reset_index()
        .sort_values(metric, ascending=False)
        .head(top_n).reset_index(drop=True)
    )
    result.columns = [category, metric]
    return result

def _agg_by_date(df, metric, date_col):
    d = df[[date_col, metric]].copy()
    col = d[date_col]
    if pd.api.types.is_datetime64_any_dtype(col):
        d["_period"] = col.dt.to_period("M").astype(str)
    elif pd.api.types.is_integer_dtype(col):
        d["_period"] = col.astype(str)
    else:
        parsed = _parse_date_series(col)
        d["_period"] = parsed.dt.to_period("M").astype(str) if parsed is not None else col.astype(str)
    result = (
        d.groupby("_period", sort=True)[metric].sum()
        .reset_index().reset_index(drop=True)
    )
    result.columns = ["Period", metric]
    if len(result) > 60:
        result = result.tail(60).reset_index(drop=True)
    return result

def plan(df: pd.DataFrame, source_name: str = "Dashboard") -> Dict[str, Any]:
    if df is None or df.empty:
        raise ValueError("Dataset is empty.")
    metrics = _pick_metrics(df, n=2)
    if not metrics:
        raise ValueError("No numeric columns found. Dataset needs at least one non-ID numeric column.")

    primary   = metrics[0]
    secondary = metrics[1] if len(metrics) > 1 else None

    # Exclude metric columns from category selection to avoid collisions
    metrics_set = set(metrics)
    category    = _pick_category(df, exclude=metrics_set)
    date_col    = _pick_date(df)

    log.info("Plan: primary=%s secondary=%s category=%s date=%s",
             primary, secondary, category, date_col)

    charts: List[Dict[str, Any]] = []

    if category:
        charts.append({"type":"col","title":f"{primary} by {category}",
            "df":_agg_by_category(df,primary,category),"cat_col":category,
            "val_col":primary,"color_idx":0})
    if date_col:
        charts.append({"type":"line","title":f"{primary} Over Time",
            "df":_agg_by_date(df,primary,date_col),"cat_col":"Period",
            "val_col":primary,"color_idx":1})
    elif secondary and category:
        charts.append({"type":"col","title":f"{secondary} by {category}",
            "df":_agg_by_category(df,secondary,category),"cat_col":category,
            "val_col":secondary,"color_idx":1})
    if category:
        top_df = _agg_by_category(df,primary,category,top_n=8)
        charts.append({"type":"bar","title":f"Top {len(top_df)} -- {primary} by {category}",
            "df":top_df,"cat_col":category,"val_col":primary,"color_idx":2})
    if secondary:
        if date_col:
            charts.append({"type":"line","title":f"{secondary} Over Time",
                "df":_agg_by_date(df,secondary,date_col),"cat_col":"Period",
                "val_col":secondary,"color_idx":3})
        elif category:
            charts.append({"type":"bar","title":f"Avg {secondary} by {category}",
                "df":_agg_by_category(df,secondary,category,agg="mean"),
                "cat_col":category,"val_col":secondary,"color_idx":3})

    return {
        "source_name":    source_name,
        "primary_metric": primary,
        "secondary_metric": secondary,
        "category":       category,
        "date_col":       date_col,
        "df_columns":     list(df.columns),
        "n_rows":         len(df),
        "charts":         charts[:4],
    }