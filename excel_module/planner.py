"""
planner.py — Dashboard planning layer for ProVA (V2).

Transactional dashboards only.
"""

from __future__ import annotations
from typing import Dict, Any, List, Optional
import pandas as pd


# -----------------------------
# Errors
# -----------------------------
class PlannerError(ValueError):
    pass


# -----------------------------
# Helpers
# -----------------------------
def _numeric_columns(df: pd.DataFrame) -> List[str]:
    return df.select_dtypes(include=["number"]).columns.tolist()


def _categorical_columns(df: pd.DataFrame, max_unique: int = 20) -> List[str]:
    cols = []
    for c in df.columns:
        if df[c].dtype == object or pd.api.types.is_string_dtype(df[c]):
            n = df[c].nunique(dropna=True)
            if 1 < n <= max_unique:
                cols.append(c)
    return cols


def _date_columns(df: pd.DataFrame) -> List[str]:
    return [
        c for c in df.columns
        if pd.api.types.is_datetime64_any_dtype(df[c])
    ]


# -----------------------------
# Auto chart selection
# -----------------------------
def auto_chart_type(metric_series, category_series):
    if pd.api.types.is_datetime64_any_dtype(category_series):
        return "line"
    if category_series.nunique() <= 6:
        return "column"
    return "bar"


# -----------------------------
# KPI computation
# -----------------------------
def compute_basic_kpis(df: pd.DataFrame, metric: str) -> Dict[str, Any]:
    s = df[metric]

    return {
        "KPI_Records": int(len(df)),
        "KPI_Total": float(s.sum(skipna=True)),
        "KPI_Average": float(s.mean(skipna=True)),
        "KPI_Min": float(s.min(skipna=True)),
        "KPI_Max": float(s.max(skipna=True)),
    }


# -----------------------------
# Planner
# -----------------------------
def plan_transactional_dashboard(
    df: pd.DataFrame,
    *,
    metric: Optional[str] = None,
    date_col: Optional[str] = None,
    category_col: Optional[str] = None,
) -> Dict[str, Any]:

    if df is None or df.empty:
        raise PlannerError("Input DataFrame is empty")

    # -----------------------------
    # Metric
    # -----------------------------
    numeric_cols = _numeric_columns(df)
    if not numeric_cols:
        raise PlannerError("No numeric columns found")

    metric = metric or numeric_cols[0]
    if metric not in df.columns:
        raise PlannerError(f"Metric column '{metric}' not found")

    # -----------------------------
    # Date
    # -----------------------------
    date_cols = _date_columns(df)
    date_col = date_col if date_col in df.columns else (date_cols[0] if date_cols else None)

    # -----------------------------
    # Category
    # -----------------------------
    cat_cols = _categorical_columns(df)
    category_col = category_col if category_col in df.columns else (cat_cols[0] if cat_cols else None)

    # -----------------------------
    # Pivots
    # -----------------------------
    pivots = []

    if date_col:
        pivots.append({
            "name": "MetricByDate",
            "sheet": "Pivot_MetricByDate",
            "rows": [date_col],
            "values": [(metric, "sum")],
            "group_dates": True,
        })

    if category_col:
        pivots.append({
            "name": "MetricByCategory",
            "sheet": "Pivot_MetricByCategory",
            "rows": [category_col],
            "values": [(metric, "sum")],
        })

    pivots.append({
        "name": "MetricDistribution",
        "sheet": "Pivot_MetricDistribution",
        "rows": [],
        "values": [(metric, "count")],
    })

    # -----------------------------
    # Charts (3 max – executive rule)
    # -----------------------------
    charts = []

    if date_col:
        charts.append({
            "pivot": "MetricByDate",
            "type": auto_chart_type(df[metric], df[date_col]),
            "placeholder": "CHART1_TL",
            "title": f"{metric} Trend",
        })

    if category_col:
        charts.append({
            "pivot": "MetricByCategory",
            "type": auto_chart_type(df[metric], df[category_col]),
            "placeholder": "CHART2_TL",
            "title": f"{metric} by {category_col}",
        })

    charts.append({
        "pivot": "MetricDistribution",
        "type": "column",
        "placeholder": "CHART3_TL",
        "title": "Record Count",
    })

    # -----------------------------
    # Slicers (safe only)
    # -----------------------------
    slicers = []
    if category_col:
        slicers.append(category_col)
    if date_col:
        slicers.append(date_col)

    # -----------------------------
    # KPIs
    # -----------------------------
    kpis = compute_basic_kpis(df, metric)

    # -----------------------------
    # Final spec
    # -----------------------------
    return {
        "pattern": "transactional",
        "metric": metric,
        "dimensions": {
            "date": date_col,
            "category": category_col,
        },
        "pivots": pivots,
        "charts": charts,
        "slicers": slicers,
        "kpis": kpis,
    }
