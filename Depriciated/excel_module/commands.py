"""
commands.py — Session orchestration for ProVA (V2).

Responsibilities:
- Load a clean dataset into memory
- Hold session state
- Invoke planner to build a DashboardSpec
- Invoke excel_engine to render dashboards

Non-responsibilities:
- Data cleaning
- Column inference
- KPI computation (delegated to planner)
"""

from __future__ import annotations
from pathlib import Path
from typing import Optional, Dict, Any
import pandas as pd

from excel_engine import create_excel_dashboard
from planner import plan_transactional_dashboard


# -----------------------------
# Utilities
# -----------------------------
def load_data(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(str(p))

    if p.suffix.lower() in (".xls", ".xlsx"):
        df = pd.read_excel(p, sheet_name=sheet_name or 0)
    elif p.suffix.lower() in (".csv", ".txt"):
        df = pd.read_csv(p)
    else:
        raise ValueError(f"Unsupported file type: {p.suffix}")

    if df.empty:
        raise ValueError("Loaded DataFrame is empty")

    return df


# -----------------------------
# Session
# -----------------------------
class Session:
    """
    Thin orchestration wrapper.

    This class does NOT:
    - clean data
    - infer structure
    - design dashboards manually
    """

    def __init__(self):
        self.df: Optional[pd.DataFrame] = None

    # -------- Data --------
    def load(self, path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        df = load_data(path, sheet_name)

        # ---- COLUMN NORMALIZATION (CRITICAL) ----
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.replace("\u00A0", "", regex=False)  # Excel non-breaking space
        )

        self.df = df
        return self.df

    # -------- Planning + Dashboard (RECOMMENDED) --------
    def plan_and_create_dashboard(
        self,
        *,
        template_path: Optional[str],
        output_path: str,
        metric: Optional[str] = None,
        date_col: Optional[str] = None,
        category_col: Optional[str] = None,
        visible: bool = False,
    ) -> Path:
        """
        High-level convenience method.

        Flow:
            DataFrame → Planner → DashboardSpec → Excel Engine
        """
        if self.df is None:
            raise RuntimeError("No data loaded. Call load() first.")

        dashboard_spec = plan_transactional_dashboard(
            self.df,
            metric=metric,
            date_col=date_col,
            category_col=category_col,
        )

        return create_excel_dashboard(
            df=self.df,
            dashboard_spec=dashboard_spec,
            template_path=template_path,
            output_path=output_path,
            visible=visible,
        )

    # -------- Dashboard (LOW-LEVEL / ADVANCED) --------
    def create_dashboard(
        self,
        dashboard_spec: Dict[str, Any],
        template_path: Optional[str],
        output_path: str,
        visible: bool = False,
    ) -> Path:
        """
        Low-level entrypoint.

        Use this ONLY if you already have a DashboardSpec
        (e.g. from a custom planner or manual construction).
        """
        if self.df is None:
            raise RuntimeError("No data loaded. Call load() first.")

        return create_excel_dashboard(
            df=self.df,
            dashboard_spec=dashboard_spec,
            template_path=template_path,
            output_path=output_path,
            visible=visible,
        )
