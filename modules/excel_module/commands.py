"""
ProVA -- excel_module/commands.py
Load + clean -> plan -> build dashboard.
"""
from __future__ import annotations
import logging
from pathlib import Path
from typing import Optional
import pandas as pd
from .planner      import plan as make_plan
from .excel_engine import build_dashboard

log = logging.getLogger("ProVA.Excel.Session")
# Project root = ProVA_final_v3/  (two parents up from this file)
_PROJECT_ROOT   = Path(__file__).parent.parent.parent
_DASHBOARDS_DIR = _PROJECT_ROOT / "modules" / "excel_module" / "out"
_CURRENCY_RE = r"[₹$€£¥\s,]"

def _clean_currency_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.select_dtypes(include=["object"]).columns:
        sample = df[col].dropna().astype(str).head(20).tolist()
        has_currency = any(s.strip().startswith(("₹","$","€","£","¥")) for s in sample)
        has_percent  = any(s.strip().endswith("%") for s in sample)
        if not (has_currency or has_percent):
            continue
        cleaned = (df[col].astype(str)
                   .str.replace(_CURRENCY_RE, "", regex=True)
                   .str.replace("%", "", regex=False)
                   .str.strip())
        numeric = pd.to_numeric(cleaned, errors="coerce")
        if numeric.notna().mean() >= 0.70:
            df[col] = numeric
            log.info("Converted '%s' from string to numeric", col)
    return df

def load_data(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {p}")
    suffix = p.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(p, sheet_name=0)
    elif suffix in (".csv", ".txt"):
        df = pd.read_csv(p)
    else:
        raise ValueError(f"Unsupported file type: '{suffix}'.")
    if df.empty:
        raise ValueError(f"File contains no data: {p.name}")
    df.columns = df.columns.astype(str).str.strip().str.replace("\u00a0","",regex=False)
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": pd.NA, "": pd.NA, "None": pd.NA})
    df = _clean_currency_strings(df)
    log.info("Loaded '%s' -- %d rows x %d cols", p.name, len(df), len(df.columns))
    return df

class Session:
    def __init__(self) -> None:
        self.df:          Optional[pd.DataFrame] = None
        self.source_path: Optional[Path]         = None

    def load(self, path: str) -> pd.DataFrame:
        self.df          = load_data(path)
        self.source_path = Path(path)
        return self.df

    def create_dashboard(self, output_path: Optional[str] = None) -> Path:
        if self.df is None:
            raise RuntimeError("No data loaded.")
        stem = self.source_path.stem if self.source_path else "dashboard"
        out  = Path(output_path) if output_path else (
            _DASHBOARDS_DIR / f"{stem}_dashboard.xlsx"
        )
        _DASHBOARDS_DIR.mkdir(parents=True, exist_ok=True)
        return build_dashboard(self.df, make_plan(self.df, source_name=stem), str(out))