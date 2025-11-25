# excel_module/commands.py
"""
Simple command/session layer for the excel module.

Session exposes:
- load(path)
- clean()
- write_to_excel(sheet_name, visible)
- create_dashboard(dashboard_name, visible, open_in_excel)  # advanced dashboard
- close()
"""

import time
import os
import logging
import traceback
from typing import Optional
import pandas as pd

from .core import load_data, clean_data, fuzzy_column_match
from .excel_engine import open_workbook, write_dataframe_to_sheet, create_advanced_dashboard, save_and_close

logger = logging.getLogger("excel_module.commands")
logger.addHandler(logging.NullHandler())


class Session:
    """Keep state while working on a single workbook / dataframe."""

    def __init__(self, path: Optional[str] = None):
        self.path = path
        self.df: Optional[pd.DataFrame] = None
        self.app = None
        self.wb = None

    # ---------- loading & cleaning ----------
    def load(self, path: str, sheet_name: Optional[str] = None):
        """Load data from CSV or Excel first sheet by default."""
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        self.df = load_data(path, sheet_name=sheet_name)
        self.path = path
        logger.info("Loaded %s (rows=%d cols=%d)", path, len(self.df), len(self.df.columns))
        return self.df

    def clean(self, fill_numeric: str = "median"):
        """Run the clean pipeline (returns cleaned df)."""
        if self.df is None:
            raise RuntimeError("No data loaded")
        self.df = clean_data(self.df, fill_numeric=fill_numeric)
        logger.info("Cleaned data (rows=%d cols=%d)", len(self.df), len(self.df.columns))
        return self.df

    # ---------- Excel IO & safe save ----------
    def _open_if_needed(self, visible: bool):
        if self.app is None or self.wb is None:
            self.app, self.wb = open_workbook(self.path, visible=visible)

    def _safe_save(self) -> bool:
        """
        Try to save via Excel COM. If that fails (modal dialogs, RPC), close Excel
        and save via pandas.to_excel as a fallback. Returns True on success.
        """
        try:
            # temporarily disable prompts to avoid modal SaveAs popups
            try:
                if self.app and hasattr(self.app, "api"):
                    self.app.api.DisplayAlerts = False
            except Exception:
                pass

            # attempt normal save
            self.wb.save(self.path)
            return True

        except Exception as ex:
            logger.warning("COM save failed, falling back to pandas: %s", ex)
            logger.debug(traceback.format_exc())
            # try to close Excel gracefully to release file locks
            try:
                if self.wb is not None:
                    try: self.wb.close()
                    except: pass
                if self.app is not None:
                    try: self.app.quit()
                    except: pass
            finally:
                self.wb = None
                self.app = None

            # fallback: use pandas to save the DataFrame (this will not preserve COM-created charts)
            try:
                if self.df is not None:
                    self.df.to_excel(self.path, index=False)
                    logger.info("Saved via pandas fallback to %s", self.path)
                    return True
            except Exception:
                logger.exception("Pandas fallback save failed")
                return False
        finally:
            # re-enable alerts if app is still present
            try:
                if self.app and hasattr(self.app, "api"):
                    self.app.api.DisplayAlerts = True
            except Exception:
                pass

    def write_to_excel(self, sheet_name: str = "SourceData", visible: bool = True):
        """Write the cleaned DataFrame to a sheet in the workbook and save."""
        if self.df is None:
            raise RuntimeError("No data loaded")
        self._open_if_needed(visible=visible)
        write_dataframe_to_sheet(self.wb, sheet_name, self.df)
        if self.path:
            ok = self._safe_save()
            if not ok:
                raise RuntimeError("Failed to save workbook")
        logger.info("Wrote DataFrame to sheet %s", sheet_name)
        return sheet_name

    # ---------- dashboards ----------
    def create_dashboard(self, dashboard_name: str = "AutoDashboard", visible: bool = True, open_in_excel: bool = False):
        """Create an advanced dashboard sheet (wrapper around excel_engine.create_advanced_dashboard)."""
        if self.df is None:
            raise RuntimeError("No data loaded")
        self._open_if_needed(visible=visible)
        name = create_advanced_dashboard(self.wb, self.df, dashboard_name=dashboard_name)
        ok = self._safe_save()
        if not ok:
            raise RuntimeError("Failed to save after creating dashboard")
        if open_in_excel and self.path:
            try:
                os.startfile(self.path)
            except Exception:
                logger.exception("Could not open file via OS")
        return name

    # ---------- cleanup ----------
    def close(self):
        """Close workbook and Excel app if open."""
        if self.app and self.wb:
            try:
                save_and_close(self.app, self.wb, self.path)
            except Exception:
                logger.exception("Error during save_and_close")
        self.app = None
        self.wb = None
