# excel_module/commands.py
"""
Session API for the excel module.

Session is a simple, stateful object for running the pipeline on ONE workbook:
- load(path)
- clean()
- open_excel(visible=True)
- write_to_excel(sheet_name)
- create_dashboard(name, visible, open_in_excel)
- close()
"""

import os
import logging
import traceback
from typing import Optional
import pandas as pd
from .core import load_data, clean_data
from .excel_engine import open_workbook, write_dataframe_to_sheet, create_advanced_dashboard, save_and_close

logger = logging.getLogger("excel_module.commands")
logger.addHandler(logging.NullHandler())


class Session:
    def __init__(self, path: Optional[str] = None):
        self.path = path
        self.df: Optional[pd.DataFrame] = None
        self.app = None
        self.wb = None

    # ---------------- load & clean ----------------
    def load(self, path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """Load CSV or Excel first sheet to memory (DataFrame)."""
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        self.df = load_data(path, sheet_name=sheet_name)
        self.path = path
        logger.info("Loaded data from %s rows=%d cols=%d", path, len(self.df), len(self.df.columns))
        return self.df

    def clean(self, fill_numeric: str = "median") -> pd.DataFrame:
        """Clean loaded DataFrame in memory."""
        if self.df is None:
            raise RuntimeError("No data loaded")
        self.df = clean_data(self.df, fill_numeric=fill_numeric)
        logger.info("Clean completed rows=%d cols=%d", len(self.df), len(self.df.columns))
        return self.df

    # ---------------- workbook lifecycle ----------------
    def open_excel(self, visible: bool = True):
        """
        Ensure a single Excel app + workbook is opened and ready.
        If the workbook does not exist on disk, create a new workbook and SaveAs immediately
        to claim the target filename (avoids later SaveAs prompts).
        """
        if self.app and self.wb:
            return
        self.app, self.wb = open_workbook(self.path, visible=visible)

        # If the path doesn't exist on disk (new), do a SaveAs to the desired path to avoid temp names later
        if self.path and not os.path.exists(self.path):
            try:
                # disable alerts, do SaveAs, re-enable
                try:
                    if hasattr(self.app, "api"):
                        self.app.api.DisplayAlerts = False
                except Exception:
                    pass
                # Use COM SaveAs to ensure workbook is bound to disk path
                self.wb.api.SaveAs(os.path.abspath(self.path))
            except Exception:
                logger.exception("SaveAs on new workbook failed")
            finally:
                try:
                    if hasattr(self.app, "api"):
                        self.app.api.DisplayAlerts = True
                except Exception:
                    pass

    def _safe_save(self) -> bool:
        """
        Try a normal COM save first (with DisplayAlerts turned off to avoid blocking modals).
        If COM save fails, close Excel and do a single pandas.to_excel fallback to preserve data.
        Returns True on success.
        """
        try:
            # disable prompts briefly
            try:
                if self.app and hasattr(self.app, "api"):
                    self.app.api.DisplayAlerts = False
            except Exception:
                pass

            # If workbook is not bound to the path (temp name), call SaveAs explicitly
            try:
                name = getattr(self.wb, "name", "") or ""
                need_saveas = False
                # heuristics: temp names often start with "Book" or "~"
                if self.path and (name.startswith("Book") or name.startswith("~")):
                    need_saveas = True
                if need_saveas:
                    self.wb.api.SaveAs(os.path.abspath(self.path))
                else:
                    self.wb.save(self.path)
            except Exception:
                # final attempt - COM SaveAs
                try:
                    self.wb.api.SaveAs(os.path.abspath(self.path))
                except Exception:
                    raise
            return True

        except Exception as ex:
            logger.warning("COM save failed: %s", ex)
            logger.debug(traceback.format_exc())
            # close Excel to release any locks
            try:
                if self.wb is not None:
                    try:
                        self.wb.close()
                    except Exception:
                        pass
                if self.app is not None:
                    try:
                        self.app.quit()
                    except Exception:
                        pass
            finally:
                self.wb = None
                self.app = None

            # pandas fallback (single controlled write)
            try:
                if self.df is not None and self.path:
                    self.df.to_excel(self.path, index=False)
                    logger.info("Saved via pandas fallback to %s", self.path)
                    return True
                return False
            except Exception:
                logger.exception("Pandas fallback save failed")
                return False
        finally:
            # re-enable alerts if possible
            try:
                if self.app and hasattr(self.app, "api"):
                    self.app.api.DisplayAlerts = True
            except Exception:
                pass

    # ---------------- writing & dashboard ----------------
    def write_to_excel(self, sheet_name: str = "SourceData", visible: bool = True):
        """
        Write the current DataFrame into the open workbook (opens it first if needed).
        This method uses the unified workbook instance so write + dashboard happen in the same file.
        """
        if self.df is None:
            raise RuntimeError("No data loaded")
        self.open_excel(visible=visible)  # ensures workbook is open and bound to path
        write_dataframe_to_sheet(self.wb, sheet_name, self.df)
        if self.path:
            ok = self._safe_save()
            if not ok:
                raise RuntimeError("Failed to save workbook")
        logger.info("Wrote DataFrame to sheet '%s'", sheet_name)
        return sheet_name

    def create_dashboard(self, dashboard_name: str = "AutoDashboard", visible: bool = True, open_in_excel: bool = False):
        """
        Create the advanced dashboard inside the same opened workbook.
        Call open_excel() first via write_to_excel() so everything uses the same workbook instance.
        """
        if self.df is None:
            raise RuntimeError("No data loaded")
        # ensure workbook is open
        self.open_excel(visible=visible)
        name = create_advanced_dashboard(self.wb, self.df, dashboard_name=dashboard_name)
        ok = self._safe_save()
        if not ok:
            raise RuntimeError("Failed to save after creating dashboard")
        if open_in_excel and self.path:
            try:
                os.startfile(self.path)
            except Exception:
                logger.exception("Could not open file via OS")
        logger.info("Dashboard created: %s", name)
        return name

    def close(self):
        """Close workbook and quit Excel app (safely)."""
        if self.app and self.wb:
            try:
                save_and_close(self.app, self.wb, self.path)
            except Exception:
                logger.exception("Error while closing Excel")
        self.app = None
        self.wb = None
