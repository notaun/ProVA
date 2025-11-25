# excel_module/excel_engine.py
"""
Excel engine using xlwings (Windows + Excel required).

Provides:
- open_workbook / save_and_close
- write_dataframe_to_sheet
- create_advanced_dashboard: builds a richer dashboard with KPIs, multiple charts, pivot + slicer
"""
from typing import Optional
import pandas as pd
import xlwings as xw
import logging
from .core import detect_column_types
from datetime import datetime

logger = logging.getLogger("excel_module.excel_engine")
logger.addHandler(logging.NullHandler())


def open_workbook(path: Optional[str] = None, visible: bool = True):
    """
    Open Excel app and workbook. If path exists, open it; otherwise create a new workbook.
    Returns (app, workbook).
    """
    app = xw.App(visible=visible)
    if path:
        wb = app.books.open(path)
    else:
        wb = app.books.add()
    return app, wb


def save_and_close(app: xw.App, wb: xw.Book, path: Optional[str] = None):
    """Save workbook (path optional) and quit Excel app."""
    try:
        if path:
            wb.save(path)
        else:
            wb.save()
    finally:
        wb.close()
        app.quit()


def write_dataframe_to_sheet(wb: xw.Book, sheet_name: str, df: pd.DataFrame, start_cell: str = "A1", clear_first: bool = True):
    """
    Write df to sheet_name (create sheet if needed). Returns the xlwings Sheet.
    Uses a simple header + values write for speed and compatibility.
    """
    if sheet_name in [s.name for s in wb.sheets]:
        sht = wb.sheets[sheet_name]
        if clear_first:
            sht.clear()
    else:
        sht = wb.sheets.add(sheet_name)
    sht.range(start_cell).value = [df.columns.tolist()] + df.values.tolist()
    try:
        sht.autofit()
    except Exception:
        pass
    return sht


def _add_native_chart(target_sheet: xw.Sheet, source_sheet: xw.Sheet, source_range: str,
                      chart_type: int = 4, left: int = 10, top: int = 60, width: int = 480, height: int = 300, title: Optional[str] = None):
    """
    Add an Excel native chart on target_sheet. chart_type uses Excel XlChartType indexes.
    """
    try:
        shape = target_sheet.api.Shapes.AddChart2(-1, chart_type, left, top, width, height)
        chart = shape.Chart
        chart.SetSourceData(source_sheet.range(source_range).api)
        if title:
            chart.ChartTitle.Text = title
        return chart
    except Exception:
        logger.exception("chart creation failed")
        return None


def create_advanced_dashboard(wb: xw.Book, df: pd.DataFrame, dashboard_name: str = "AdvancedDashboard",
                              metrics: list | None = None, category_field: str | None = None, date_field: str | None = None, top_n: int = 10):
    """
    Build a dashboard sheet with:
    - KPI summary
    - Time-series trend (monthly)
    - Monthly column chart for second metric
    - Top-N category breakdown
    - Scatter (first two numeric cols)
    - Pivot table + Slicer for category_field
    """
    # main sheet
    if dashboard_name in [s.name for s in wb.sheets]:
        main = wb.sheets[dashboard_name]
        main.clear()
    else:
        main = wb.sheets.add(dashboard_name)
    main.range("A1").value = dashboard_name
    main.range("A1").api.Font.Size = 16
    main.range("A1").api.Font.Bold = True

    df = df.copy()
    types = detect_column_types(df)
    date_cols = types["date_cols"]
    num_cols = types["numeric_cols"]
    cat_cols = types["categorical_cols"]

    if metrics is None:
        metrics = num_cols[:2]
    if category_field is None and cat_cols:
        category_field = cat_cols[0]
    if date_field is None and date_cols:
        date_field = date_cols[0]

    # KPI cards
    kpi_row = 2
    kpi_col = 2
    if metrics:
        m = metrics[0]
        kpis = [("Total", df[m].sum()), ("Average", round(df[m].mean(), 2)), ("Max", df[m].max()), ("Min", df[m].min())]
    else:
        kpis = [("Rows", len(df)), ("Cols", df.shape[1])]

    for i, (lbl, val) in enumerate(kpis):
        r = kpi_row + (i // 2)
        c = kpi_col + (i % 2) * 2
        main.range((r, c)).value = lbl
        main.range((r, c + 1)).value = val
        main.range((r, c)).api.Font.Bold = True

    # write a hidden source sheet for charts/pivots
    src_name = f"{dashboard_name}_src"
    src_sheet = write_dataframe_to_sheet(wb, src_name, df, start_cell="A1", clear_first=True)
    try:
        src_sheet.api.Visible = False
    except Exception:
        pass

    # Trend (monthly) if date and primary metric exist
    left = 10; top = 120
    if date_field and metrics:
        tmp = df.copy()
        tmp[date_field] = pd.to_datetime(tmp[date_field], errors="coerce")
        tmp = tmp.dropna(subset=[date_field])
        tmp["_period"] = tmp[date_field].dt.to_period("M").dt.to_timestamp()
        agg = tmp.groupby("_period")[metrics[0]].sum().reset_index()
        trend_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_trend", agg)
        try: trend_sheet.api.Visible = False
        except: pass
        _add_native_chart(main, trend_sheet, f"A1:B{len(agg)+1}", chart_type=4, left=left, top=top, title=f"{metrics[0]} Trend (Monthly)")
        top += 280

    # Monthly for second metric (if present)
    if date_field and len(metrics) > 1:
        tmp = df.copy()
        tmp[date_field] = pd.to_datetime(tmp[date_field], errors="coerce")
        tmp = tmp.dropna(subset=[date_field])
        tmp["_period"] = tmp[date_field].dt.to_period("M").dt.to_timestamp()
        agg2 = tmp.groupby("_period")[metrics[1]].sum().reset_index()
        mo_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_mo2", agg2)
        try: mo_sheet.api.Visible = False
        except: pass
        _add_native_chart(main, mo_sheet, f"A1:B{len(agg2)+1}", chart_type=51, left=left+540, top=120, title=f"{metrics[1]} Monthly")

    # Category top-N
    if category_field and metrics:
        catagg = df.groupby(category_field)[metrics[0]].sum().reset_index().sort_values(metrics[0], ascending=False).head(top_n)
        cat_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_cat", catagg)
        try: cat_sheet.api.Visible = False
        except: pass
        _add_native_chart(main, cat_sheet, f"A1:B{len(catagg)+1}", chart_type=51, left=10, top=420, width=1000, height=320, title=f"Top {top_n} {category_field}")

    # Scatter of two numeric cols
    if len(num_cols) >= 2:
        xcol, ycol = num_cols[0], num_cols[1]
        sc = df[[xcol, ycol]].dropna()
        sc_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_scatter", sc)
        try: sc_sheet.api.Visible = False
        except: pass
        _add_native_chart(main, sc_sheet, f"A1:B{len(sc)+1}", chart_type=-4169, left=540, top=420, title=f"{ycol} vs {xcol}")

    # Pivot + slicer (best-effort)
    if category_field:
        try:
            p_name = f"{dashboard_name}_pivot"
            p_sheet = wb.sheets[p_name] if p_name in [s.name for s in wb.sheets] else wb.sheets.add(p_name)
            p_sheet.clear()
            src_api = src_sheet.api
            pc = wb.api.PivotCaches().Create(1, src_api.Range("A1").CurrentRegion)
            pt = p_sheet.api.PivotTables().Add(PivotCache=pc, TableDestination=p_sheet.api.Range("A3"), TableName="PT_"+dashboard_name)
            # try set row and data field
            try:
                pt.PivotFields(category_field).Orientation = 1
                pt.AddDataField(pt.PivotFields(metrics[0]), f"Sum of {metrics[0]}", -4157)
            except Exception:
                pass
            # slicer
            try:
                slicer_cache = wb.api.SlicerCaches.Add(pt.PivotFields(category_field))
                slicer_cache.Slicers.Add(p_sheet.api, category_field + "_Slicer", category_field, p_sheet.api.Range("H3"))
            except Exception:
                pass
            try: p_sheet.autofit()
            except: pass
        except Exception:
            logger.exception("Pivot creation failed")

    try:
        main.autofit()
    except Exception:
        pass

    return main.name
