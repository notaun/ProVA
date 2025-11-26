# excel_module/excel_engine.py
"""
Excel engine using xlwings. Requires Windows + Excel installed.

Provides:
- open_workbook(): start Excel and open or create workbook
- write_dataframe_to_sheet(): write DataFrame into workbook
- create_advanced_dashboard(): builds KPI cards, charts, pivot & slicer
- save_and_close(): save and quit Excel
"""
from typing import Optional
import pandas as pd
import xlwings as xw
import logging
import os
from .core import detect_column_types

logger = logging.getLogger("excel_module.excel_engine")
logger.addHandler(logging.NullHandler())


def open_workbook(path: Optional[str] = None, visible: bool = True):
    """
    Start Excel (xlwings App) and open existing workbook or create new.
    Returns (app, wb).
    """
    app = xw.App(visible=visible)
    if path and os.path.exists(path):
        wb = app.books.open(path)
    else:
        wb = app.books.add()
    return app, wb


def save_and_close(app: xw.App, wb: xw.Book, path: Optional[str] = None):
    """
    Save workbook (to path if provided) and close Excel app.
    This is a straightforward close used when Session.close() calls it.
    """
    try:
        if path:
            wb.save(path)
        else:
            wb.save()
    finally:
        try:
            wb.close()
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass


def write_dataframe_to_sheet(wb: xw.Book, sheet_name: str, df: pd.DataFrame, start_cell: str = "A1", clear_first: bool = True):
    """
    Write DataFrame to sheet. Creates the sheet if missing.
    Returns the xlwings Sheet object.
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
    Add an Excel native chart (COM) to target_sheet using source_range on source_sheet.
    chart_type corresponds to Excel XlChartType integer (4=line, 51=col clustered, -4169=scatter).
    """
    try:
        shape = target_sheet.api.Shapes.AddChart2(-1, chart_type, left, top, width, height)
        chart = shape.Chart
        chart.SetSourceData(source_sheet.range(source_range).api)
        if title:
            try:
                chart.ChartTitle.Text = title
            except Exception:
                pass
        # basic formatting: legend off for single-series, gridlines off
        try:
            chart.HasLegend = False
        except Exception:
            pass
        return chart
    except Exception:
        logger.exception("Failed to add native chart")
        return None


def _rgb(r, g, b):
    """
    Convert RGB to Excel color integer.
    In COM the color is B + G*256 + R*65536 or vice-versa depending on host;
    using (r + g*256 + b*65536) is safe across many setups, but we wrap in try/except.
    """
    try:
        return int(r) + int(g) * 256 + int(b) * 65536
    except Exception:
        return None


def _format_kpi_cell(sheet, row, col, label, value, currency=False):
    """
    Write a KPI label and value, apply nice formatting (big bold value, subtle fill).
    row,col specify the label cell; value placed in (row, col+1)
    """
    try:
        sheet.range((row, col)).value = label
        sheet.range((row, col + 1)).value = value
        # formatting
        try:
            # label style
            sheet.range((row, col)).api.Font.Bold = True
            sheet.range((row, col)).api.Font.Size = 10
            # value style
            vcell = sheet.range((row, col + 1))
            vcell.api.Font.Bold = True
            vcell.api.Font.Size = 18
            # number format
            if currency:
                vcell.number_format = '"₹"#,##0' if False else '#,##0.00'  # you can swap to currency symbol you prefer
            else:
                vcell.number_format = '#,##0'
            # cell fill / border
            try:
                color = _rgb(245, 245, 245)
                vcell.api.Interior.Color = color
                vcell.api.Borders.Weight = 2
            except Exception:
                pass
        except Exception:
            pass
    except Exception:
        pass


def _style_chart_object(xl_chart):
    """
    Apply a few safe, modern chart tweaks:
    - ChartStyle (built-in Excel style index)
    - hide legend for single-series charts
    - set axis tick label font size
    - enable data labels for column charts
    """
    try:
        # try a built-in style that looks modern; if it fails it's non-fatal
        try:
            xl_chart.ChartStyle = 201  # 201 is a fairly clean modern style on many Excel versions
        except Exception:
            pass

        # attempt series formatting (first series)
        try:
            series = xl_chart.SeriesCollection(1)
            # For column/area apply data labels (value)
            try:
                series.HasDataLabels = True
                series.DataLabels().ShowValue = True
            except Exception:
                pass
            # reduce marker size for scatter/line
            try:
                series.MarkerSize = 6
            except Exception:
                pass
        except Exception:
            pass

        # Axis label sizes
        try:
            xl_chart.Axes(1).TickLabels.Font.Size = 9  # category axis
            xl_chart.Axes(2).TickLabels.Font.Size = 9  # value axis
        except Exception:
            pass

        # turn off major gridlines for readability
        try:
            xl_chart.Axes(2).MajorGridlines.Format.Line.Visible = False
        except Exception:
            try:
                xl_chart.Axes(2).MajorGridlines.Format.Visible = False
            except Exception:
                pass

    except Exception:
        pass


def create_advanced_dashboard(wb: xw.Book, df: pd.DataFrame, dashboard_name: str = "AdvancedDashboard",
                              metrics: list | None = None, category_field: str | None = None,
                              date_field: str | None = None, top_n: int = 10):
    """
    Enhanced dashboard builder: KPIs, styled charts, pivot + slicer placed on dashboard,
    and best-effort formatting to make it look polished.
    """
    # Create/clear dashboard sheet
    if dashboard_name in [s.name for s in wb.sheets]:
        main = wb.sheets[dashboard_name]
        main.clear()
    else:
        main = wb.sheets.add(dashboard_name)

    # Big title
    main.range("A1").value = dashboard_name
    try:
        main.range("A1").api.Font.Size = 18
        main.range("A1").api.Font.Bold = True
    except Exception:
        pass

    # Prepare df and detect types
    df = df.copy()
    from .core import detect_column_types
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

    # KPI area coordinates
    kpi_row = 2
    kpi_col = 2

    # Build KPI definitions (label, value, currency_flag)
    kpi_defs = []
    if metrics:
        p = metrics[0]
        kpi_defs = [
            (f"Total {p}", float(df[p].sum()), True),
            (f"Avg {p}", float(df[p].mean()), True),
            (f"Max {p}", float(df[p].max()), True),
            (f"Min {p}", float(df[p].min()), True),
        ]
    else:
        kpi_defs = [("Rows", len(df), False), ("Cols", df.shape[1], False)]

    # write KPI cards with formatting
    for idx, (label, val, currency) in enumerate(kpi_defs):
        r = kpi_row + (idx // 2)
        c = kpi_col + (idx % 2) * 3
        _format_kpi_cell(main, r, c, label, val, currency=currency)
        # add subtle border around the range covering label+value
        try:
            cell_range = main.range((r, c), (r, c + 1))
            cell_range.api.Borders.Weight = 2
        except Exception:
            pass

    # hidden source sheet
    src_name = f"{dashboard_name}_src"
    src = write_dataframe_to_sheet(wb, src_name, df, start_cell="A1", clear_first=True)
    try:
        src.api.Visible = False
    except Exception:
        pass

    # Chart positions and sizes
    left_base = 12
    top_base = 120
    width_big = 640
    width_med = 480
    height_med = 280
    height_small = 240

    # 1) Trend chart (monthly)
    if date_field and metrics:
        tmp = df.copy()
        tmp[date_field] = pd.to_datetime(tmp[date_field], errors="coerce")
        tmp = tmp.dropna(subset=[date_field])
        tmp["_period"] = tmp[date_field].dt.to_period("M").dt.to_timestamp()
        agg = tmp.groupby("_period")[metrics[0]].sum().reset_index()
        trend_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_trend", agg)
        try: trend_sheet.api.Visible = False
        except: pass
        rng = f"A1:B{len(agg)+1}"
        chart = _add_native_chart(main, trend_sheet, rng, chart_type=4, left=left_base, top=top_base, width=width_med, height=height_med, title=f"{metrics[0]} Trend (Monthly)")
        try:
            _style_chart_object(chart)
            # format x-axis as dates
            try:
                chart.Axes(1).TickLabels.NumberFormat = "mmm-yy"
            except Exception:
                pass
        except Exception:
            pass
        top_base += height_med + 20

    # 2) Monthly column for second metric
    if date_field and len(metrics) > 1:
        tmp = df.copy()
        tmp[date_field] = pd.to_datetime(tmp[date_field], errors="coerce")
        tmp = tmp.dropna(subset=[date_field])
        tmp["_period"] = tmp[date_field].dt.to_period("M").dt.to_timestamp()
        agg2 = tmp.groupby("_period")[metrics[1]].sum().reset_index()
        mo_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_mo2", agg2)
        try: mo_sheet.api.Visible = False
        except: pass
        rng2 = f"A1:B{len(agg2)+1}"
        chart2 = _add_native_chart(main, mo_sheet, rng2, chart_type=51, left=left_base + width_med + 20, top=120, width=width_med, height=height_med, title=f"{metrics[1]} Monthly")
        try:
            _style_chart_object(chart2)
            # add data labels for top points
            try:
                chart2.SeriesCollection(1).HasDataLabels = True
            except Exception:
                pass
        except Exception:
            pass

    # 3) Category top-N chart
    if category_field and metrics:
        catagg = df.groupby(category_field)[metrics[0]].sum().reset_index().sort_values(metrics[0], ascending=False).head(top_n)
        cat_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_cat", catagg)
        try: cat_sheet.api.Visible = False
        except: pass
        rng3 = f"A1:B{len(catagg)+1}"
        cat_chart = _add_native_chart(main, cat_sheet, rng3, chart_type=51, left=10, top=top_base, width=width_big, height=height_med, title=f"Top {top_n} {category_field} by {metrics[0]}")
        try:
            _style_chart_object(cat_chart)
            # rotate category axis labels if long
            try:
                cat_chart.Axes(1).TickLabels.Orientation = -45
            except Exception:
                pass
        except Exception:
            pass

    # 4) Scatter plot
    if len(num_cols) >= 2:
        xcol = num_cols[0]; ycol = num_cols[1]
        sc = df[[xcol, ycol]].dropna()
        sc_sheet = write_dataframe_to_sheet(wb, f"{dashboard_name}_scatter", sc)
        try: sc_sheet.api.Visible = False
        except: pass
        rng_sc = f"A1:B{len(sc)+1}"
        scatter = _add_native_chart(main, sc_sheet, rng_sc, chart_type=-4169, left=720, top=420, width=520, height=320, title=f"{ycol} vs {xcol}")
        try:
            _style_chart_object(scatter)
        except Exception:
            pass

    # 5) Pivot + Slicer: create pivot on pivot sheet but add slicer on main dashboard
    if category_field:
        try:
            pivot_sheet_name = f"{dashboard_name}_pivot"
            if pivot_sheet_name in [s.name for s in wb.sheets]:
                p_sht = wb.sheets[pivot_sheet_name]
                p_sht.clear()
            else:
                p_sht = wb.sheets.add(pivot_sheet_name)

            src_api = src.api
            pc = wb.api.PivotCaches().Create(1, src_api.Range("A1").CurrentRegion)
            pt = p_sht.api.PivotTables().Add(PivotCache=pc, TableDestination=p_sht.api.Range("A3"), TableName="PT_"+dashboard_name)
            try:
                pt.PivotFields(category_field).Orientation = 1
                pt.AddDataField(pt.PivotFields(metrics[0]), f"Sum of {metrics[0]}", -4157)
            except Exception:
                pass

            # Add slicer BUT place the slicer on the main dashboard sheet for convenience
            try:
                slicer_cache = wb.api.SlicerCaches.Add(pt.PivotFields(category_field))
                # place on main sheet (top-right area)
                slicer_shape = slicer_cache.Slicers.Add(main.api, category_field + "_Slicer", category_field, main.api.Range("H2"))
                # style slicer shape (size, fill, border)
                try:
                    slicer_shape.Shape.Height = 180
                    slicer_shape.Shape.Width = 200
                    # fill color
                    color = _rgb(250, 250, 250)
                    slicer_shape.Shape.Fill.ForeColor.RGB = color
                    slicer_shape.Shape.Line.Weight = 1
                except Exception:
                    pass
            except Exception:
                pass

            try:
                p_sht.autofit()
            except Exception:
                pass

        except Exception:
            logger.exception("Pivot/slicer creation failed")

    # final autofit
    try:
        main.autofit()
    except Exception:
        pass

    return main.name