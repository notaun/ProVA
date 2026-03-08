
Project: excel_module — voice/GUI-driven Excel dashboard automation

Overview:
- core.py: data loading/cleaning/type-detection helpers (pandas).
- excel_engine.py: Excel automation via xlwings (open workbook, write sheet, create charts/pivots).
- commands.py: Session class exposing a small API for a single pipeline run:
    Session.load(path) -> loads data to memory
    Session.clean() -> cleans the data
    Session.write_to_excel(sheet_name, visible) -> writes cleaned data to sheet and saves
    Session.create_dashboard(dashboard_name, visible, open_in_excel) -> creates the dashboard sheet
    Session.close() -> saves/closes workbook and quits Excel app

Main design principles:
- Open workbook once and perform write + dashboard operations in the same Excel session to avoid split files.
- Try COM save first; if it fails, do a single pandas fallback to preserve data.
- Keep UI thin: launch_gui.py uses Session API; all business logic is package-contained.

How to run (developer):
1. Install deps (Windows + Excel required): pip install pandas xlwings numpy
2. Open `launch_gui.py` (double-click or `python launch_gui.py`), pick an .xlsx and press Run.
3. For CLI: run `python -m excel_module.app --generate-sample` or `python -m excel_module.app --input path\to\file.xlsx`.

Important notes:
- Pivot tables and slicers require Excel on Windows.
- If COM save repeatedly fails, look for Excel modal dialogs (add-ins or permissions). Use `visible=True` to debug.
- pandas fallback will **not** preserve charts/pivots; it's a data-preservation safety net.

Files of interest:
- excel_module/core.py
- excel_module/excel_engine.py
- excel_module/commands.py
- excel_module/app.py
- launch_gui.py (UI)
