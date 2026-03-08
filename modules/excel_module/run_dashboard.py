"""
run_dashboard.py — Standalone test/utility script for the ProVA Excel module.

Usage (run from the ProVA root directory):
    python run_dashboard.py path/to/your_data.xlsx

If no path is given the script opens a file-picker dialog.

The Session class has been removed — it served no purpose here because
commands.py already exposes load_data() as a plain function, and
build_dashboard() is called directly via planner + excel_engine.
The one-liner pipeline below is clearer than a class wrapper that only
delegated every call through unchanged.
"""

from pathlib import Path
import sys

# Ensure ProVA root is on the path when running directly
sys.path.insert(0, str(Path(__file__).parent))

from modules.excel_module.commands     import load_data
from modules.excel_module.planner      import plan as make_plan
from modules.excel_module.excel_engine import build_dashboard
from modules.excel_module.file_finder  import pick_file_dialog

_OUT_DIR = Path(__file__).parent / "modules" / "excel_module" / "out"


def run(data_path: Path) -> Path:
    """Load data → plan → build dashboard. Returns the output path."""
    print(f"Loading: {data_path}")
    df = load_data(str(data_path))

    print(f"Columns : {list(df.columns)}")
    print(f"Shape   : {df.shape}")
    print(f"Dtypes  :\n{df.dtypes}\n")

    plan = make_plan(df, source_name=data_path.stem)

    _OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = _OUT_DIR / f"{data_path.stem}_dashboard.xlsx"

    result = build_dashboard(df, plan, str(out_path))
    print(f"Dashboard saved → {result}")
    return result


if __name__ == "__main__":
    if len(sys.argv) > 1:
        path = Path(sys.argv[1])
        if not path.exists():
            sys.exit(f"File not found: {path}")
    else:
        print("No file specified — opening file picker…")
        path = pick_file_dialog()
        if path is None:
            sys.exit("No file selected. Exiting.")

    run(path)