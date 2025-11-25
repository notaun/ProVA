# excel_module/app.py
"""
Small CLI entrypoint and sample generator (keeps things simple).
"""
import argparse
import logging
import pandas as pd
import numpy as np
import os
from .commands import Session

logger = logging.getLogger("excel_module.app")
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def generate_sample_xlsx(path: str = "sample_sales.xlsx", rows: int = 120) -> str:
    rng = pd.date_range("2020-01-01", periods=rows, freq="D")
    np.random.seed(0)
    data = {
        "date": rng,
        "region": np.random.choice(["North", "South", "East", "West"], size=len(rng)),
        "sales": np.random.poisson(lam=200.0, size=len(rng)) * np.random.choice([1, 1, 1, 2], size=len(rng)),
        "expenses": np.random.poisson(lam=150.0, size=len(rng)),
        "customer": np.random.choice([f"Cust_{i}" for i in range(1, 51)], size=len(rng)),
    }
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)
    logger.info("Generated sample XLSX: %s", os.path.abspath(path))
    return path


def example_flow(path: str, dashboard_name: str = "AutoDashboard", open_in_excel: bool = False):
    sess = Session(path=path)
    sess.load(path)
    sess.clean()
    sess.write_to_excel(sheet_name="SourceData")
    sess.create_dashboard(dashboard_name=dashboard_name, visible=False, open_in_excel=open_in_excel)
    sess.close()
    logger.info("Example flow complete. Check workbook: %s", os.path.abspath(path))


def main():
    parser = argparse.ArgumentParser(description="Excel module minimal CLI")
    parser.add_argument("--generate-sample", action="store_true", help="Create a sample XLSX for testing")
    parser.add_argument("--input", "-i", type=str, help="Path to Excel (.xlsx) file")
    parser.add_argument("--dashboard", "-d", type=str, default="AutoDashboard", help="Dashboard sheet name")
    parser.add_argument("--open", action="store_true", help="Open the saved workbook in Excel after creation (Windows only)")
    args = parser.parse_args()

    if args.generate_sample:
        path = generate_sample_xlsx()
        example_flow(path, dashboard_name=args.dashboard, open_in_excel=args.open)
    elif args.input:
        if not os.path.exists(args.input):
            print(f"Input file not found: {args.input}")
            return
        example_flow(args.input, dashboard_name=args.dashboard, open_in_excel=args.open)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
