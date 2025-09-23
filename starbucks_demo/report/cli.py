import argparse
import pandas as pd
from .io_load import collect_store_frames
from .metrics import kpis_for_window
from .excel_report import write_excel
from .pdf_export import export_excel_to_pdf


def run():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days-week", type=int, default=7)
    ap.add_argument("--days-month", type=int, default=30)
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()

    stores = collect_store_frames()
    today = pd.Timestamp.today().normalize()
    week_start = today - pd.Timedelta(days=args.days_week)
    month_start = today - pd.Timedelta(days=args.days_month)

    weekly_rows, monthly_rows, weekly_tabs, monthly_tabs = [], [], {}, {}

    for store, df in stores.items():
        # Weekly KPIs
        w = kpis_for_window(df, week_start, today)
        weekly_rows.append({
            "Store": store,
            "Week Range": f"{week_start.date():%b %d} - {today.date():%b %d, %Y}",
            **{k: v for k, v in w.items()
               if k not in ["Category Revenue", "Daily Revenue", "Top 3 Items", "Bottom 3 Items", "AOV by Category"]},
        })
        weekly_tabs[store] = {k: w[k] for k in
                              ["Category Revenue", "Top 3 Items", "Bottom 3 Items", "AOV by Category"]}

        # Monthly KPIs
        m = kpis_for_window(df, month_start, today)
        monthly_rows.append({
            "Store": store,
            "Month Range": f"{month_start.date():%b %d} - {today.date():%b %d, %Y}",
            **{k: v for k, v in m.items()
               if k not in ["Category Revenue", "Daily Revenue", "Top 3 Items", "Bottom 3 Items", "AOV by Category"]},
        })
        monthly_tabs[store] = {k: m[k] for k in
                               ["Category Revenue", "Top 3 Items", "Bottom 3 Items", "AOV by Category"]}

    weekly_df = pd.DataFrame(weekly_rows, columns=[
        "Store", "Week Range", "Revenue", "Orders", "AOV",
        "Drinks %", "Food %", "Seasonal %", "% Orders w/ Food", "Peak Day",
    ])
    monthly_df = pd.DataFrame(monthly_rows, columns=[
        "Store", "Month Range", "Revenue", "Orders", "AOV",
        "Drinks %", "Food %", "Seasonal %", "% Orders w/ Food", "Peak Day",
    ])

    xlsx_path = write_excel(weekly_df, monthly_df, weekly_tabs, monthly_tabs)
    if not args.no_pdf:
        export_excel_to_pdf(xlsx_path, xlsx_path.with_suffix(".pdf"))
