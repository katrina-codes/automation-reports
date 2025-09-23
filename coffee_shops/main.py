import datetime as dt
import pandas as pd
from pathlib import Path

DATA = Path("data/sales.csv")
OUTDIR = Path("output")
OUTDIR.mkdir(exist_ok=True)

# --- Add a simple category map (edit as you like) ---
ITEM_CATEGORY = {
    "House Blend": "Drink",
    "Latte": "Drink",
    "Mocha": "Drink",
    "Cold Brew": "Drink",
    "Matcha Latte": "Drink",
    "Blueberry Muffin": "Pastry",
    "Croissant": "Pastry",
}

def compute_metrics(df: pd.DataFrame):
    df["date"] = pd.to_datetime(df["date"])
    today = pd.Timestamp.today().normalize()

    # Windows: past 7 days = "this week" window; prior 7 = last week
    this_week_mask = df["date"] >= (today - pd.Timedelta(days=7))
    last_week_mask = (df["date"] < (today - pd.Timedelta(days=7))) & \
                     (df["date"] >= (today - pd.Timedelta(days=14)))

    this_week = df[this_week_mask].copy()
    last_week = df[last_week_mask].copy()

    # ---- Base metrics ----
    rev_this = float(this_week["revenue"].sum())
    rev_last = float(last_week["revenue"].sum())

    orders_this = int(this_week["order_id"].nunique())
    orders_last = int(last_week["order_id"].nunique())

    aov_this = round(rev_this / orders_this, 2) if orders_this else 0.0
    aov_last = round(rev_last / orders_last, 2) if orders_last else 0.0

    wow_pct = round(((rev_this - rev_last) / rev_last * 100.0), 1) if rev_last else None

    # ---- Extra KPIs ----
    units_this = int(len(this_week))  # each row is a line item
    # Daily revenue (date-by-date)
    daily_rev = (this_week.groupby(this_week["date"].dt.date)["revenue"]
                 .sum().reset_index().rename(columns={"date":"day","revenue":"revenue"}))
    # Peak day
    peak_day = daily_rev.sort_values("revenue", ascending=False).head(1) if not daily_rev.empty else None

    # Top items by revenue (already had)
    top_items_revenue = (this_week.groupby("item")["revenue"]
                         .sum().sort_values(ascending=False).reset_index())

    # Top items by units sold
    top_items_units = (this_week.groupby("item")["item"]
                       .count().sort_values(ascending=False)
                       .reset_index(name="units"))

    # Category revenue (map unknowns to "Other")
    this_week["category"] = this_week["item"].map(ITEM_CATEGORY).fillna("Other")
    category_rev = (this_week.groupby("category")["revenue"]
                    .sum().sort_values(ascending=False).reset_index())

    metrics = {
        "rev_this": round(rev_this, 2),
        "rev_last": round(rev_last, 2),
        "orders_this": orders_this,
        "orders_last": orders_last,
        "aov_this": aov_this,
        "aov_last": aov_last,
        "wow_pct": wow_pct,
        "units_this": units_this,
        "daily_rev": daily_rev,
        "peak_day": peak_day,
        "top_items_revenue": top_items_revenue,
        "top_items_units": top_items_units,
        "category_rev": category_rev,
    }
    return metrics

def save_report(m):
    ts = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")  # unique filename
    xlsx = OUTDIR / f"weekly_report_{ts}.xlsx"

    with pd.ExcelWriter(xlsx) as xw:
        # Summary
        pd.DataFrame([
            ["Revenue (this week)", m["rev_this"]],
            ["Revenue (last week)", m["rev_last"]],
            ["WoW Change (%)", m["wow_pct"]],
            ["Orders (this week)", m["orders_this"]],
            ["Orders (last week)", m["orders_last"]],
            ["Avg Order Value (this week)", m["aov_this"]],
            ["Avg Order Value (last week)", m["aov_last"]],
            ["Units sold (this week)", m["units_this"]],
        ], columns=["Metric","Value"]).to_excel(xw, index=False, sheet_name="Summary")

        # Daily revenue
        m["daily_rev"].to_excel(xw, index=False, sheet_name="Daily Revenue")

        # Peak day (single row)
        if m["peak_day"] is not None and not m["peak_day"].empty:
            m["peak_day"].to_excel(xw, index=False, sheet_name="Peak Day")

        # Top items by revenue / by units
        m["top_items_revenue"].to_excel(xw, index=False, sheet_name="Top Items (Revenue)")
        m["top_items_units"].to_excel(xw, index=False, sheet_name="Top Items (Units)")

        # Category revenue
        m["category_rev"].to_excel(xw, index=False, sheet_name="Category Revenue")

    print("Report generated:", xlsx)

def main():
    df = pd.read_csv(DATA)
    metrics = compute_metrics(df)
    save_report(metrics)

if __name__ == "__main__":
    main()
