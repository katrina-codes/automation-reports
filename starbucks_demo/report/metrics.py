import pandas as pd


def kpis_for_window(df: pd.DataFrame, start, end):
    win = df[(df["date"] >= start) & (df["date"] < end)].copy()

    revenue = float(win["revenue"].sum()) if not win.empty else 0.0
    orders = int(win["order_id"].nunique()) if not win.empty else 0
    aov = round(revenue / orders, 2) if orders else 0.0

    cat_rev = (
        win.groupby("category", dropna=False)["revenue"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
        if not win.empty
        else pd.DataFrame(columns=["category", "revenue"])
    )
    total = revenue or 1.0
    if not cat_rev.empty:
        cat_rev["percent_of_total"] = (cat_rev["revenue"] / total * 100).round(1)

    def pct(cat: str) -> float:
        if cat_rev.empty:
            return 0.0
        r = cat_rev[cat_rev["category"].str.lower() == cat.lower()]
        return float(r["percent_of_total"].iloc[0]) if not r.empty else 0.0

    # Daily revenue (for peak day)
    daily_rev = (
        win.groupby(win["date"].dt.date)["revenue"]
        .sum()
        .reset_index()
        .rename(columns={"date": "day"})
        if not win.empty
        else pd.DataFrame(columns=["day", "revenue"])
    )
    peak_day = ""
    if not daily_rev.empty:
        p = daily_rev.sort_values("revenue", ascending=False).iloc[0]
        peak_day = f"{pd.to_datetime(p['day']).strftime('%a')} (${p['revenue']:.2f})"

    # Items
    items_rev = (
        win.groupby("item", dropna=False)["revenue"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
        if not win.empty
        else pd.DataFrame(columns=["item", "revenue"])
    )

    # AOV by category helper
    def aov_by(cat: str):
        sub = win[win["category"].str.lower() == cat.lower()]
        o = int(sub["order_id"].nunique()) if not sub.empty else 0
        rev = float(sub["revenue"].sum()) if not sub.empty else 0.0
        return [cat, o, rev, round(rev / o, 2) if o else 0.0]

    return {
        "Revenue": round(revenue, 2),
        "Orders": orders,
        "AOV": aov,
        "Drinks %": pct("drink"),
        "Food %": pct("food"),
        "Seasonal %": pct("seasonal"),
        "% Orders w/ Food": round(
            (int(win[win["category"].str.lower() == "food"]["order_id"].nunique()) / orders * 100), 1
        ) if orders else 0.0,
        "Peak Day": peak_day,
        "Category Revenue": cat_rev,
        "Daily Revenue": daily_rev,
        "Top 3 Items": items_rev.head(3),
        "Bottom 3 Items": items_rev.tail(3) if len(items_rev) >= 3 else items_rev,
        "AOV by Category": pd.DataFrame(
            [aov_by(c) for c in ["Drink", "Food", "Seasonal"]],
            columns=["category", "orders", "revenue", "aov"],
        ),
    }
