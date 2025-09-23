# report/excel_report.py
import datetime as dt
from pathlib import Path
import pandas as pd
from .config import OUTDIR


def _prep(df: pd.DataFrame) -> pd.DataFrame:
    """Add % of Total and Rank; sort by Revenue desc."""
    if df.empty:
        return df
    total = df["Revenue"].sum() or 1.0
    df = df.copy()
    df["% of Total"] = (df["Revenue"] / total * 100).round(1)
    df = df.sort_values("Revenue", ascending=False).reset_index(drop=True)
    df.insert(0, "Rank", df.index + 1)
    return df


def _set_col_widths(ws, widths: list[float]):
    """Apply a list of widths to columns 0..N."""
    for idx, w in enumerate(widths):
        ws.set_column(idx, idx, w)


def _freeze_and_page(ws, nrows: int, ncols: int):
    """Consistent page setup + freeze header row."""
    ws.freeze_panes(1, 0)        # freeze first row (header)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)        # 1 page wide, N tall
    ws.set_margins(0.5, 0.5, 0.6, 0.6)
    ws.repeat_rows(0)
    ws.hide_gridlines(2)
    if nrows and ncols:
        ws.print_area(0, 0, nrows - 1, ncols - 1)


def write_excel(weekly_df: pd.DataFrame,
                monthly_df: pd.DataFrame,
                store_tabs_week: dict,
                store_tabs_month: dict) -> Path:
    """Build a polished XLSX with summaries, charts, and per-store tabs."""
    # Output name
    week_range = weekly_df["Week Range"].iloc[0] if not weekly_df.empty else "No_Week"
    slug = week_range.replace(" ", "_").replace(",", "")
    ts = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out_xlsx = OUTDIR / f"franchise_report_{slug}_{ts}.xlsx"

    weekly_ranked = _prep(weekly_df)
    monthly_ranked = _prep(monthly_df)

    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as xw:
        wb = xw.book

        # ---- Formats
        f_hdr   = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1,
                                 "align": "center", "valign": "vcenter"})
        f_cell  = wb.add_format({"border": 1, "valign": "vcenter"})
        f_money = wb.add_format({"num_format": "$#,##0", "border": 1, "valign": "vcenter"})
        f_aov   = wb.add_format({"num_format": "$#,##0.00", "border": 1, "valign": "vcenter"})
        f_pct   = wb.add_format({"num_format": "0.0%", "border": 1, "valign": "vcenter"})
        f_title = wb.add_format({"bold": True, "font_size": 14})
        f_sub   = wb.add_format({"italic": True, "font_color": "#666666"})

        # ---- Helpers
        def write_table(ws, df, title=None,
                        money=("Revenue",),
                        pct=("Drinks %", "Food %", "Seasonal %", "% of Total", "% Orders w/ Food"),
                        aov=("AOV",)):
            """Write a DataFrame with header/body formats; return (last_row, ncols)."""
            cols = list(df.columns)
            r0 = 0
            if title:
                ws.write(r0, 0, title, f_title)
                r0 += 1
            # header
            for j, c in enumerate(cols):
                ws.write(r0, j, c, f_hdr)
            # body
            for i in range(len(df)):
                for j, c in enumerate(cols):
                    v = df.iloc[i, j]
                    if c in money:
                        ws.write(r0 + 1 + i, j, v, f_money)
                    elif c in aov:
                        ws.write(r0 + 1 + i, j, v, f_aov)
                    elif c in pct:
                        ws.write(r0 + 1 + i, j,
                                 (v / 100 if isinstance(v, (int, float)) and v > 1 else v),
                                 f_pct)
                    else:
                        ws.write(r0 + 1 + i, j, v, f_cell)

            # sensible widths for summary tables
            # [Rank, Store, Week/Month Range, Revenue, Orders, AOV, Drinks %, Food %, Seasonal %,
            #  % Orders w/ Food, Peak Day, % of Total]
            _set_col_widths(ws, [6, 16, 24, 12, 10, 10, 10, 10, 12, 16, 16, 11])

            last_row = r0 + 1 + len(df)   # zero-based index of last written row
            ncols = len(cols)
            _freeze_and_page(ws, last_row, ncols)
            return last_row, ncols

        def add_bar(ws, title, cat_col, val_col, n, anchor):
            if not n:
                return
            ch = wb.add_chart({"type": "column"})
            ch.add_series({
                "name": title,
                "categories": [ws.get_name(), 2, cat_col, n + 1, cat_col],
                "values":     [ws.get_name(), 2, val_col, n + 1, val_col],
                "data_labels": {"value": True},
            })
            ch.set_title({"name": title})
            ch.set_y_axis({"num_format": "$#,##0"})
            ch.set_legend({"position": "none"})
            ws.insert_chart(anchor, ch, {"x_scale": 1.25, "y_scale": 1.1})

        # ---- Summary – Weekly
        ws_w = wb.add_worksheet("Summary – Weekly")
        t_end_w, t_cols_w = write_table(ws_w, weekly_ranked, title="Weekly Summary")
        n_w = len(weekly_ranked)
        add_bar(ws_w, "Weekly Revenue by Store", 1, 3, n_w, "J3")
        add_bar(ws_w, "Weekly AOV by Store",     1, 5, n_w, "J22")
        # extend print area to include charts (buffer ~30 rows, min 13 cols to capture charts)
        ws_w.print_area(0, 0, max(t_end_w + 30, t_end_w), max(t_cols_w - 1, 12))

        # ---- Summary – Monthly
        ws_m = wb.add_worksheet("Summary – Monthly")
        t_end_m, t_cols_m = write_table(ws_m, monthly_ranked, title="Monthly Summary")
        n_m = len(monthly_ranked)
        add_bar(ws_m, "Monthly Revenue by Store", 1, 3, n_m, "J3")
        add_bar(ws_m, "Monthly AOV by Store",     1, 5, n_m, "J22")
        ws_m.print_area(0, 0, max(t_end_m + 30, t_end_m), max(t_cols_m - 1, 12))

        # ---- Per-store sheets
        for store in sorted(store_tabs_week.keys()):
            ws = wb.add_worksheet(store[:31])
            # give breathing room; Excel ignores extra if fewer cols exist
            _set_col_widths(ws, [18, 12, 12, 12, 12, 12, 12, 12])

            r = 0
            max_cols = 1
            sections = [
                (f"{store} – WEEKLY • Category Revenue", store_tabs_week[store]["Category Revenue"]),
                (f"{store} – WEEKLY • Top 3 Items",      store_tabs_week[store]["Top 3 Items"]),
                (f"{store} – WEEKLY • Bottom 3 Items",   store_tabs_week[store]["Bottom 3 Items"]),
                (f"{store} – WEEKLY • AOV by Category",  store_tabs_week[store]["AOV by Category"]),
                (f"{store} – MONTHLY • Category Revenue", store_tabs_month[store]["Category Revenue"]),
                (f"{store} – MONTHLY • Top 3 Items",      store_tabs_month[store]["Top 3 Items"]),
                (f"{store} – MONTHLY • Bottom 3 Items",   store_tabs_month[store]["Bottom 3 Items"]),
                (f"{store} – MONTHLY • AOV by Category",  store_tabs_month[store]["AOV by Category"]),
            ]
            for label, df in sections:
                ws.write(r, 0, label, f_sub)
                r += 1
                cols = list(df.columns)
                max_cols = max(max_cols, len(cols))
                # header
                for j, c in enumerate(cols):
                    ws.write(r, j, c, f_hdr)
                # body
                for i in range(len(df)):
                    for j, c in enumerate(cols):
                        ws.write(r + 1 + i, j, df.iloc[i, j], f_cell)
                r += len(df) + 3

            _freeze_and_page(ws, r, max_cols)

    return out_xlsx
