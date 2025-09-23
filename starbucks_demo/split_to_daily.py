from pathlib import Path
import pandas as pd
import re

DATA_DIR = Path("data")

def store_folder_from_name(name: str) -> str:
    # works for names like "starbucks_store101_week (1).csv" or "store101_sales.csv"
    m = re.search(r"store\s*[_-]?(\d+)", name, re.IGNORECASE)
    if m:
        return f"store{m.group(1)}"
    # fallback
    stem = Path(name).stem.replace(" ", "_").lower()
    return stem

def split_weekly_to_daily():
    csvs = list(DATA_DIR.glob("*.csv"))  # your weekly files are directly under data/
    if not csvs:
        print("No weekly CSVs found in ./data/")
        return

    for f in csvs:
        print(f"Processing {f.name} ...")
        df = pd.read_csv(f)
        # normalize required columns
        required = {"date", "order_id", "revenue"}
        if not required.issubset(df.columns):
            raise ValueError(f"{f.name} must include columns: {required}")
        if "category" not in df.columns: df["category"] = "Unknown"
        if "item" not in df.columns:     df["item"] = "Unknown Item"

        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"])

        store_folder = DATA_DIR / store_folder_from_name(f.name)
        store_folder.mkdir(exist_ok=True)

        # write one CSV per calendar day
        for d, day_df in df.groupby(df["date"].dt.date):
            out_path = store_folder / f"{d.isoformat()}.csv"
            day_df.to_csv(out_path, index=False)
        print(f"  -> wrote daily files to {store_folder}")

    print("Done splitting weekly files.")

if __name__ == "__main__":
    split_weekly_to_daily()