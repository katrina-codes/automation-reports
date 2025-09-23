import re
from pathlib import Path
import pandas as pd
from .config import DATA_DIR


def pretty_store_name_from_path(p: Path) -> str:
    # Prefer parent folder as the store id (store101). Fallback to filename.
    stem = p.parent.name if p.parent != DATA_DIR else p.stem
    m = re.search(r"store\s*[_-]?(\d+)", stem, re.IGNORECASE)
    return f"Store{m.group(1)}" if m else re.sub(r"[_\-]+", " ", stem).title()[:31]


def load_csv_normalized(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    need = {"date", "order_id", "revenue"}
    if not need.issubset(df.columns):
        raise ValueError(f"{path} must include columns: {need}")
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    if "category" not in df.columns:
        df["category"] = "Unknown"
    if "item" not in df.columns:
        df["item"] = "Unknown Item"
    return df


def collect_store_frames() -> dict[str, pd.DataFrame]:
    """
    Returns dict: {store_name: concatenated DataFrame of all daily CSVs}
    Accepts nested folders under data/ or flat CSVs inside data/.
    """
    all_csvs = list(DATA_DIR.rglob("*.csv"))
    if not all_csvs:
        raise SystemExit("No CSVs found under ./data/. Add daily EOD files per store.")

    stores: dict[str, list[pd.DataFrame]] = {}
    for csv in all_csvs:
        store = pretty_store_name_from_path(csv)
        stores.setdefault(store, []).append(load_csv_normalized(csv))

    return {k: pd.concat(v, ignore_index=True).sort_values("date") for k, v in stores.items()}
