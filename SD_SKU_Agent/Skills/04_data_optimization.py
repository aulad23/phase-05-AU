"""
SKILL: Data Optimization
Duplicate removal, column standardization, batch processing patterns.
"""

import pandas as pd
from pathlib import Path


# ── DEDUPLICATION ─────────────────────────────────────────────────────────────
def dedup_by_url(df: pd.DataFrame, url_col: str = "Product URL") -> pd.DataFrame:
    before = len(df)
    df = df.drop_duplicates(subset=[url_col], keep="first").reset_index(drop=True)
    print(f"  Dedup: {before} → {len(df)} rows")
    return df


def dedup_by_sku(df: pd.DataFrame, sku_col: str = "SKU") -> pd.DataFrame:
    before = len(df)
    df = df.dropna(subset=[sku_col])
    df = df.drop_duplicates(subset=[sku_col], keep="first").reset_index(drop=True)
    print(f"  Dedup (SKU): {before} → {len(df)} rows")
    return df


# ── COLUMN FILL ───────────────────────────────────────────────────────────────
def fill_manufacturer(df: pd.DataFrame, brand: str) -> pd.DataFrame:
    df["Manufacturer"] = brand
    return df


def fill_source(df: pd.DataFrame, url: str) -> pd.DataFrame:
    df["Source"] = df.get("Product URL", url)
    return df


# ── BATCH PROCESSING ──────────────────────────────────────────────────────────
def batch_iter(items: list, size: int = 50):
    """Yield batches of `size` from a list."""
    for i in range(0, len(items), size):
        yield items[i:i + size]


# ── COLUMN REORDER ────────────────────────────────────────────────────────────
STANDARD_COLUMNS = [
    "Manufacturer", "Source", "Image URL", "Product Name", "SKU",
    "Product Family Id", "Description", "Weight", "Width", "Depth",
    "Diameter", "Length", "Height", "Seat Width", "Seat Depth",
    "Seat Height", "Arm Height", "Arm Width", "List Price",
]

def reorder_columns(df: pd.DataFrame, extra_cols: list = None) -> pd.DataFrame:
    cols = STANDARD_COLUMNS + (extra_cols or [])
    final = [c for c in cols if c in df.columns]
    missing = [c for c in cols if c not in df.columns]
    for c in missing:
        df[c] = ""
    return df[final + [c for c in df.columns if c not in cols]]


# ── SUMMARY REPORT ────────────────────────────────────────────────────────────
def print_summary(df: pd.DataFrame, vendor: str):
    print(f"\n  === {vendor} Summary ===")
    print(f"  Total rows   : {len(df)}")
    print(f"  Columns      : {len(df.columns)}")
    filled = {c: df[c].notna().sum() for c in df.columns}
    empty_cols = [c for c, n in filled.items() if n == 0]
    if empty_cols:
        print(f"  Empty cols   : {', '.join(empty_cols)}")
    print(f"  ========================")
