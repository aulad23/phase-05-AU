# step3_rating_logic_update.py
# Same logic, only updated to detect Rating from <td data-label-key="Y3JpdGVyaWFfMTE=">...</td>
# Added three new columns: Extension, Backplate, Base

import re
import pandas as pd

INPUT_XLSX  = "visualcomfort_bulbs_full_data.xlsx"
OUTPUT_XLSX = "visualcomfort_bulbs_final_data.xlsx"

# ------------------------ Label policy ------------------------
DIMENSION_LABELS = {
    "Height", "O/A Height", "Overall Height", "Fixture Height", "Min. Custom Height",
    "Max Custom Height", "Min Custom Height", "Width", "Depth", "Extension", "Projection",
    "Diameter", "Overall", "Backplate", "Canopy", "Shade Details", "Shade",
    "Length", "Drop", "Backplate Width", "Backplate Height", "Backplate Depth",
    "Overall Width", "Overall Depth", "Overall Length",
    "Chain Length", "O/A Height", "Fixture Height", "Overall Height"
}

NON_DIMENSION_ALWAYS_KEEP = {
    "Socket", "Wattage", "Lightsource", "Light Source", "Bulb", "Bulb Type",
    "Weight", "Finish", "Material", "Notes", "Care", "Collection", "Certifications",
    "Switch", "Switch Type", "Hardwire Portable", "Shipped Via","Product Family Id",
    "General Delivery If Not In Stock", "Brand", "Type", "Match", "Reversible",
    "Features", "Minimum Order Quantity", "Order Increment","List Price",
    "CFA Requests", "Freight Code",
    # keep Rating always
    "Rating"
}

SPECIAL_KEEP_FEET_OK = {"Chain Length"}

INCH_WORDS = (" inch", " inches", " in.", " in ")

# ------------------------ Helpers ------------------------
def normalize_text(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        s = s[1:-1]
    s = s.replace('""', '"')
    s = re.sub(r'\r\n?|\n', '\n', s)
    return s.strip()

def clean_text(x: str) -> str:
    return re.sub(r'\s+', ' ', x).strip()

def looks_inch(value: str) -> bool:
    v = value.lower()
    return ('"' in value) or any(w in v for w in INCH_WORDS)

def looks_feet(value: str) -> bool:
    v = value.lower()
    return ("'" in v) or (" ft" in v) or ("feet" in v)

def parse_lines_to_multimap(spec_text: str):
    """
    Parse plain text or HTML-like blocks from the 'Specifications' field.
    - Normal lines: 'Label: Value'
    - HTML line for Rating: <td data-label-key="Y3JpdGVyaWFfMTE=">Rating: Damp Rated</td>
    """
    mm = {}
    s = normalize_text(spec_text)
    if not s:
        return mm

    # --- check for embedded HTML rating cell ---
    rating_match = re.search(r'<td[^>]*data-label-key="Y3JpdGVyaWFfMTE="[^>]*>(.*?)</td>', s, re.IGNORECASE | re.DOTALL)
    if rating_match:
        text = clean_text(rating_match.group(1))
        # expected format: Rating: Damp Rated
        if text.lower().startswith("rating:"):
            val = clean_text(text.split(":", 1)[1])
            mm.setdefault("Rating", []).append(val)

    # --- fallback: parse text lines ---
    for raw in s.split("\n"):
        line = raw.strip()
        if not line or ":" not in line:
            continue
        label, val = line.split(":", 1)
        label = clean_text(label)
        val = clean_text(val)
        if not label or not val:
            continue
        if val.lower() == label.lower():
            continue
        mm.setdefault(label, []).append(val)

    return mm

def choose_final_value(label: str, values: list[str]) -> str | None:
    if label in DIMENSION_LABELS:
        for v in values:
            if looks_inch(v):
                return v
        if label in SPECIAL_KEEP_FEET_OK:
            for v in values:
                if looks_feet(v) or looks_inch(v):
                    return v
        return None
    for v in values:
        if v:
            return v
    return None

def expand_specs_column(df: pd.DataFrame, specs_col: str = "Specifications") -> pd.DataFrame:
    if specs_col not in df.columns:
        print(f'Column "{specs_col}" not found. No changes made.')
        return df
    parsed_rows = []
    all_keys = set()
    for spec_text in df[specs_col].fillna(""):
        mm = parse_lines_to_multimap(spec_text)
        chosen = {}
        for label, vals in mm.items():
            final_val = choose_final_value(label, vals)
            if final_val is not None and label not in chosen:
                chosen[label] = final_val
                all_keys.add(label)
        parsed_rows.append(chosen)
    for key in sorted(all_keys):
        col_name = key
        if col_name in df.columns:
            base = col_name
            i = 2
            while col_name in df.columns:
                col_name = f"{base} ({i})"
                i += 1
        df[col_name] = [row.get(key, None) for row in parsed_rows]
    return df

# ------------------------ Run ------------------------
df = pd.read_excel(INPUT_XLSX)
df = expand_specs_column(df, specs_col="Specifications")

# ----- Enforce final column order -----
DESIRED_ORDER = [
    "Product URL",
    "Image URL",
    "Product Name",
    "SKU",
    "Product Family Id"
    "List Price",
    "Description",
    "Height",
    "Width",
    "Length",
    "Diameter",
    "Weight",
    "Finish",
    "Canopy",
    "Socket",
    "Wattage",
    "Chain Length",
    "Shade Details",
    "Lightsource",
    "Rating",
    "O/A Height",
    "Fixture Height",
    "Overall Height",
    # new columns
    "Extension",
    "Backplate",
    "Base"
]

for col in DESIRED_ORDER:
    if col not in df.columns:
        df[col] = None

remaining = [c for c in df.columns if c not in DESIRED_ORDER]
df = df[DESIRED_ORDER + remaining]

df.to_excel(OUTPUT_XLSX, index=False)
print(f"Done. Wrote: {OUTPUT_XLSX}")
