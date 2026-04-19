"""
Eichholtz Column Re-Parser (No Scraping)
=========================================
Already scrape করা Excel file নিয়ে Finish, Socket, Dimension columns
re-parse/clean করে updated columns populate করে।

Usage: python eichholtz_reparse.py

INPUT_FILE আর OUTPUT_FILE নিচে change করো তোমার file অনুযায়ী।
"""

import re
import pandas as pd

# ===================== CONFIG =====================
INPUT_FILE  = "eichholtz_Bar_Stools_Final.xlsx"
OUTPUT_FILE = "eichholtz_Bar_Stools_Update.xlsx"
# ==================================================

# ─── Final column order ───
FINAL_COLUMN_ORDER = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family Id",
    "Description", "Weight", "Width", "Depth", "Diameter", "Length", "Height",
    "Seat Depth", "Seat Width", "Seat Height", "Arm Height", "Shade Details",
    "Dimension", "Finish", "Wattage", "Socket", "Fabric", "Specifications",
]

# Dimension abbreviation → column name
DIM_ABBR_MAP = {
    "W": "Width", "D": "Depth", "H": "Height", "L": "Length",
    "SD": "Seat Depth", "SH": "Seat Height", "SW": "Seat Width", "AH": "Arm Height",
}

DIM_COLS = [
    "Width", "Depth", "Height", "Diameter", "Length",
    "Seat Depth", "Seat Width", "Seat Height", "Arm Height", "Shade Details",
]

SPEC_KEYS = ["Finish", "Weight", "Fabric", "Wattage", "Socket"]

# ═══════════════════════════════════════════════════════════
# BOUNDARY KEYS — এগুলো দেখলে value কেটে দেবে
# ═══════════════════════════════════════════════════════════

BOUNDARY_KEYS = [
    "extra info", "lamp holder qty", "lamp holder quantity",
    "light bulbs included", "light bulbs", "bulbs included",
    "voltage", "plug type", "fabric composition shade",
    "shade dimensions", "shade dimension",
    "indoor/outdoor", "indoor outdoor",
    "dimmable", "ip rating", "cord length", "cable length",
    "switch type", "switch", "power source",
    "color temperature", "colour temperature",
    "number of lights", "number of light sources",
    "max wattage", "max weight load", "max weight load kg",
    "light source included",
]


def clean_column_with_boundaries(raw_val):
    """
    Boundary key দেখলে সেখানে কেটে দেয়।

    "Antique brass | alabaster | Extra info: blah blah" → "Antique brass | alabaster"
    "E27 | Lamp holder qty: 1 | Light bulbs..."         → "E27"
    "E27,Integrated LED | Lamp holder qty: 1..."         → "E27,Integrated LED"
    """
    val = str(raw_val).strip()
    if not val or val == "nan":
        return ""

    parts = val.split(" | ")
    clean_parts = []

    for part in parts:
        if ":" in part:
            potential_key = part.split(":", 1)[0].strip().lower()
            if any(potential_key == bk for bk in BOUNDARY_KEYS):
                break
        clean_parts.append(part)

    return " | ".join(clean_parts) if clean_parts else ""


# ═══════════════════════════════════════════════════════════
# PARSE DIMENSION STRING
# ═══════════════════════════════════════════════════════════

def parse_shade_from_text(text):
    """
    "Shade: Bottom Ø 22.83″ | Top Ø 9.84″ | H. 11.81″" → "22.83,9.84,11.81"
    """
    if not text:
        return ""
    m = re.search(r'[Ss]hade[:\s]*(.*?)(?:\n[A-Z]|\Z)', text, re.DOTALL)
    if not m:
        return ""
    shade_text = m.group(1).strip()
    numbers = re.findall(r'(\d+(?:\.\d+)?)(?:\s*[″"])', shade_text)
    return ",".join(numbers) if numbers else ""


def parse_dimension_string(dim_str):
    """
    Parse dimension string → separate columns.
    Handles: W, D, H, L, SD, SH, SW, AH, Ø (Diameter), Shade
    """
    result = {col: "" for col in DIM_COLS}

    if not dim_str or str(dim_str).strip() in ("", "nan"):
        return result

    dim_str = str(dim_str).strip()

    # Diameter: Ø 31.50″
    m_dia = re.search(r'[ØøΦφ]\s*([\d.]+)', dim_str)
    if m_dia:
        result["Diameter"] = str(round(float(m_dia.group(1)), 2))

    # Match longer abbreviations first
    for abbr in ["SD", "SH", "SW", "AH", "W", "D", "H", "L"]:
        pattern = r'(?<![A-Za-z])' + abbr + r'\.\s*([\d.]+)'
        m = re.search(pattern, dim_str)
        if m:
            col_name = DIM_ABBR_MAP.get(abbr, "")
            if col_name:
                result[col_name] = str(round(float(m.group(1)), 2))

    # Shade Details
    result["Shade Details"] = parse_shade_from_text(dim_str)

    return result


# ═══════════════════════════════════════════════════════════
# PARSE SPECIFICATIONS STRING
# ═══════════════════════════════════════════════════════════

SPEC_KEY_ALIASES = {
    "Finish":  ["finish", "general info", "general_info"],
    "Weight":  ["weight", "max weight load", "max weight load lbs", "max_weight_load"],
    "Fabric":  ["fabric", "fabric composition", "fabric_composition", "material composition"],
    "Wattage": ["wattage", "max wattage", "max_wattage"],
    "Socket":  ["socket", "lamp holder", "lamp_holder", "lampholder"],
    "_skip":   BOUNDARY_KEYS,
}


def _match_spec_key(raw_key):
    raw_lower = raw_key.strip().lower()
    for target, aliases in SPEC_KEY_ALIASES.items():
        for alias in aliases:
            if raw_lower == alias:
                return target
    return None


def clean_finish(v):
    v = re.sub(r'^general\s*info\s*[:]\s*', '', v.strip(), flags=re.IGNORECASE).strip()
    return v


def clean_weight(v):
    v = v.strip()
    if re.search(r'\bKG\b', v, re.IGNORECASE) and not re.search(r'\bLBS\b', v, re.IGNORECASE):
        return ""
    m = re.search(r'LBS\s*[:]\s*([\d.]+)', v, re.IGNORECASE)
    if m: return m.group(1)
    m = re.search(r'([\d.]+)\s*lbs', v, re.IGNORECASE)
    if m: return m.group(1)
    v = re.sub(r'^max\s*weight\s*load\s*(lbs|kg)?\s*[:]\s*', '', v, flags=re.IGNORECASE).strip()
    m = re.search(r'([\d.]+)', v)
    return m.group(1) if m else v


def clean_fabric(v):
    v = re.sub(r'^fabric\s*composition\s*[:]\s*', '', v.strip(), flags=re.IGNORECASE).strip()
    return re.sub(r'^fabric\s*[:]\s*', '', v, flags=re.IGNORECASE).strip()


def clean_socket(v):
    return re.sub(r'^lamp\s*holder\s*[:]\s*', '', v.strip(), flags=re.IGNORECASE).strip()


def clean_wattage(v):
    v = re.sub(r'^max\s*wattage\s*[:]\s*', '', v.strip(), flags=re.IGNORECASE).strip()
    m = re.search(r'([\d.]+)\s*(watt|w)\b', v, re.IGNORECASE)
    return f"{m.group(1)} Watt" if m else v


SPEC_CLEANERS = {
    "Finish": clean_finish, "Weight": clean_weight,
    "Fabric": clean_fabric, "Wattage": clean_wattage, "Socket": clean_socket,
}


def parse_specifications_string(spec_str):
    result = {k: "" for k in SPEC_KEYS}
    if not spec_str or str(spec_str).strip() in ("", "nan"):
        return result

    raw_parts = [p.strip() for p in str(spec_str).strip().split(" | ")]
    parsed_pairs = []
    current_key = None
    current_val_parts = []

    for part in raw_parts:
        found_key = False
        if ":" in part:
            potential_key = part.split(":", 1)[0].strip()
            potential_val = part.split(":", 1)[1].strip()
            target = _match_spec_key(potential_key)
            if target:
                if current_key is not None:
                    parsed_pairs.append((current_key, " | ".join(current_val_parts)))
                current_key = target
                current_val_parts = [potential_val] if potential_val else []
                found_key = True
        if not found_key:
            if current_key is not None:
                current_val_parts.append(part)

    if current_key is not None:
        parsed_pairs.append((current_key, " | ".join(current_val_parts)))

    for target_key, raw_val in parsed_pairs:
        if target_key in SPEC_KEYS:
            cleaner = SPEC_CLEANERS.get(target_key, lambda x: x.strip())
            cleaned = cleaner(raw_val)
            if cleaned:
                result[target_key] = cleaned

    return result


# ═══════════════════════════════════════════════════════════
# REORDER COLUMNS
# ═══════════════════════════════════════════════════════════

def reorder_columns(df):
    ordered = [c for c in FINAL_COLUMN_ORDER if c in df.columns]
    for c in df.columns:
        if c not in ordered:
            ordered.append(c)
    return df[ordered]


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    print(f"Reading: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
    total = len(df)
    print(f"Total rows: {total}\n")

    # Ensure all target columns exist
    for col in DIM_COLS + SPEC_KEYS + ["Shade Details", "Wattage"]:
        if col not in df.columns:
            df[col] = ""

    finish_fixed = 0
    socket_fixed = 0
    dim_parsed_count = 0
    spec_parsed_count = 0

    for idx, row in df.iterrows():

        # ── 1) If "Dimension" raw column exists → parse into separate columns ──
        dim_raw = str(row.get("Dimension", "")).strip()
        if dim_raw and dim_raw != "nan":
            dim_parsed = parse_dimension_string(dim_raw)
            for col_name, val in dim_parsed.items():
                if val:
                    df.at[idx, col_name] = val
            dim_parsed_count += 1

        # ── 2) If "Specifications" raw column exists → parse into separate columns ──
        spec_raw = str(row.get("Specifications", "")).strip()
        if spec_raw and spec_raw != "nan":
            spec_parsed = parse_specifications_string(spec_raw)
            for col_name, val in spec_parsed.items():
                if val:
                    df.at[idx, col_name] = val
            spec_parsed_count += 1

        # ── 3) Clean Finish column (remove Extra info etc.) ──
        old_finish = str(row.get("Finish", "")).strip()
        if old_finish and old_finish != "nan":
            new_finish = clean_column_with_boundaries(old_finish)
            if new_finish != old_finish:
                df.at[idx, "Finish"] = new_finish
                finish_fixed += 1

        # ── 4) Clean Socket column (remove Lamp holder qty etc.) ──
        old_socket = str(row.get("Socket", "")).strip()
        if old_socket and old_socket != "nan":
            new_socket = clean_column_with_boundaries(old_socket)
            if new_socket != old_socket:
                df.at[idx, "Socket"] = new_socket
                socket_fixed += 1

    # Final cleanup
    for col in DIM_COLS + SPEC_KEYS:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: "" if str(x).strip() in ("", "nan", "0", "0.0") else x
            )

    df_final = reorder_columns(df)
    df_final.to_excel(OUTPUT_FILE, index=False)

    print(f"{'='*55}")
    print(f"Done! {total} rows processed")
    if dim_parsed_count:
        print(f"  Dimension parsed : {dim_parsed_count} rows")
    if spec_parsed_count:
        print(f"  Specifications parsed: {spec_parsed_count} rows")
    print(f"  Finish cleaned   : {finish_fixed} rows")
    print(f"  Socket cleaned   : {socket_fixed} rows")
    print(f"\nSaved: '{OUTPUT_FILE}'")
    print(f"\nColumn order:")
    for i, col in enumerate(df_final.columns, 1):
        print(f"  {i}. {col}")


if __name__ == "__main__":
    main()