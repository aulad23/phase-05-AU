# -*- coding: utf-8 -*-
import os
import re
import pandas as pd
import math

# =============== CONFIG ===============
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE_NAME  = "rowefurniture_sectional_veriation.xlsx"
OUTPUT_FILE_NAME = "rowefurniture_sectional_Final.xlsx"

INPUT_FILE  = os.path.join(SCRIPT_DIR, INPUT_FILE_NAME)
OUTPUT_FILE = os.path.join(SCRIPT_DIR, OUTPUT_FILE_NAME)
# =====================================

# Now also capture L for Length
DIM_TOKEN_RE = re.compile(r'\b(L|W|DIA|Dia|D|H)\s*([0-9]+(?:\.[0-9]+)?)')

def clean_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

def safe_str(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x)

def parse_specs_block_text(spec_text: str) -> dict:
    """
    Specifications column theke:
    Weight, Dimension, Seat Height, Seat Depth, Arm Height,
    Cushion, Construction, Content, Color,
    Width, Depth, Diameter, Length, Height extract korbe.
    """
    result = {
        "Weight": "",
        "Dimension": "",
        "Seat Height": "",
        "Seat Depth": "",
        "Arm Height": "",
        "Cushion": "",
        "Construction": "",
        "Content": "",
        "Color": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Length": "",
        "Height": "",
    }

    if not spec_text:
        return result

    lines = [clean_text(line) for line in spec_text.splitlines() if clean_text(line)]
    cushion_number = ""
    cushion_standard = ""

    for line in lines:
        low = line.lower()

        # ---------- Weight ----------
        # Example: "Weight (LB): 247"
        if "weight" in low and "(lb" in low and not result["Weight"]:
            m = re.search(r'weight\s*\(lb\).*?([0-9]+(?:\.[0-9]+)?)', line, re.I)
            if m:
                result["Weight"] = m.group(1)
            continue

        # ---------- Dimension ----------
        # Example: "Dimensions (IN): L 75\" x D 40\" x H 35\""
        if "dimension" in low and "(in" in low and not result["Dimension"]:
            # part after ":" if exists, otherwise after label word
            if ":" in line:
                dim_part = line.split(":", 1)[1].strip()
            else:
                # fallback, remove label text
                dim_part = re.sub(r'(?i)dimensions?\s*\(in\)\s*', "", line).strip()

            # normalize quotes
            dim_part = dim_part.replace("”", '"').replace("“", '"')
            # example wants: L 75" x D 40" x H 35  (last quote chere)
            if dim_part.endswith('"'):
                dim_part = dim_part[:-1].strip()

            result["Dimension"] = dim_part

            # now parse Width/Depth/Diameter/Length/Height from dim_part
            tokens = DIM_TOKEN_RE.findall(dim_part)
            for k, v in tokens:
                ku = k.upper()
                if   ku == "W":
                    result["Width"] = v
                elif ku == "D":
                    result["Depth"] = v
                elif ku == "H":
                    result["Height"] = v
                elif ku == "DIA":
                    result["Diameter"] = v
                elif ku == "L":
                    result["Length"] = v
            continue

        # ---------- Seat Height ----------
        if "seat height" in low and "(in" in low and not result["Seat Height"]:
            m = re.search(r'seat height\s*\(in\).*?([0-9]+(?:\.[0-9]+)?)', line, re.I)
            if m:
                result["Seat Height"] = m.group(1)
            else:
                # fallback: last number in line
                m2 = re.search(r'([0-9]+(?:\.[0-9]+)?)', line)
                if m2:
                    result["Seat Height"] = m2.group(1)
            continue

        # ---------- Seat Depth ----------
        if "seat depth" in low and "(in" in low and not result["Seat Depth"]:
            m = re.search(r'seat depth\s*\(in\).*?([0-9]+(?:\.[0-9]+)?)', line, re.I)
            if m:
                result["Seat Depth"] = m.group(1)
            else:
                m2 = re.search(r'([0-9]+(?:\.[0-9]+)?)', line)
                if m2:
                    result["Seat Depth"] = m2.group(1)
            continue

        # ---------- Arm Height ----------
        if "arm height" in low and "(in" in low and not result["Arm Height"]:
            m = re.search(r'arm height\s*\(in\).*?([0-9]+(?:\.[0-9]+)?)', line, re.I)
            if m:
                result["Arm Height"] = m.group(1)
            else:
                m2 = re.search(r'([0-9]+(?:\.[0-9]+)?)', line)
                if m2:
                    result["Arm Height"] = m2.group(1)
            continue

        # ---------- Cushion ----------
        # "Number Of Cushions: 1 Cushion"
        if "number of cushions" in low:
            # try to take first number in the line
            m = re.search(r'([0-9]+)', line)
            if m:
                cushion_number = m.group(1)
            else:
                # or value after :
                if ":" in line:
                    cushion_number = line.split(":", 1)[1].strip()
            continue

        # "Standard Cushion: BLISS"
        if "standard cushion" in low:
            if ":" in line:
                cushion_standard = line.split(":", 1)[1].strip()
            else:
                # remove label
                cushion_standard = re.sub(r'(?i)standard cushion', "", line).strip()
            continue

        # ---------- Construction ----------
        # "Kd Construction: Yes"
        if "kd construction" in low and not result["Construction"]:
            if ":" in line:
                val = line.split(":", 1)[1].strip()
            else:
                val = re.sub(r'(?i)kd construction', "", line).strip()
            result["Construction"] = val
            continue

        # ---------- Content ----------
        # "Content: 100% POLYESTER"
        if low.startswith("content") and not result["Content"]:
            if ":" in line:
                val = line.split(":", 1)[1].strip()
            else:
                val = re.sub(r'(?i)content', "", line).strip()
            result["Content"] = val
            continue

        # ---------- Color ----------
        # "Color: Bark"
        if low.startswith("color") and not result["Color"]:
            if ":" in line:
                val = line.split(":", 1)[1].strip()
            else:
                val = re.sub(r'(?i)color', "", line).strip()
            result["Color"] = val
            continue

    # Cushion final value
    if cushion_number or cushion_standard:
        if cushion_number and cushion_standard:
            result["Cushion"] = f"{cushion_number}, {cushion_standard}"
        elif cushion_number:
            result["Cushion"] = cushion_number
        else:
            result["Cushion"] = cushion_standard

    return result

def main():
    print(f"🔎 Reading input: {INPUT_FILE}")

    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input file not found: {INPUT_FILE}")
        return

    df = pd.read_excel(INPUT_FILE)

    # ----- Product Family Id = Product Name -----
    if "Product Name" in df.columns:
        # Force Product Family Id to be exactly Product Name
        df["Product Family Id"] = df["Product Name"]
    else:
        df["Product Family Id"] = ""

    # ensure Specifications column exists
    if "Specifications" not in df.columns:
        df["Specifications"] = ""

    # prepare new columns with default empty
    new_cols = [
        "Weight",
        "Dimension",
        "Seat Height",
        "Seat Depth",
        "Arm Height",
        "Cushion",
        "Construction",
        "Content",
        "Color",
        "Width",
        "Depth",
        "Diameter",
        "Length",
        "Height",
    ]
    for col in new_cols:
        if col not in df.columns:
            df[col] = ""

    total = len(df)
    for idx in range(total):
        spec_text = safe_str(df.at[idx, "Specifications"])
        parsed = parse_specs_block_text(spec_text)

        for col in new_cols:
            df.at[idx, col] = parsed.get(col, "")

    # ----- Final column order -----
    desired_order = [
        "Product URL",
        "Image URL",
        "Product Name",
        "SKU",
        "Product Family Id",
        "Description",
        "Weight",
        "Width",
        "Depth",
        "Diameter",
        "Length",
        "Height",
        "Finish",
        "Seat Height",
        "Seat Depth",
        "Arm Height",
        "Cushion",
        "Construction",
        "Content",
        "Color",
    ]

    # keep only those that actually exist, then append remaining columns
    current_cols = list(df.columns)
    ordered_cols = []

    for col in desired_order:
        if col in df.columns:
            ordered_cols.append(col)

    for col in current_cols:
        if col not in ordered_cols:
            ordered_cols.append(col)

    df = df[ordered_cols]

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Final file saved: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
