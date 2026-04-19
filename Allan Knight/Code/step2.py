# ============================================================
# Allan Knight — STEP-2 (FINAL FULL + Column Order + SKU from Product Page)
# ✅ INPUT : allan_knight_cocktail_tables.xlsx
# ✅ OUTPUT: allan_knight_cocktail_tables_step2.xlsx
#
# OUTPUT COLUMN ORDER (as requested):
# Product URL, Image URL, Product Name, SKU, Product Family Id, Description,
# Weight, Width, Depth, Diameter, Height, COM, Shade Details, Seat Height, Arm Height, Wattage, Dimension
# ============================================================

import os
import re
import time
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# =========================
# CONFIG
# =========================
BASE_URL = "https://allan-knight.com"

INPUT_FILE = r"D:\phase-05 (AU)\Allan Knight\Code\AllanKnight_Trays.xlsx"
OUTPUT_FILE = r"D:\phase-05 (AU)\Allan Knight\Code\AllanKnight_Trays_Final .xlsx"

TIMEOUT = 25
SLEEP_BETWEEN = 0.25

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
}

# =========================
# TEXT HELPERS
# =========================
def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = txt.replace("\r", "\n")
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r"\n{2,}", "\n", txt)
    return txt.strip()

def _clean_desc(desc: str) -> str:
    return clean_text(desc or "")

def product_family_from_name(product_name: str) -> str:
    if not product_name:
        return ""
    parts = [p.strip() for p in product_name.split(",") if p.strip()]
    if len(parts) >= 2:
        return ", ".join(parts[:-1]).strip()
    return product_name.strip()

# =========================
# FRACTION → DECIMAL
# =========================
def normalize_fraction_to_decimal(val: str) -> str:
    if not val:
        return ""
    val = val.strip()

    fraction_map = {"¼": 0.25, "½": 0.5, "¾": 0.75}

    for f, dec in fraction_map.items():
        if f in val:
            base_part = val.replace(f, "").strip()
            base = float(base_part) if base_part else 0.0
            return str(round(base + dec, 2))

    m = re.match(r"^\s*(\d+)?\s*(\d+)\s*/\s*(\d+)\s*$", val)
    if m:
        whole = float(m.group(1)) if m.group(1) else 0.0
        num = float(m.group(2))
        den = float(m.group(3))
        if den != 0:
            return str(round(whole + (num / den), 2))

    try:
        return str(float(val))
    except Exception:
        return val

# =========================
# NETWORK
# =========================
def fetch_html(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return r.text

# =========================
# STOP MARKERS (template-B)
# =========================
def is_stop_marker_tag(tag) -> bool:
    if not tag or not getattr(tag, "name", None):
        return False

    name = tag.name.lower()

    if name == "a":
        href = (tag.get("href") or "").lower()
        txt = clean_text(tag.get_text(" ", strip=True)).lower()
        if ".pdf" in href or "download pdf" in txt:
            return True
        if "purchase or inquire" in txt:
            return True

    if name in ("h2", "h3"):
        t = clean_text(tag.get_text(" ", strip=True)).lower()
        if "suggested items" in t:
            return True

    return False

# =========================
# DESCRIPTION EXTRACTION (TABLE-AWARE)
# =========================
def extract_table_text(container) -> str:
    if not container:
        return ""
    table = container.find("table") if hasattr(container, "find") else None
    if not table:
        return ""
    td = table.find("td")
    if td:
        return clean_text(td.get_text("\n", strip=True))
    return clean_text(table.get_text("\n", strip=True))

def extract_desc_from_holder(holder) -> str:
    if not holder:
        return ""

    desc_block = None
    for d in holder.find_all("div", recursive=False):
        if d.find("table") or d.find("p"):
            desc_block = d
            break

    if not desc_block:
        t = extract_table_text(holder)
        if t:
            return t
        p = holder.select_one("p")
        return clean_text(p.get_text("\n", strip=True)) if p else ""

    t = extract_table_text(desc_block)
    if t:
        return t

    return clean_text(desc_block.get_text("\n", strip=True))

def extract_description_after_h1(h1_tag) -> str:
    if not h1_tag:
        return ""

    lines = []
    for el in h1_tag.next_elements:
        if hasattr(el, "name") and el.name:
            if is_stop_marker_tag(el):
                break

            if el.name.lower() == "table":
                t = extract_table_text(el)
                if t:
                    return t

            if el.name.lower() == "div" and el.find("table"):
                t = extract_table_text(el)
                if t:
                    return t

            if el.name.lower() in ("p", "div", "span", "strong", "li"):
                t = clean_text(el.get_text("\n", strip=True))
                if not t:
                    continue
                if "download pdf" in t.lower():
                    break
                if t not in lines:
                    lines.append(t)

        if len(lines) >= 12:
            break

    return clean_text("\n".join(lines))

# =========================
# ✅ NEW: SKU extraction from product page
# =========================
def extract_sku_from_soup(soup: BeautifulSoup) -> str:
    holder = soup.select_one("div.product-desc-holder")
    if holder:
        sku_tag = holder.select_one("strong.category")
        if sku_tag:
            sku = clean_text(sku_tag.get_text(" ", strip=True))
            if sku and sku.lower() not in ("", "category"):
                return sku

        for t in holder.select("strong.category"):
            sku = clean_text(t.get_text(" ", strip=True))
            if sku:
                return sku

    sku_hidden = soup.select_one("strong.sku.hidden")
    if sku_hidden:
        sku = clean_text(sku_hidden.get_text(" ", strip=True))
        if sku:
            return sku

    any_cat = soup.select_one("strong.category")
    if any_cat:
        sku = clean_text(any_cat.get_text(" ", strip=True))
        if sku:
            return sku

    return ""

def extract_step2_fields(product_url: str):
    html = fetch_html(product_url)
    soup = BeautifulSoup(html, "html.parser")

    sku = extract_sku_from_soup(soup)

    holder = soup.select_one("div.product-desc-holder")
    if holder:
        h1 = holder.select_one("h1")
        product_name = clean_text(h1.get_text(" ", strip=True)) if h1 else ""
        family_id = product_family_from_name(product_name)
        description = extract_desc_from_holder(holder)
        return product_name, sku, family_id, description

    h1 = soup.select_one("h1")
    product_name = clean_text(h1.get_text(" ", strip=True)) if h1 else ""
    family_id = product_family_from_name(product_name)

    description = extract_description_after_h1(h1)

    if not description:
        t = extract_table_text(soup)
        if t:
            description = t
        else:
            p = soup.select_one("p")
            description = clean_text(p.get_text("\n", strip=True)) if p else ""

    if description and product_name and description.strip() == product_name.strip():
        description = ""

    return product_name, sku, family_id, description

# =========================
# PARSE DESCRIPTION → COLUMNS
# =========================
def extract_dimension(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""

    num = r"(?:\d+(?:\.\d+)?(?:\s*(?:¼|½|¾))?(?:\s+\d/\d)?|\d+/\d)"

    dia_h = re.findall(rf"(?i)\b({num})\s*dia\s*x\s*({num})\s*h\b", d)
    if dia_h:
        a, b = dia_h[-1]
        return f"{a.strip()} dia x {b.strip()} h"

    xyz_h = re.findall(rf"(?i)\b({num})\s*x\s*({num})\s*x\s*({num})\s*h\b", d)
    if xyz_h:
        a, b, c = xyz_h[-1]
        return f"{a.strip()} x {b.strip()} x {c.strip()} h"

    dia_d = re.findall(rf"(?i)\b({num})\s*dia\s*x\s*({num})\s*d\b", d)
    if dia_d:
        a, b = dia_d[-1]
        return f"{a.strip()} dia x {b.strip()} d"

    h_dia = re.findall(rf"(?i)\b({num})\s*h\s*x\s*({num})\s*dia\b", d)
    if h_dia:
        h_val, dia_val = h_dia[-1]
        return f"{dia_val.strip()} dia x {h_val.strip()} h"

    dia_num = re.findall(rf"(?i)\b({num})\s*dia\s*x\s*({num})\s*(?:\b|$)", d)
    if dia_num:
        dia_val, other_val = dia_num[-1]
        return f"{dia_val.strip()} dia x {other_val.strip()} h"

    return ""

def extract_com(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    m = re.search(r"(?i)\bCOM\b[^0-9]{0,15}(\d+(?:\.\d+)?)\s*(?:yds|yd)\b", d)
    return m.group(1).strip() if m else ""

# =========================
# ✅ Shade Details extraction
# Handles all patterns:
#   shade 20 x 21 x 11, white        → 20 x 21 x 11
#   shade 17 t, 18 b, 10 s           → 17 t, 18 b, 10 s
#   silk shade 15 t, 16 b, 9 s, white → 15 t, 16 b, 9 s
#   oval shade 9 ½ x 17 t, 10 x 18 b, 10s → 9 ½ x 17 t, 10 x 18 b, 10s
#   rectangular shade 6 x 9 t, ...    → 6 x 9 t, ...
#   round silk shade, 15 t, ...       → 15 t, ...
# =========================
def extract_shade_details(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    m = re.search(r"(?i)(?:round\s+silk|silk|oval|rectangular)?\s*shade[,]?\s*(.+)", d)
    if not m:
        return ""
    raw = m.group(1).strip()
    # Remove trailing color info
    raw = re.sub(r",?\s*(?:white|oyster|cream|ivory|black)(?:/(?:white|oyster|cream|ivory|black))?\s*$", "", raw, flags=re.IGNORECASE)
    # Remove trailing "View this product..." text
    raw = re.sub(r"\s*View\s+this.*$", "", raw, flags=re.IGNORECASE)
    raw = raw.strip().rstrip(",").strip()
    return raw

# =========================
# ✅ OAH (Overall Height) extraction
# e.g. "36 ½ oah" → "36.5"
# Falls back into Height when dimension-based height is empty
# =========================
def extract_oah(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    num = r"(\d+(?:\.\d+)?(?:\s*(?:¼|½|¾))?(?:\s+\d/\d)?|\d+/\d)"
    m = re.search(rf"(?i)\b{num}\s*oah\b", d)
    if m:
        return normalize_fraction_to_decimal(m.group(1).strip())
    return ""

def extract_seat_height(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    m = re.search(r"(?i)\b(\d+(?:\.\d+)?)\s*(?:\"|in)?\s*sh\b", d)
    if m:
        return m.group(1).strip()
    m = re.search(r"(?i)\bsh\s*(\d+(?:\.\d+)?)\b", d)
    return m.group(1).strip() if m else ""

def extract_arm_height(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    m = re.search(r"(?i)\b(\d+(?:\.\d+)?)\s*(?:\"|in)?\s*ah\b", d)
    if m:
        return m.group(1).strip()
    m = re.search(r"(?i)\bah\s*(\d+(?:\.\d+)?)\b", d)
    return m.group(1).strip() if m else ""

def extract_wattage(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    m = re.search(r"(?i)\b(\d+(?:\.\d+)?)\s*w\s*max\b", d)
    if m:
        return m.group(1).strip()
    m = re.search(r"(?i)\b(\d+(?:\.\d+)?)w\s*max\b", d)
    if m:
        return m.group(1).strip()
    m = re.search(r"(?i)\bw\s*max\s*(\d+(?:\.\d+)?)\b", d)
    return m.group(1).strip() if m else ""

def extract_weight(desc: str) -> str:
    d = _clean_desc(desc)
    if not d:
        return ""
    m = re.search(r"(?i)\b(\d+(?:\.\d+)?)\s*(lb|lbs|kg)\b", d)
    return f"{m.group(1).strip()} {m.group(2).lower()}" if m else ""

# =========================
# SPLIT DIMENSION → Width/Depth/Diameter/Height
# =========================
def split_dimension(dim: str):
    dim = clean_text(dim)
    if not dim:
        return "", "", "", ""

    num = r"(?:\d+(?:\.\d+)?(?:\s*(?:¼|½|¾))?(?:\s+\d/\d)?|\d+/\d)"

    m = re.search(rf"(?i)\b({num})\s*dia\s*x\s*({num})\s*h\b", dim)
    if m:
        diameter = normalize_fraction_to_decimal(m.group(1).strip())
        height   = normalize_fraction_to_decimal(m.group(2).strip())
        return "", "", diameter, height

    m = re.search(rf"(?i)\b({num})\s*dia\s*x\s*({num})\s*d\b", dim)
    if m:
        diameter = normalize_fraction_to_decimal(m.group(1).strip())
        depth    = normalize_fraction_to_decimal(m.group(2).strip())
        return "", depth, diameter, ""

    m = re.search(rf"(?i)\b({num})\s*dia\s*x\s*({num})\b", dim)
    if m:
        diameter = normalize_fraction_to_decimal(m.group(1).strip())
        depth    = normalize_fraction_to_decimal(m.group(2).strip())
        return "", depth, diameter, ""

    m = re.search(rf"(?i)\b({num})\s*x\s*({num})\s*x\s*({num})\s*h\b", dim)
    if m:
        width  = normalize_fraction_to_decimal(m.group(1).strip())
        depth  = normalize_fraction_to_decimal(m.group(2).strip())
        height = normalize_fraction_to_decimal(m.group(3).strip())
        return width, depth, "", height

    m = re.search(rf"(?i)\b({num})\s*x\s*({num})\s*h\b", dim)
    if m:
        width  = normalize_fraction_to_decimal(m.group(1).strip())
        height = normalize_fraction_to_decimal(m.group(2).strip())
        return width, "", "", height

    return "", "", "", ""

# =========================
# MAIN
# =========================
def main():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)

    if "Product URL" not in df.columns:
        raise ValueError("Input Excel must have a 'Product URL' column.")

    base_cols = ["Product URL", "Image URL", "Product Name", "SKU"]
    for c in base_cols:
        if c not in df.columns:
            df[c] = ""

    needed = [
        "Product Family Id", "Description", "Weight", "Dimension",
        "COM", "Shade Details", "Seat Height", "Arm Height", "Wattage",
        "Width", "Depth", "Diameter", "Height"
    ]
    for c in needed:
        if c not in df.columns:
            df[c] = ""

    total = len(df)
    print(f"✅ Loaded: {total} rows")

    for i, (idx, row) in enumerate(df.iterrows(), start=1):
        url = str(row.get("Product URL", "")).strip()
        if not url:
            continue
        if url.startswith("/"):
            url = urljoin(BASE_URL, url)

        try:
            pname, sku, fam, desc = extract_step2_fields(url)

            if sku:
                df.at[idx, "SKU"] = sku

            if not str(df.at[idx, "Product Name"]).strip():
                df.at[idx, "Product Name"] = pname

            df.at[idx, "Product Family Id"] = fam
            df.at[idx, "Description"] = desc

            df.at[idx, "Weight"] = extract_weight(desc)

            dim = extract_dimension(desc)
            df.at[idx, "Dimension"] = dim

            df.at[idx, "COM"] = extract_com(desc)
            df.at[idx, "Shade Details"] = extract_shade_details(desc)
            df.at[idx, "Seat Height"] = extract_seat_height(desc)
            df.at[idx, "Arm Height"] = extract_arm_height(desc)
            df.at[idx, "Wattage"] = extract_wattage(desc)

            w, d, dia, h = split_dimension(dim)
            # ✅ If no height from dimension, fallback to OAH
            if not h:
                h = extract_oah(desc)
            df.at[idx, "Width"] = w
            df.at[idx, "Depth"] = d
            df.at[idx, "Diameter"] = dia
            df.at[idx, "Height"] = h

        except Exception as e:
            print(f"❌ Row {i}/{total} failed: {url} | {e}")

        if i % 20 == 0:
            print(f"  ...processed {i}/{total}")

        time.sleep(SLEEP_BETWEEN)

    # ✅ FINAL COLUMN ORDER
    final_order = [
        "Product URL", "Image URL", "Product Name", "SKU",
        "Product Family Id", "Description",
        "Weight", "Width", "Depth", "Diameter", "Height",
        "COM", "Shade Details", "Seat Height", "Arm Height", "Wattage",
        "Dimension"
    ]

    extras = [c for c in df.columns if c not in final_order]
    df = df[[*final_order, *extras]]

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()