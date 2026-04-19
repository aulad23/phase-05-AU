import requests
import pandas as pd
import re
from bs4 import BeautifulSoup

INPUT_FILE = "pierrefrey_carpets.xlsx"
OUTPUT_FILE = "pierrefrey_carpets_final.xlsx"

headers = {"User-Agent": "Mozilla/5.0"}

# ✅ NEW: batch save every N products
SAVE_EVERY = 10

# ================= HELPERS ================= #

def clean_text(t: str) -> str:
    if t is None:
        return ""
    return re.sub(r"\s+", " ", str(t)).strip()

def normalize_line(s: str) -> str:
    # IMPORTANT: do NOT remove ":" here (needed for inline label detection)
    return clean_text(s).lower().strip()

def extract_block_from_details(details_raw: str, label: str, known_labels: list) -> str:
    """
    Robust block extractor:
    - Label alone line: "Dimensions"
    - Inline label: "Construction : Traditional assembly" / "Seat: ..."
    - Multi-line values (Dimensions etc.)
    - Stops when next label starts (also detects inline labels)
    """
    if not details_raw:
        return ""

    lines = [ln.strip() for ln in re.split(r"[\r\n]+", details_raw) if ln.strip()]
    label_l = normalize_line(label)
    known_l = [normalize_line(x) for x in known_labels if x]

    def is_label_line(raw_line: str) -> str:
        rl = normalize_line(raw_line)
        for k in known_l:
            if rl == k:
                return k
            if rl.startswith(k) and (":" in raw_line[:len(k) + 6]):
                return k
        return ""

    for i, raw in enumerate(lines):
        rl = normalize_line(raw)

        if rl == label_l or (rl.startswith(label_l) and (":" in raw[:len(label_l) + 6])):
            vals = []

            # inline value: "Label: value"
            if ":" in raw:
                parts = raw.split(":", 1)
                if len(parts) == 2:
                    inline_val = parts[1].strip()
                    if inline_val:
                        vals.append(inline_val)

            j = i + 1
            while j < len(lines):
                if is_label_line(lines[j]):
                    break
                vals.append(lines[j])
                j += 1

            return "\n".join(vals).strip()

    return ""

def parse_com_keep_comma(block_text: str) -> str:
    """Fabric Qty: '2,17 yds' -> '2,17' (yds removed, comma kept)"""
    if not block_text:
        return ""
    t = clean_text(block_text)
    m = re.search(r"(\d+(?:[.,]\d+)?)", t)
    return m.group(1) if m else ""

def parse_numeric_keep_comma(block_text: str) -> str:
    """
    Generic numeric extractor (comma kept):
      '10.81 oz/ft' -> '10.81'
      '33,08 Lbs'   -> '33,08'
    """
    if not block_text:
        return ""
    t = clean_text(block_text)
    m = re.search(r"(\d+(?:[.,]\d+)?)", t)
    return m.group(1) if m else ""

def extract_inch_value(text: str) -> str:
    """
    Pick the number that is tied to 'inch' when present.
    Examples:
      '134 cm / 52,75 inch' -> '52,75'
      '0.39 inch' -> '0.39'
      '10.00 mm 0.39 inch' -> '0.39'
    Fallback: first numeric if no inch found.
    """
    if not text:
        return ""
    t = str(text)

    matches = re.findall(r"(\d+(?:[.,]\d+)?)\s*(?:\"|\binc?h(?:es)?\b)", t, flags=re.IGNORECASE)
    if matches:
        return matches[-1]

    return parse_numeric_keep_comma(t)

def get_inch_line(dim_text: str) -> str:
    """Dimensions multi-line text থেকে inch/inches/" আছে এমন লাইন বের করে"""
    if not dim_text:
        return ""
    lines = [x.strip() for x in re.split(r"[\r\n]+", str(dim_text)) if x.strip()]
    for ln in lines:
        if re.search(r'\b(inch|inches)\b|"', ln, flags=re.IGNORECASE):
            return ln
    return ""

def parse_inch_dimensions_from_text(dim_text: str):
    """
    ONLY inch line থেকে parse করবে:
      L -> Length
      D বা P -> Depth
      W -> Width
      H -> Height
      Dia/DIA/Diameter -> Diameter
    """
    out = {"width": "", "depth": "", "diameter": "", "length": "", "height": ""}

    inch_line = get_inch_line(dim_text)
    if not inch_line:
        return out

    t = clean_text(inch_line)
    t = t.replace("×", "x").replace("X", "x")

    t = re.sub(r"\bDIA\.?\b", "DIA", t, flags=re.IGNORECASE)
    t = re.sub(r"\bDiameter\b", "DIA", t, flags=re.IGNORECASE)

    pattern = re.compile(r"\b(L|P|D|W|H|DIA)\b\s*:?\s*(\d+(?:[.,]\d+)?)", flags=re.IGNORECASE)

    for lab, val in pattern.findall(t):
        lab = lab.upper()
        if lab == "L":
            out["length"] = val
        elif lab in ("P", "D"):
            out["depth"] = val
        elif lab == "W":
            out["width"] = val
        elif lab == "H":
            out["height"] = val
        elif lab == "DIA":
            out["diameter"] = val

    return out

# ================= MAIN ================= #

df = pd.read_excel(INPUT_FILE)
final_rows = []

total_products = len(df)  # ✅ NEW: total count

KNOWN_LABELS = [
    "Dimensions", "Fabric Qty", "Fabric Quantity", "COM", "Com",
    "Weight", "Total Weight",
    "Seat", "Back", "Cushion",
    "Frame", "Construction", "Finish", "Finishes", "Frame Finishes",
    "Comfort", "Repeat", "Thickness",
    "Width", "Total Width",
    "Total Height",
]

OUTPUT_COLS = [
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
    "Total Height",
    "Finish",
    "Com",
    "Seat",
    "Back",
    "Cushion",
    "Construction",
    "Repeat",
    "Comfort",
    "Thickness",
]

processed_ok = 0  # ✅ NEW: successful count

for idx, row in df.iterrows():
    current_no = idx + 1  # ✅ NEW: progress number (1-based)

    product_url = row["Product URL"]
    image_url = row["Image URL"]
    product_name = row["Product Name"]
    sku = row["SKU"]

    # ✅ NEW: show progress in console
    print(f"[{current_no}/{total_products}] Scraping: {product_url}")

    try:
        res = requests.get(product_url, headers=headers, timeout=20)
        res.raise_for_status()
    except Exception as e:
        print(f"❌ Failed: {e}")
        continue

    soup = BeautifulSoup(res.text, "html.parser")

    # DESCRIPTION
    description = ""
    desc_div = soup.find("div", class_="productInfos__description")
    if desc_div:
        description = desc_div.get_text(" ", strip=True)

    popin_div = soup.find("div", class_="popinText__container")
    if popin_div:
        wysiwyg = popin_div.find("div", class_="wysiwyg")
        if wysiwyg:
            description = wysiwyg.get_text(" ", strip=True)

    description = clean_text(description)

    # DETAILS RAW (from specTechInfos)
    spec_block = soup.find("div", class_="specTechInfos")
    details_raw = spec_block.get_text(separator="\n", strip=True) if spec_block else ""

    # Dimensions (needed for inch parsing)
    dimensions = extract_block_from_details(details_raw, "Dimensions", KNOWN_LABELS)
    dim_parts = parse_inch_dimensions_from_text(dimensions)

    # Weight: sometimes comes as Total Weight (e.g., 10.81 oz/ft)
    weight_block = extract_block_from_details(details_raw, "Weight", KNOWN_LABELS)
    weight_value = parse_numeric_keep_comma(weight_block)

    if not weight_value:
        total_weight_block = extract_block_from_details(details_raw, "Total Weight", KNOWN_LABELS)
        weight_value = parse_numeric_keep_comma(total_weight_block)

    # Total Height column (take inch value if present)
    total_height_block = extract_block_from_details(details_raw, "Total Height", KNOWN_LABELS)
    total_height_value = extract_inch_value(total_height_block)

    # Width sometimes present in details like: "134 cm / 52,75 inch" (or Total Width)
    width_details_block = extract_block_from_details(details_raw, "Width", KNOWN_LABELS)
    if not width_details_block:
        width_details_block = extract_block_from_details(details_raw, "Total Width", KNOWN_LABELS)

    width_from_details_inch = extract_inch_value(width_details_block)

    # If details width exists, override Width column
    width_final = width_from_details_inch if width_from_details_inch else dim_parts["width"]

    # COM
    fabric_block = extract_block_from_details(details_raw, "Fabric Qty", KNOWN_LABELS)
    if not fabric_block:
        fabric_block = extract_block_from_details(details_raw, "Fabric Quantity", KNOWN_LABELS)
    com_value = parse_com_keep_comma(fabric_block)

    # Finish / Construction / Repeat / Comfort / Thickness
    finish = extract_block_from_details(details_raw, "Frame Finishes", KNOWN_LABELS) or \
             extract_block_from_details(details_raw, "Finish", KNOWN_LABELS) or \
             extract_block_from_details(details_raw, "Finishes", KNOWN_LABELS)

    construction = extract_block_from_details(details_raw, "Construction", KNOWN_LABELS)
    repeat = extract_block_from_details(details_raw, "Repeat", KNOWN_LABELS)
    comfort = extract_block_from_details(details_raw, "Comfort", KNOWN_LABELS)
    thickness = extract_block_from_details(details_raw, "Thickness", KNOWN_LABELS)

    # Seat / Back / Cushion
    seat = extract_block_from_details(details_raw, "Seat", KNOWN_LABELS)
    back = extract_block_from_details(details_raw, "Back", KNOWN_LABELS)
    cushion = extract_block_from_details(details_raw, "Cushion", KNOWN_LABELS)

    final_rows.append({
        "Product URL": product_url,
        "Image URL": image_url,
        "Product Name": product_name,
        "SKU": sku,
        "Product Family Id": product_name,
        "Description": description,
        "Weight": weight_value,
        "Width": width_final,
        "Depth": dim_parts["depth"],
        "Diameter": dim_parts["diameter"],
        "Length": dim_parts["length"],
        "Height": dim_parts["height"],
        "Total Height": total_height_value,
        "Finish": finish,
        "Com": com_value,
        "Seat": seat,
        "Back": back,
        "Cushion": cushion,
        "Construction": construction,
        "Repeat": repeat,
        "Comfort": comfort,
        "Thickness": thickness,
    })

    processed_ok += 1  # ✅ NEW

    # ✅ NEW: batch save every 10 successful rows (overwrite output)
    if processed_ok % SAVE_EVERY == 0:
        final_df = pd.DataFrame(final_rows)
        for c in OUTPUT_COLS:
            if c not in final_df.columns:
                final_df[c] = ""
        final_df = final_df[OUTPUT_COLS]
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"💾 Saved batch: {processed_ok} done (out of total {total_products}) -> {OUTPUT_FILE}")

# ---------------- FINAL SAVE OUTPUT ---------------- #

final_df = pd.DataFrame(final_rows)

for c in OUTPUT_COLS:
    if c not in final_df.columns:
        final_df[c] = ""

final_df = final_df[OUTPUT_COLS]
final_df.to_excel(OUTPUT_FILE, index=False)

print(f"✅ Completed successfully: {OUTPUT_FILE}")
print(f"✅ Total successful scraped: {processed_ok} / {total_products}")
