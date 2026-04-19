"""
Hennepin Made - Step 2 FINAL (HYBRID: JSON API + Page Scrape)
COMPLETE FIX WITH:
- Smart quote handling (fancy quotes from Excel)
- All dimension patterns (W, H, D, L, Dia - any case)
- Proper parsing that fills Width, Height, Depth, Diameter, Length columns
- ADDED: Fallback parser for ProductMeta__Description style pages (bulb products)
  Maps: BASE→Base, LUMENS→Lumens, TEMPERATURE→Color Temp, VOLTAGE→Input Voltage,
        WATTAGE→Wattage, ASSOCIATED COLLECTION→Associated Collection
"""

import requests
import time
import re
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook

# =========================
# CONFIG
# =========================
INPUT_XLSX  = "hennepinmade_Bulbs.xlsx"
OUTPUT_XLSX = "hennepinmade_Bulbs_Final.xlsx"
MANUFACTURER = "Hennepin Made"
BASE_URL = "https://hennepinmade.com"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}


def num_only(s):
    if not s:
        return ""
    m = re.search(r'([\d]+(?:\.\d+)?)', str(s))
    return m.group(1) if m else ""


def clean_price(p):
    if not p:
        return ""
    return str(p).replace("$", "").replace(",", "").strip()


# =========================
# DIMENSION PARSER - COMPLETE & TESTED
# =========================
def parse_dimensions(dim_source):
    """
    Parse dimensions from ANY format.
    Handles: W, H, D, L, Dia (case insensitive, smart quotes)

    Examples:
    - "12\" dia x 2.9\"" → {Diameter: 12, Height: 2.9}
    - "15\" W x 13.5\" H" → {Width: 15, Height: 13.5}
    - "4\" x 13\"" → {Width: 4, Height: 13}
    - "5.5\" W x 18.5\" H x 3.75\" D" → {Width: 5.5, Height: 18.5, Depth: 3.75}
    """

    specs = {}

    if not dim_source:
        return specs

    dim_source = str(dim_source).strip()

    # ===== CRITICAL: Replace ALL smart quotes with regular quotes =====
    dim_source = dim_source.replace(chr(0x201d), '"')
    dim_source = dim_source.replace(chr(0x201c), '"')
    dim_source = dim_source.replace(chr(0x2019), "'")
    dim_source = dim_source.replace(chr(0x2018), "'")
    dim_source = dim_source.replace(chr(0x2033), '"')

    dim_source = re.sub(r'\n', ' ', dim_source)
    dim_source = re.sub(r'(\d)\s*"\s*', r'\1" ', dim_source)
    dim_source = re.sub(r'(\d)\s+(["\u2033])', r'\1\2', dim_source)
    dim_source = re.sub(r'\s+', ' ', dim_source).strip()

    if not dim_source or dim_source == '':
        return specs

    SEP = r'[\s]*[xX×][\s]*'

    # ===== PATTERN 1: THREE labeled dimensions =====
    m = re.search(r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]' + SEP + r'([\d.]+)["\u2033]?\s*[Dd]', dim_source, re.I)
    if m:
        specs["Width"] = m.group(1); specs["Height"] = m.group(2); specs["Depth"] = m.group(3)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Dd]' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]', dim_source, re.I)
    if m:
        specs["Width"] = m.group(1); specs["Depth"] = m.group(2); specs["Height"] = m.group(3)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]' + SEP + r'([\d.]+)["\u2033]?\s*[Ll]', dim_source, re.I)
    if m:
        specs["Width"] = m.group(1); specs["Height"] = m.group(2); specs["Length"] = m.group(3)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Ll]' + SEP + r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]', dim_source, re.I)
    if m:
        specs["Length"] = m.group(1); specs["Width"] = m.group(2); specs["Height"] = m.group(3)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Dd]' + SEP + r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]', dim_source, re.I)
    if m:
        specs["Depth"] = m.group(1); specs["Width"] = m.group(2); specs["Height"] = m.group(3)
        return specs

    # ===== PATTERN 2: TWO labeled dimensions =====
    m = re.search(r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]', dim_source, re.I)
    if m:
        specs["Width"] = m.group(1); specs["Height"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Dd]', dim_source, re.I)
    if m:
        specs["Width"] = m.group(1); specs["Depth"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Ww]' + SEP + r'([\d.]+)["\u2033]?\s*[Ll]', dim_source, re.I)
    if m:
        specs["Width"] = m.group(1); specs["Length"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Hh]' + SEP + r'([\d.]+)["\u2033]?\s*[Dd]', dim_source, re.I)
    if m:
        specs["Height"] = m.group(1); specs["Depth"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Hh]' + SEP + r'([\d.]+)["\u2033]?\s*[Ll]', dim_source, re.I)
    if m:
        specs["Height"] = m.group(1); specs["Length"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*[Dd]' + SEP + r'([\d.]+)["\u2033]?\s*[Ll]', dim_source, re.I)
    if m:
        specs["Depth"] = m.group(1); specs["Length"] = m.group(2)
        return specs

    # ===== PATTERN 3: DIAMETER patterns =====
    m = re.search(r'([\d.]+)["\u2033]?\s*(?:[Dd]ia|[Dd]iameter)' + SEP + r'([\d.]+)["\u2033]?\s*[Hh]', dim_source, re.I)
    if m:
        specs["Diameter"] = m.group(1); specs["Height"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*(?:[Dd]ia|[Dd]iameter)' + SEP + r'([\d.]+)["\u2033]?\s*[Dd]', dim_source, re.I)
    if m:
        specs["Diameter"] = m.group(1); specs["Depth"] = m.group(2)
        return specs

    m = re.search(r'([\d.]+)["\u2033]?\s*(?:[Dd]ia|[Dd]iameter)' + SEP + r'([\d.]+)["\u2033]', dim_source, re.I)
    if m:
        specs["Diameter"] = m.group(1); specs["Height"] = m.group(2)
        return specs

    # ===== PATTERN 4: Unlabeled dimensions =====
    unlabeled = re.findall(r'([\d.]+)["\u2033]', dim_source)

    if len(unlabeled) >= 3:
        specs["Width"] = unlabeled[0]; specs["Height"] = unlabeled[1]; specs["Depth"] = unlabeled[2]
        return specs
    elif len(unlabeled) == 2:
        specs["Width"] = unlabeled[0]; specs["Height"] = unlabeled[1]
        return specs
    elif len(unlabeled) == 1:
        if re.search(r'(?:[Dd]ia|[Dd]iameter)', dim_source, re.I):
            specs["Diameter"] = unlabeled[0]
        elif re.search(r'[Ww]\b', dim_source, re.I):
            specs["Width"] = unlabeled[0]
        elif re.search(r'[Hh]\b', dim_source, re.I):
            specs["Height"] = unlabeled[0]
        elif re.search(r'[Dd]\b', dim_source, re.I):
            specs["Depth"] = unlabeled[0]
        elif re.search(r'[Ll]\b', dim_source, re.I):
            specs["Length"] = unlabeled[0]
        else:
            specs["Width"] = unlabeled[0]
        return specs

    return specs


# =========================
# FALLBACK PARSER: ProductMeta__Description style (e.g. bulb pages)
# Handles bold labels: BASE, LUMENS, TEMPERATURE, VOLTAGE, WATTAGE,
#                      ASSOCIATED COLLECTION, etc.
# =========================
def parse_product_meta_description(soup):
    """
    Parses specs from the .ProductMeta__Description > .Rte div.
    Used for pages that don't have a "Technical Specs" section (e.g. bulb products).

    Label mapping:
        BASE                → Base
        LUMENS              → Lumens
        TEMPERATURE         → Color Temp
        VOLTAGE             → Input Voltage
        WATTAGE             → Wattage
        ASSOCIATED COLLECTION → Associated Collection
    """
    LABEL_MAP = {
        "BASE":                 "Base",
        "LUMENS":               "Lumens",
        "TEMPERATURE":          "Color Temp",
        "VOLTAGE":              "Input Voltage",
        "WATTAGE":              "Wattage",
        "ASSOCIATED COLLECTION": "Associated Collection",
    }

    specs = {}
    description_extra = []

    rte_div = soup.select_one(".ProductMeta__Description .Rte")
    if not rte_div:
        return specs, ""

    for p_tag in rte_div.find_all("p"):
        strong = p_tag.find("strong")
        if strong:
            raw_label = strong.get_text(strip=True).rstrip(":").upper()
            # Remove the <strong> so we can get the value text cleanly
            strong.extract()
            value = p_tag.get_text(separator=" ", strip=True).strip()

            mapped_key = LABEL_MAP.get(raw_label)
            if mapped_key:
                specs[mapped_key] = value
            else:
                # Unknown bold label — store as-is (title-cased) so nothing is lost
                title_label = raw_label.title()
                specs[title_label] = value
        else:
            # Plain paragraph with no bold label = description text or dimmable note
            text = p_tag.get_text(strip=True)
            if text:
                description_extra.append(text)

    description = " ".join(description_extra).strip()
    return specs, description


# =========================
# PAGE HTML PARSER
# =========================
def parse_page_specs(page_html):
    if not page_html:
        return {}, ""

    soup = BeautifulSoup(page_html, "html.parser")
    full_text = soup.get_text(separator="\n", strip=True)

    # --- DESCRIPTION ---
    description = ""
    desc_match = re.search(
        r'(?:Made to order[.\s]*Lead time[^.]*\.\s*(?:Lookbook\s*)?)(.*?)(?=\s*(?:Share\s*)?Technical Specs)',
        full_text, re.DOTALL | re.I
    )
    if desc_match:
        description = desc_match.group(1).strip()
    else:
        desc_match2 = re.search(
            r'Lookbook\s+(.*?)(?=\s*Share\s*Technical Specs)',
            full_text, re.DOTALL | re.I
        )
        if desc_match2:
            description = desc_match2.group(1).strip()

    description = re.sub(r'\s+', ' ', description).strip()

    # --- TECH SPECS ---
    specs_text = ""
    specs_match = re.search(
        r'Technical Specs\s*(.*?)(?=\s*Downloads\b)',
        full_text, re.DOTALL | re.I
    )
    if specs_match:
        specs_text = specs_match.group(1).strip()
    else:
        specs_match2 = re.search(r'Technical Specs\s*(.*?)(?=\s*(?:Join the trade|Resources|More from))',
                                  full_text, re.DOTALL | re.I)
        if specs_match2:
            specs_text = specs_match2.group(1).strip()

    # -------------------------------------------------------
    # FALLBACK: If no Technical Specs section found, try the
    # ProductMeta__Description bold-label pattern (bulb pages)
    # -------------------------------------------------------
    if not specs_text:
        meta_specs, meta_desc = parse_product_meta_description(soup)
        if meta_specs:
            # Merge description: prefer page-level description if already found
            if not description and meta_desc:
                description = meta_desc
            return meta_specs, description
        # Truly nothing found
        return {}, description

    # --- Parse sections (existing logic) ---
    sections = {}
    current_label = ""
    current_value = []

    for line in specs_text.split('\n'):
        line = line.strip()
        if not line:
            continue

        label_match = re.match(r'^([A-Z][A-Z\s\(\)\/&]+)$', line)
        if label_match:
            if current_label:
                sections[current_label] = '\n'.join(current_value).strip()
            current_label = label_match.group(1).strip()
            current_value = []
        else:
            merged = re.match(r'^([A-Z][A-Z\s\(\)\/&]+?)(\d.*)$', line)
            if merged:
                if current_label:
                    sections[current_label] = '\n'.join(current_value).strip()
                current_label = merged.group(1).strip()
                current_value = [merged.group(2).strip()]
            else:
                current_value.append(line)

    if current_label:
        sections[current_label] = '\n'.join(current_value).strip()

    specs = {}

    # DIMENSIONS
    dim_lights = sections.get("DIMENSIONS (LIGHTS)", "")
    dim_canopy = sections.get("DIMENSIONS (CANOPY)", "")
    dim_plain = sections.get("DIMENSIONS", "")

    dim_source = dim_plain or dim_lights or dim_canopy
    if dim_source:
        dim_specs = parse_dimensions(dim_source)
        specs.update(dim_specs)

    all_dim = dim_lights
    if dim_canopy:
        all_dim += (" | CANOPY: " + dim_canopy) if all_dim else dim_canopy
    if dim_plain:
        all_dim = dim_plain if not all_dim else all_dim

    specs["Dimensions"] = all_dim.replace('\n', ' | ')[:500] if all_dim else ""

    # LAMPING
    specs["Lamping"] = sections.get("LAMPING", "").replace('\n', ' | ')

    # LUMENS
    lum = sections.get("LUMENS", "")
    specs["Lumens"] = lum.split('\n')[0].strip() if lum else ""

    # COLOR TEMP
    ct = sections.get("COLOR TEMP", sections.get("COLOR TEMPERATURE", ""))
    specs["Color Temp"] = ct.split('\n')[0].strip() if ct else ""

    # INPUT VOLTAGE
    iv = sections.get("INPUT VOLTAGE", sections.get("VOLTAGE", ""))
    specs["Input Voltage"] = iv.split('\n')[0].strip() if iv else ""

    # MOUNTING
    mount = sections.get("MOUNTING", "")
    specs["Mounting Info"] = mount.replace('\n', ' ').strip()[:300]

    # CORD LENGTH
    specs["Cord Length"] = sections.get("CORD LENGTH", sections.get("CORD", "")).replace('\n', ' ').strip()

    # DROP ROD
    specs["Drop Rod"] = sections.get("DROP ROD", "").replace('\n', ' ').strip()

    # OVERALL HEIGHT
    oah_text = sections.get("OVERALL HEIGHT", "")
    oah_max = re.search(r'OAH\s*max[:\s]*([\d.]+)"?', oah_text, re.I)
    oah_min = re.search(r'OAH\s*min[:\s]*([\d.]+)"?', oah_text, re.I)
    specs["OAH Max"] = oah_max.group(1) if oah_max else ""
    specs["OAH Min"] = oah_min.group(1) if oah_min else ""

    # WEIGHT
    wt_text = sections.get("WEIGHT (CANOPY AND LIGHTS)",
              sections.get("WEIGHT (CANOPY & LIGHTS)",
              sections.get("WEIGHT", "")))
    wt_m = re.search(r'([\d.]+)\s*(?:lbs?|pounds?)', wt_text, re.I)
    specs["Weight"] = wt_m.group(1) if wt_m else num_only(wt_text)

    # MATERIAL
    specs["Material"] = sections.get("MATERIAL", sections.get("MATERIALS", "")).replace('\n', ' ').strip()

    # CANOPY
    specs["Canopy"] = sections.get("CANOPY", "").replace('\n', ' ').strip()

    # SOCKET
    specs["Socket"] = sections.get("SOCKET", sections.get("SOCKET TYPE", "")).replace('\n', ' ').strip()

    # WATTAGE
    wattage_text = sections.get("WATTAGE", "")
    wt_match = re.search(r'(\d+\s*W(?:att)?)', wattage_text, re.I)
    specs["Wattage"] = wt_match.group(1) if wt_match else wattage_text.replace('\n', ' ').strip()

    # SHADE DETAILS
    shade = sections.get("SHADE DIMENSIONS", sections.get("SHADE", ""))
    specs["Shade Details"] = shade.replace('\n', ' ').strip()[:200] if shade else ""

    # BACKPLATE
    specs["Backplate"] = sections.get("BACKPLATE", sections.get("WALL PLATE", "")).replace('\n', ' ').strip()

    # DOWNLOADS
    downloads = {}
    dl_match = re.search(r'Downloads\s*(.*?)(?=\s*(?:Meet |More from|Join the trade|Trade Program|Resources))',
                          full_text, re.DOTALL | re.I)
    if dl_match:
        dl_section = None
        for tag in soup.find_all(string=re.compile(r'Downloads', re.I)):
            parent = tag.find_parent(['div', 'section', 'h2', 'h3'])
            if parent:
                dl_section = parent.find_parent(['div', 'section'])
                break

        if dl_section:
            for a_tag in dl_section.find_all("a", href=True):
                href = a_tag["href"]
                link_text = a_tag.get_text(strip=True)
                if link_text and (".pdf" in href.lower() or ".step" in href.lower() or
                                  ".dwg" in href.lower() or ".skp" in href.lower() or
                                  ".rfa" in href.lower()):
                    downloads[link_text] = href

    if not downloads:
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            link_text = a_tag.get_text(strip=True)
            if link_text and ".pdf" in href.lower() and "cdn.shopify" in href:
                name_lower = link_text.lower()
                if any(k in name_lower for k in ["spec", "install", "lookbook", "tear"]):
                    downloads[link_text] = href

    specs["Downloads"] = downloads

    return specs, description


# =========================
# PRODUCT FETCHER
# =========================
def get_product_json(handle):
    json_url = f"{BASE_URL}/products/{handle}.json"
    try:
        resp = requests.get(json_url, headers=HEADERS, timeout=30)
        if resp.status_code == 200:
            return resp.json().get("product", {})
    except Exception as e:
        print(f"  ❌ JSON error: {e}")
    return None


def get_page_html(handle):
    page_url = f"{BASE_URL}/products/{handle}"
    try:
        resp = requests.get(page_url, headers=HEADERS, timeout=30)
        if resp.status_code == 200:
            return resp.text
    except Exception as e:
        print(f"  ❌ Page fetch error: {e}")
    return None


def build_variant_image_map(product_data):
    images = product_data.get("images", [])
    variant_img = {}
    for img in images:
        img_src = img.get("src", "")
        for vid in img.get("variant_ids", []):
            if vid not in variant_img:
                variant_img[vid] = img_src
    default_img = images[0]["src"] if images else ""
    return variant_img, default_img


def process_product(product_url):
    handle = product_url.rstrip('/').split('/')[-1]

    product_data = get_product_json(handle)
    if not product_data:
        return []

    page_html = get_page_html(handle)
    specs, description = parse_page_specs(page_html)

    title = product_data.get("title", "").strip()
    product_type = product_data.get("product_type", "")
    tags = product_data.get("tags", [])

    variant_img_map, default_img = build_variant_image_map(product_data)

    downloads = specs.pop("Downloads", {})
    spec_sheet_url = ""
    install_url = ""
    lookbook_url = ""
    for name, url in downloads.items():
        nl = name.lower()
        if "spec" in nl or "tear" in nl:
            spec_sheet_url = url
        elif "install" in nl:
            install_url = url
        elif "lookbook" in nl:
            lookbook_url = url

    variants = product_data.get("variants", [])
    rows = []

    def build_row(sku="", price="", finish="", img_url=""):
        return {
            "Manufacturer": MANUFACTURER,
            "Source": product_url,
            "Image URL": img_url or default_img,
            "Product Name": title,
            "Product Family Id": title,
            "SKU": sku,
            "Description": description,
            "Price": clean_price(price),
            "Finish": finish,
            "Dimensions": specs.get("Dimensions", ""),
            "Weight": specs.get("Weight", ""),
            "Width": specs.get("Width", ""),
            "Depth": specs.get("Depth", ""),
            "Diameter": specs.get("Diameter", ""),
            "Length": specs.get("Length", ""),
            "Height": specs.get("Height", ""),
            # Standard lighting fields
            "Lamping": specs.get("Lamping", ""),
            "Lumens": specs.get("Lumens", ""),
            "Color Temp": specs.get("Color Temp", ""),
            "Input Voltage": specs.get("Input Voltage", ""),
            "Mounting Info": specs.get("Mounting Info", ""),
            "Canopy": specs.get("Canopy", ""),
            "Canopy Dimensions": specs.get("Canopy Dimensions", ""),
            "OAH Min": specs.get("OAH Min", ""),
            "OAH Max": specs.get("OAH Max", ""),
            "Cord Length": specs.get("Cord Length", ""),
            "Drop Rod": specs.get("Drop Rod", ""),
            "Shade Details": specs.get("Shade Details", ""),
            "Socket": specs.get("Socket", ""),
            "Wattage": specs.get("Wattage", ""),
            "Material": specs.get("Material", ""),
            "Backplate": specs.get("Backplate", ""),
            # Bulb-specific fields (populated by fallback parser)
            "Base": specs.get("Base", ""),
            "Associated Collection": specs.get("Associated Collection", ""),
            # Meta
            "Product Type": product_type,
            "Tags": ", ".join(tags) if tags else "",
            "Spec Sheet": spec_sheet_url,
            "Install Instructions": install_url,
            "Lookbook": lookbook_url,
        }

    if not variants:
        rows.append(build_row(img_url=default_img))
        return rows

    for v in variants:
        vid = v.get("id", 0)
        sku = v.get("sku", "")
        price_raw = v.get("price", "")
        img_url = variant_img_map.get(vid, default_img)

        opt1 = v.get("option1", "") or ""
        opt2 = v.get("option2", "") or ""
        opt3 = v.get("option3", "") or ""

        finish_parts = [p for p in [opt1, opt2, opt3] if p]
        finish = " / ".join(finish_parts)

        rows.append(build_row(sku=sku, price=price_raw, finish=finish, img_url=img_url))

    return rows


# =========================
# COLUMN ORDER
# =========================
COLUMN_ORDER = [
    "Manufacturer", "Source", "Image URL", "Product Name", "SKU", "Product Family Id",
    "Description", "Dimensions", "Weight", "Width", "Depth", "Diameter", "Length", "Height",
    "Price", "Finish", "Lamping", "Lumens", "Color Temp", "Input Voltage", "Mounting Info",
    "Canopy", "Canopy Dimensions", "OAH Min", "OAH Max", "Cord Length", "Drop Rod",
    "Shade Details", "Socket", "Wattage", "Material", "Backplate",
    # Bulb-specific columns added at end
    "Base", "Associated Collection",
    # Meta
    "Product Type", "Tags", "Spec Sheet", "Install Instructions", "Lookbook",
]


def save_to_excel(all_rows, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(COLUMN_ORDER)
    for row in all_rows:
        ws.append([row.get(col, "") for col in COLUMN_ORDER])
    wb.save(filename)


def main():
    print("=" * 70)
    print("  HENNEPIN MADE - STEP 2 FINAL (HYBRID JSON + PAGE SCRAPE)")
    print(f"  Input : {INPUT_XLSX}")
    print(f"  Output: {OUTPUT_XLSX}")
    print("=" * 70)

    df = pd.read_excel(INPUT_XLSX)

    if "Source" in df.columns:
        urls = df["Source"].dropna().unique().tolist()
    elif "Product URL" in df.columns:
        urls = df["Product URL"].dropna().unique().tolist()
    else:
        print("❌ No 'Source' column found!")
        return

    total = len(urls)
    print(f"\n  Total unique products: {total}\n")

    all_rows = []

    for idx, url in enumerate(urls, 1):
        print(f"[{idx}/{total}] {url}")
        try:
            rows = process_product(url)
            if rows:
                all_rows.extend(rows)
                r0 = rows[0]
                # Show which parser was used based on what fields got filled
                if r0.get("Base") or r0.get("Associated Collection"):
                    parser_note = f"[BULB PARSER] Base={r0.get('Base','')} Lumens={r0.get('Lumens','')} Temp={r0.get('Color Temp','')}"
                else:
                    dims = f"W={r0.get('Width','')} H={r0.get('Height','')} D={r0.get('Depth','')} Dia={r0.get('Diameter','')}"
                    parser_note = f"[STD PARSER] {dims}"
                print(f"  ✓ {r0['Product Name']} → {len(rows)} variants | {parser_note}")
            else:
                print(f"  ⚠ No data")
        except Exception as e:
            print(f"  ❌ Error: {e}")

        if idx % 5 == 0:
            save_to_excel(all_rows, OUTPUT_XLSX)
            print(f"  💾 Auto-saved ({len(all_rows)} rows)")

        time.sleep(1)

    save_to_excel(all_rows, OUTPUT_XLSX)
    print(f"\n{'='*70}")
    print(f"  ✅ DONE! {len(all_rows)} rows → {OUTPUT_XLSX}")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()