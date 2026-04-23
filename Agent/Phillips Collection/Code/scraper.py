# -*- coding: utf-8 -*-
"""
scraper.py - Phillips Collection
Listing page  → SKU, Name, Price, Image, W/D/H dims (server-rendered HTML, ?p=N pagination)
Product page  → Description, Weight, Material, Finish
"""

import re
import time
import requests
import pandas as pd
from pathlib import Path
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# ── CONFIG ────────────────────────────────────────────────────────────────────
DEMO_MODE    = False
DEMO_PER_CAT = 3
BASE_URL     = "https://www.phillipscollection.com"
MANUFACTURER = "Phillips Collection"

DEMO_FILE   = Path("d:/phase-05 (AU)/Agent/Phillips Collection/Demo/Phillips_Collection_demo.xlsx")
OUTPUT_FILE = Path("d:/phase-05 (AU)/Agent/Phillips Collection/Data/Phillips_Collection.xlsx")

COLUMNS = [
    "Index", "Category", "Manufacturer", "Source", "Image URL",
    "Product Name", "SKU", "Product Family Id", "Description",
    "Weight", "Width", "Depth", "Diameter", "Height",
    "Seat Width", "Seat Depth", "Seat Height", "Arm Height",
    "Price", "Finish", "Special Order", "Location",
    "Materials", "Tags", "Notes",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

CATEGORIES = [
    ("Coffee Tables",      "https://www.phillipscollection.com/furniture/coffee-tables/"),
    ("Dining Tables",      "https://www.phillipscollection.com/furniture/dining-tables/"),
    ("Consoles",           "https://www.phillipscollection.com/furniture/consoles/"),
    ("Side Tables",        "https://www.phillipscollection.com/furniture/side-tables/"),
    ("Desks",              "https://www.phillipscollection.com/furniture/desks/"),
    ("Seating",            "https://www.phillipscollection.com/furniture/seating/"),
    ("Benches",            "https://www.phillipscollection.com/furniture/benches/"),
    ("Stools",             "https://www.phillipscollection.com/furniture/stools/"),
    ("Bar",                "https://www.phillipscollection.com/furniture/bar/"),
    ("Botanical",          "https://www.phillipscollection.com/wall-decor/botanical/"),
    ("Figurative",         "https://www.phillipscollection.com/wall-decor/figurative/"),
    ("Mirrors",            "https://www.phillipscollection.com/wall-decor/mirrors/"),
    ("Modular",            "https://www.phillipscollection.com/wall-decor/modular/"),
    ("Panels",             "https://www.phillipscollection.com/wall-decor/panels/"),
    ("Shelves",            "https://www.phillipscollection.com/wall-decor/shelves/"),
    ("Statement Pieces",   "https://www.phillipscollection.com/wall-decor/statement-pieces/"),
    ("Abstract",           "https://www.phillipscollection.com/sculpture/abstract/"),
    ("Figures",            "https://www.phillipscollection.com/sculpture/figures/"),
    ("Animals",            "https://www.phillipscollection.com/sculpture/animals/"),
    ("Tabletop Sculpture", "https://www.phillipscollection.com/sculpture/tabletop/"),
    ("Fountains",          "https://www.phillipscollection.com/sculpture/fountains/"),
    ("Vases",              "https://www.phillipscollection.com/accessories/vases/"),
    ("Bowls",              "https://www.phillipscollection.com/accessories/bowls/"),
    ("Planters",           "https://www.phillipscollection.com/accessories/planters/"),
    ("Screens",            "https://www.phillipscollection.com/accessories/screens/"),
    ("Pedestals",          "https://www.phillipscollection.com/accessories/pedestals/"),
    ("Tabletop",           "https://www.phillipscollection.com/accessories/tabletop/"),
    ("Rugs",               "https://www.phillipscollection.com/accessories/rugs/"),
    ("Floor Lamps",        "https://www.phillipscollection.com/lighting/floor-lamps/"),
    ("Table Lamps",        "https://www.phillipscollection.com/lighting/table-lamps/"),
    ("Hanging Lamps",      "https://www.phillipscollection.com/lighting/hanging-lamps/"),
    ("Wall Sconces",       "https://www.phillipscollection.com/lighting/sconces/"),
    ("Outdoor",            "https://www.phillipscollection.com/outdoor/"),
]

# ── PARSE DIMENSIONS "39x39x15\"h" → (W, D, H) ───────────────────────────────
def parse_dims(text):
    text = text.replace('"', '').replace("'", '').lower().strip()
    # Remove packed size in parens
    text = re.sub(r'\(.*?\)', '', text).strip()
    # Remove trailing letters like 'h', 'w', 'd'
    nums = re.findall(r'\d+\.?\d*', text)
    w = nums[0] if len(nums) > 0 else ""
    d = nums[1] if len(nums) > 1 else ""
    h = nums[2] if len(nums) > 2 else ""
    return w, d, h

# ── PARSE WEIGHT "312 lbs (390 lbs packed)" → "312" ──────────────────────────
def parse_weight(text):
    m = re.search(r'(\d+\.?\d*)\s*lbs?', text, re.IGNORECASE)
    return m.group(1) if m else ""

# ── PARSE PRICE "$2,609.00" → "2609.00" ──────────────────────────────────────
def parse_price(text):
    text = re.sub(r'[^0-9.]', '', text.replace(',', ''))
    return text if text else ""

# ── SCRAPE ONE LISTING PAGE → list of product dicts ───────────────────────────
def scrape_listing_page(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.text, 'html.parser')
    except Exception as e:
        print(f"  [ERR] listing {url}: {e}")
        return [], None

    products = []
    for item in soup.select('.item'):
        a = item.find('a', href=True)
        if not a:
            continue
        prod_url = a['href']
        if not prod_url.startswith('http'):
            prod_url = urljoin(BASE_URL, prod_url)

        name = item.find('h4')
        name = name.get_text(strip=True) if name else ""

        # "ID65143 / 39x39x15\"h"
        info_p = item.select_one('p.museo_sans300')
        sku, dims_raw = "", ""
        if info_p:
            info_txt = info_p.get_text(strip=True)
            parts = info_txt.split(' / ', 1)
            sku = parts[0].strip()
            dims_raw = parts[1].strip() if len(parts) > 1 else ""

        price_span = item.select_one('span.museo_sans700')
        price = ""
        if price_span:
            price = parse_price(price_span.get_text())

        # First non-hover image
        img = item.find('img', class_=lambda c: c != 'hover-image' if c else True)
        image_url = ""
        if img and img.get('src'):
            src = img['src']
            image_url = src if src.startswith('http') else urljoin(BASE_URL + '/', src)

        w, d, h = parse_dims(dims_raw) if dims_raw else ("", "", "")

        if sku or name:
            products.append({
                "url": prod_url,
                "name": name,
                "sku": sku,
                "price": price,
                "image": image_url,
                "width": w,
                "depth": d,
                "height": h,
            })

    # Next page link
    next_url = None
    pag = soup.select_one('.pagination')
    if pag:
        for a in pag.find_all('a', href=True):
            href = a['href']
            if not href.startswith('http'):
                href = urljoin(BASE_URL + '/', href)
            # Detect next page link (not current active)
            parent_cls = ' '.join(a.get('class', []))
            if 'active' not in parent_cls:
                next_url = href  # will pick last — but we paginate manually below

    return products, soup

# ── GET ALL PRODUCTS FOR A CATEGORY (all pages) ───────────────────────────────
def get_all_products(cat_url, demo=False):
    all_products = []
    page = 1
    seen_skus = set()

    while True:
        url = cat_url if page == 1 else f"{cat_url}?p={page}"
        products, soup = scrape_listing_page(url)

        if not products:
            break

        new_added = 0
        for p in products:
            if p['sku'] in seen_skus:
                continue
            seen_skus.add(p['sku'])
            all_products.append(p)
            new_added += 1
            if demo and len(all_products) >= DEMO_PER_CAT:
                return all_products

        print(f"    page {page}: {new_added} products (total: {len(all_products)})")

        # Check if next page exists
        if soup:
            pag = soup.select_one('.pagination')
            if pag:
                active_links = pag.select('a')
                # Find if there's a page beyond current
                has_next = False
                for a in active_links:
                    href = a.get('href', '')
                    if f'p={page+1}' in href or (f'&p={page+1}' in href):
                        has_next = True
                        break
                if not has_next:
                    # Try checking if any link has higher page number
                    all_pages = []
                    for a in active_links:
                        m = re.search(r'[?&]p=(\d+)', a.get('href', ''))
                        if m:
                            all_pages.append(int(m.group(1)))
                    if all_pages and page >= max(all_pages):
                        break
                    elif not all_pages:
                        break
            else:
                break  # No pagination → single page

        page += 1
        time.sleep(0.5)

    return all_products

# ── SCRAPE PRODUCT DETAIL PAGE → description, weight, material, finish ────────
def scrape_product_page(url):
    result = {"description": "", "weight": "", "material": "", "finish": ""}
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.text, 'html.parser')
    except Exception as e:
        print(f"  [ERR] product {url}: {e}")
        return result

    panels = soup.select('.panel_content')
    for panel in panels:
        rows = panel.select('.two-columns-container')
        if rows:
            # This is the specs panel
            for row in rows:
                divs = row.find_all('div', recursive=False)
                if len(divs) < 2:
                    continue
                label = divs[0].get_text(strip=True).lower().rstrip(':')
                value_div = divs[1]
                # Remove packed size span (gray)
                for gray in value_div.select('.gray'):
                    gray.decompose()
                value = value_div.get_text(strip=True)

                if 'weight' in label:
                    result['weight'] = parse_weight(value)
                elif 'material' in label:
                    result['material'] = value
                elif 'finish' in label:
                    result['finish'] = value
        else:
            # Text-only panel → description (first one) or care instructions
            text = panel.get_text(strip=True)
            if text and not result['description'] and len(text) > 30:
                result['description'] = text

    return result

# ── BUILD ONE ROW ─────────────────────────────────────────────────────────────
def build_row(idx, category, listing, detail):
    return {
        "Index":            idx,
        "Category":         category,
        "Manufacturer":     MANUFACTURER,
        "Source":           listing['url'],
        "Image URL":        listing['image'],
        "Product Name":     listing['name'],
        "SKU":              listing['sku'],
        "Product Family Id":"",
        "Description":      detail['description'],
        "Weight":           listing.get('weight') or detail['weight'],
        "Width":            listing['width'],
        "Depth":            listing['depth'],
        "Diameter":         "",
        "Height":           listing['height'],
        "Seat Width":       "",
        "Seat Depth":       "",
        "Seat Height":      "",
        "Arm Height":       "",
        "Price":            listing['price'],
        "Finish":           detail['finish'],
        "Special Order":    "",
        "Location":         "",
        "Materials":        detail['material'],
        "Tags":             "",
        "Notes":            "",
    }

# ── SAVE EXCEL (Alfonso Marina format) ───────────────────────────────────────
def save_excel(all_rows, save_path):
    save_path.parent.mkdir(parents=True, exist_ok=True)
    df_all = pd.DataFrame(all_rows, columns=COLUMNS)

    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        for cat_name, cat_df in df_all.groupby("Category", sort=False):
            cat_df = cat_df.reset_index(drop=True)
            # Re-index within sheet
            cat_df["Index"] = range(1, len(cat_df) + 1)

            meta = [
                [MANUFACTURER] + [""] * (len(COLUMNS) - 1),
                [BASE_URL]     + [""] * (len(COLUMNS) - 1),
                [""]           * len(COLUMNS),
                COLUMNS,
            ]
            sheet_name = str(cat_name)[:31]
            pd.DataFrame(meta + cat_df.values.tolist()).to_excel(
                writer, sheet_name=sheet_name, index=False, header=False
            )

    print(f"\nSaved {len(all_rows)} products across {df_all['Category'].nunique()} sheets -> {save_path}")

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    mode = "DEMO" if DEMO_MODE else "FULL RUN"
    save_path = DEMO_FILE if DEMO_MODE else OUTPUT_FILE
    print(f"Phillips Collection — {mode}")
    print(f"Output: {save_path}\n")

    all_rows = []
    seen_skus = set()
    global_idx = 1

    for cat_name, cat_url in CATEGORIES:
        print(f"[{cat_name}] {cat_url}")
        products = get_all_products(cat_url, demo=DEMO_MODE)

        cat_count = 0
        for p in products:
            sku = p['sku']
            if not DEMO_MODE and sku in seen_skus:
                continue
            seen_skus.add(sku)

            detail = scrape_product_page(p['url'])
            row = build_row(global_idx, cat_name, p, detail)
            all_rows.append(row)
            global_idx += 1
            cat_count += 1

            print(f"  [{global_idx-1}] {p['name']}  W:{p['width']} D:{p['depth']} H:{p['height']}  "
                  f"Wt:{detail['weight']}  Finish:{detail['finish']}  Mat:{detail['material'][:15]}")
            time.sleep(0.3)

        print(f"  -> {cat_count} products added\n")

    save_excel(all_rows, save_path)

if __name__ == "__main__":
    main()
