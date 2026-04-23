# -*- coding: utf-8 -*-
"""
scraper.py — Gabby
All categories use Shopify API on gabriellawhite.com (same backend as gabby.com).
SKU + Price + Description from API; Dimensions/Finish/Color from product page.
Demo:   d:/phase-05 (AU)/Agent/Gabby/Demo/Gabby_demo.xlsx
Full:   d:/phase-05 (AU)/Agent/Gabby/Data/Gabby.xlsx
"""

import re
import time
import requests
import pandas as pd
from pathlib import Path
from bs4 import BeautifulSoup

# ── CONFIG ─────────────────────────────────────────────────────────────────────
DEMO_MODE    = True
DEMO_PER_CAT = 3
SHOPIFY_BASE = "https://gabriellawhite.com"
GABBY_BASE   = "https://gabby.com"
MANUFACTURER = "Gabby"
DEMO_FILE    = Path("d:/phase-05 (AU)/Agent/Gabby/Demo/Gabby_demo.xlsx")
OUTPUT_FILE  = Path("d:/phase-05 (AU)/Agent/Gabby/Data/Gabby.xlsx")

# Client-provided display links (overrides auto-generated gabby.com/collections/... link)
CAT_LINKS_OVERRIDE = {
    "Cabinets":        "https://gabby.com/products/indoor-dining/cabinets, https://gabby.com/products/storage/sideboards",
    "Dressers & Chests": "https://gabby.com/products/bedroom/dressers, https://gabby.com/products/bedroom/chests",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

COLUMNS = [
    "Index", "Category", "Manufacturer", "Source", "Image URL",
    "Product Name", "SKU", "Product Family Id", "Description",
    "Weight", "Width", "Depth", "Diameter", "Length", "Height",
    "Price", "Special Order", "Location",
    "Assembly Required",
    "More Information", "Warranty", "Care Instructions",
    "Specification Sheet", "Pillow Size", "Pillow Fill Material",
    "Tags", "Notes",
]
_FIXED_COLS = set(COLUMNS)

# ── COLLECTIONS — (category_name, [handles], dedup_key, product_type_filter) ──
# product_type_filter: substring match against Shopify product_type (case-insensitive)
#                      None = no filter (take all products from collection)
# dedup_key: "global" = skip globally seen SKUs | "local" = allow across categories
COLLECTIONS = [
    ("Nightstands",              ["indoor-nightstands"],                         "global", None),
    ("Coffee & Cocktail Tables", ["indoor-coffee-tables"],                       "global", None),
    ("Side & End Tables",        ["indoor-side-end-tables"],                     "global", None),
    ("Dining Tables",            ["indoor-dining-tables"],                       "global", None),
    ("Consoles",                 ["indoor-console-tables"],                      "global", None),
    ("Beds & Headboards",        ["beds-and-headboards"],                        "global", None),
    ("Desks",                    ["desks"],                                      "global", None),
    ("Cabinets",                 ["indoor-cabinets"],                            "global", None),
    ("Bookcases",                ["indoor-bookcases"],                           "global", None),
    ("Accent Tables",            ["dining-room"],                                "global", None),
    ("Bar Carts",                ["indoor-serving-bar-carts"],                   "global", None),
    ("Dressers & Chests",        ["indoor-dressers", "indoor-storage"],          "global", None),
    ("Dining Chairs",            ["indoor-dining-chairs"],                       "global", None),
    ("Bar Stools",               ["indoor-bar-counter-height-stools"],           "global", None),
    ("Benches",                  ["indoor-benches-banquettes"],                  "global", None),
    ("Ottomans",                 ["indoor-ottomans-stools"],                     "global", None),
    ("Chandeliers",              ["hanging-lighting"],                           "local",  "chandelier"),
    ("Pendants",                 ["ceiling-lights"],                             "local",  "pendant"),
    ("Sconces",                  ["wall-lights"],                                "local",  "sconce"),
    ("Flush Mount",              ["ceiling-lights"],                             "local",  "flush"),
    ("Table Lamps",              ["lamps"],                                      "local",  "table lamp"),
    ("Floor Lamps",              ["lamps"],                                      "local",  "floor lamp"),
    ("Mirrors",                  ["mirrors"],                                    "global", None),
    ("Pillows & Throws",         ["decorative-accessories"],                     "local",  "pillow"),
    ("Rugs",                     ["decorative-accessories"],                     "local",  "rug"),
]

# ── DIMENSION KEY MAP — only 6 fixed columns; all other dim keys go dynamic ────
DIM_MAP = {
    "product weight":   "Weight",
    "product width":    "Width",
    "product depth":    "Depth",
    "product diameter": "Diameter",
    "product length":   "Length",
    "product height":   "Height",
}

# ── HELPERS ────────────────────────────────────────────────────────────────────
def clean_dim(value: str) -> str:
    value = re.sub(r"[^\d.,]+", "", value)
    return value.split(",")[0].strip()


def price_to_int(raw: str) -> str:
    """'1549.00' or '$1,549' -> '1549'"""
    raw = str(raw).replace("$", "").replace(",", "").strip()
    try:
        return str(int(float(raw)))
    except Exception:
        return raw


def strip_html(html: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html or "")).strip()


# ── SHOPIFY API — get all products from a collection ──────────────────────────
def api_collection(handle: str, limit: int = 0) -> list[dict]:
    products = []
    page = 1
    while True:
        url = f"{SHOPIFY_BASE}/collections/{handle}/products.json?limit=250&page={page}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            if r.status_code != 200:
                break
            batch = r.json().get("products", [])
            if not batch:
                break
            for p in batch:
                phandle = p.get("handle", "")
                name    = p.get("title", "")
                images  = p.get("images") or []
                img     = images[0].get("src", "") if images else ""
                variants = p.get("variants") or [{}]
                v0       = variants[0]
                sku      = v0.get("sku", "")
                price    = price_to_int(v0.get("price", ""))
                body     = strip_html(p.get("body_html", ""))
                tags     = ", ".join(p.get("tags") or [])
                products.append({
                    "handle":        phandle,
                    "Product Name":  name,
                    "SKU":           sku,
                    "Image URL":     img,
                    "Description":   body,
                    "Price":         price,
                    "Tags":          tags,
                    "product_type":  p.get("product_type", ""),
                })
                if limit and len(products) >= limit:
                    break
            if limit and len(products) >= limit:
                break
            if len(batch) < 250:
                break
            page += 1
            time.sleep(0.4)
        except Exception as e:
            print(f"    API error ({handle}): {e}")
            break
    return products


# ── PRODUCT PAGE — dimensions + specs ─────────────────────────────────────────
def scrape_specs(url: str) -> dict:
    data  = {}
    extra = {}   # dynamic columns: key = column name, value = cell value
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            data["_extra"] = extra
            return data
        soup = BeautifulSoup(r.text, "html.parser")

        for sec in soup.find_all("div", class_="specs-attributes-section"):
            h3 = sec.find("h3")
            if not h3:
                continue
            title = h3.get_text(strip=True).upper()
            spans = sec.find_all("span")
            pairs = [
                (spans[i].get_text(strip=True), spans[i+1].get_text(strip=True) if i+1 < len(spans) else "")
                for i in range(0, len(spans), 2)
            ]

            if "DIMENSIONS" in title:
                for k, v in pairs:
                    key_lower = k.lower().strip()
                    col = DIM_MAP.get(key_lower)
                    if col:
                        cleaned = clean_dim(v)
                        if cleaned:
                            data[col] = cleaned
                    elif k.strip() and v.strip():
                        col_name = k.strip().title()
                        cleaned  = clean_dim(v)
                        if cleaned:
                            # route to fixed col if name matches, else dynamic
                            if col_name in _FIXED_COLS:
                                data.setdefault(col_name, cleaned)
                            else:
                                extra[col_name] = cleaned

            elif "FEATURES" in title:
                for k, v in pairs:
                    if k.strip() and v.strip():
                        col_name = k.strip().title()
                        if col_name in _FIXED_COLS:
                            data.setdefault(col_name, v.strip())
                        else:
                            extra[col_name] = v.strip()

            elif "ASSEMBLY" in title:
                for k, v in pairs:
                    if "assembly required" in k.lower():
                        data["Assembly Required"] = v
                        break

        # Accordion sections: Warranty, Care Instructions, More Information
        for accordion in soup.find_all("div", class_="accordion-item"):
            h2 = accordion.find("h2")
            if not h2:
                continue
            heading = h2.get_text(strip=True).lower()
            content_div = accordion.find("div", class_=lambda c: c and "pb-5" in c)
            if not content_div:
                continue

            if "warranty" in heading:
                data["Warranty"] = content_div.get_text(separator=" ", strip=True)
            elif "care" in heading:
                ul = accordion.find("ul")
                if ul:
                    data["Care Instructions"] = "\n".join(li.get_text(strip=True) for li in ul.find_all("li"))
                else:
                    data["Care Instructions"] = content_div.get_text(separator=" ", strip=True)
            elif "more information" in heading:
                ul = accordion.find("ul")
                if ul:
                    data["More Information"] = "\n".join(li.get_text(strip=True) for li in ul.find_all("li"))
                else:
                    data["More Information"] = content_div.get_text(separator="\n", strip=True)

        # Specification Sheet URL
        spec_div = soup.find("div", attrs={"sub-section-id": lambda v: v and "spec_sheet_link" in v})
        if spec_div:
            a = spec_div.find("a", href=True)
            if a:
                href = a["href"]
                data["Specification Sheet"] = (SHOPIFY_BASE + href) if href.startswith("/") else href

        # Pillow Size (radio inputs for pillow size)
        sizes = []
        for inp in soup.find_all("input", attrs={"name": "properties[pillow_size]"}):
            label_id = inp.get("id")
            if label_id:
                lbl = soup.find("label", attrs={"for": label_id})
                if lbl:
                    p = lbl.find("p")
                    if p:
                        t = p.get_text(strip=True)
                        if t and t not in sizes:
                            sizes.append(t)
        if sizes:
            data["Pillow Size"] = "\n".join(sizes)

        # Pillow Fill Material (radio inputs for fill_material)
        fills = []
        for inp in soup.find_all("input", attrs={"name": "properties[fill_material]"}):
            label_id = inp.get("id")
            if label_id:
                lbl = soup.find("label", attrs={"for": label_id})
                if lbl:
                    p = lbl.find("p")
                    if p:
                        t = p.get_text(strip=True)
                        if t and t not in fills:
                            fills.append(t)
        if fills:
            data["Pillow Fill Material"] = ", ".join(fills)

    except Exception as e:
        print(f"    spec error: {e}")
    data["_extra"] = extra
    return data


# ── SAVE EXCEL ─────────────────────────────────────────────────────────────────
def save_excel(all_rows: list[dict], save_path: Path,
               extra_keys: list[str] = None,
               cat_links: dict = None) -> None:
    save_path.parent.mkdir(parents=True, exist_ok=True)
    cat_order  = list(dict.fromkeys(r["Category"] for r in all_rows))
    extra_keys = extra_keys or []
    cat_links  = cat_links  or {}

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        for cat_name in cat_order:
            cat_rows = [r for r in all_rows if r["Category"] == cat_name]

            # Per-sheet: only include dynamic cols that have at least one non-empty value
            sheet_extra = [k for k in extra_keys if any(r.get(k, "") for r in cat_rows)]
            final_cols  = COLUMNS + sheet_extra
            n           = len(final_cols)

            cat_link = cat_links.get(cat_name, "")
            meta = [
                ["Brand:", MANUFACTURER] + [""] * (n - 2),
                ["Link:",  cat_link]     + [""] * (n - 2),
                [""] * n,
                final_cols,
            ]

            cat_df = pd.DataFrame(cat_rows, columns=final_cols).fillna("")
            cat_df["Index"] = range(1, len(cat_df) + 1)
            rows = meta + cat_df.values.tolist()
            pd.DataFrame(rows).to_excel(
                writer, sheet_name=cat_name[:31], index=False, header=False
            )

    print(f"Saved {len(all_rows)} products across {len(cat_order)} sheets -> {save_path}")


# ── MAIN ───────────────────────────────────────────────────────────────────────
def main():
    print(f"Gabby scraper | DEMO_MODE={DEMO_MODE}")
    all_rows         = []
    global_seen      = set()
    extra_keys       = []
    extra_keys_seen  = set()
    limit            = DEMO_PER_CAT if DEMO_MODE else 0
    cat_links        = {
                            cat: CAT_LINKS_OVERRIDE.get(cat, f"{GABBY_BASE}/collections/{handles[0]}")
                            for cat, handles, *_ in COLLECTIONS
                        }

    for i, (category, handles, dedup_key, type_filter) in enumerate(COLLECTIONS):
        print(f"\n[{i+1}/{len(COLLECTIONS)}] {category}  handles={handles}")

        # Collect all raw products from all handles
        raw_products = []
        for handle in handles:
            batch = api_collection(handle, limit=0)  # always fetch all, slice later
            print(f"  API [{handle}]: {len(batch)} products")
            raw_products.extend(batch)

        if not raw_products:
            print("  No products, skipping.")
            continue

        # Apply product_type filter for categories sharing a collection handle
        if type_filter:
            raw_products = [p for p in raw_products
                            if type_filter in p.get("product_type", "").lower()]
            print(f"  After type filter '{type_filter}': {len(raw_products)} products")

        # Deduplicate within the raw batch (by SKU)
        seen_in_cat = set()
        unique = []
        for p in raw_products:
            key = p["SKU"] or p["handle"]
            if key and key not in seen_in_cat:
                seen_in_cat.add(key)
                unique.append(p)

        # Apply demo limit
        if limit:
            unique = unique[:limit]

        added = 0
        for p in unique:
            sku = p["SKU"] or p["handle"]

            # Global dedup (furniture categories — skip if already in another sheet)
            if dedup_key == "global" and sku in global_seen:
                continue
            if dedup_key == "global":
                global_seen.add(sku)

            phandle      = p["handle"]
            product_url  = f"{GABBY_BASE}/collections/{handles[0]}/products/{phandle}"
            name         = p["Product Name"]
            family_id    = re.sub(r"\s*-\s*.*", "", name).strip()

            # Scrape specs from product page
            spec_url = f"{SHOPIFY_BASE}/products/{phandle}"
            specs    = scrape_specs(spec_url)
            extra    = specs.pop("_extra", {})

            # Register new dynamic keys (preserve insertion order)
            for k in extra:
                if k not in extra_keys_seen:
                    extra_keys_seen.add(k)
                    extra_keys.append(k)

            row = {c: "" for c in COLUMNS}
            row.update({
                "Category":          category,
                "Manufacturer":      MANUFACTURER,
                "Source":            product_url,
                "Image URL":         p["Image URL"],
                "Product Name":      name,
                "SKU":               p["SKU"],
                "Product Family Id": family_id,
                "Description":       p["Description"],
                "Price":             p["Price"],
                "Tags":              p["Tags"],
            })
            row.update(specs)
            row.update(extra)
            all_rows.append(row)
            added += 1

            print(
                f"  [{len(all_rows)}] {name[:45]}  SKU:{row['SKU']}  "
                f"W:{row.get('Width','')} D:{row.get('Depth','')} H:{row.get('Height','')}  "
                f"Price:{row['Price']}"
            )
            time.sleep(0.8)

        print(f"  -> {added} added")

        if not DEMO_MODE and len(all_rows) % 100 == 0 and all_rows:
            save_excel(all_rows, OUTPUT_FILE, extra_keys=extra_keys, cat_links=cat_links)
            print("  Autosaved.")

    if not all_rows:
        print("WARNING: 0 products found.")
        return

    out = DEMO_FILE if DEMO_MODE else OUTPUT_FILE
    save_excel(all_rows, out, extra_keys=extra_keys, cat_links=cat_links)
    print(f"\n{'Demo' if DEMO_MODE else 'Full run'} done -> {out}")


if __name__ == "__main__":
    main()
