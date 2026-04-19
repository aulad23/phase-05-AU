# -*- coding: utf-8 -*-
"""
Zuo Modern — Step 2: Product Detail Scraper
============================================
Input:  zuo_products_step1.xlsx (from Step 1)
Output: zuo_products_final.xlsx

Extracts from each product page:
- Image URL (high-res from detail page)
- SKU
- Product Family Id
- Description
- Product Details (dynamic key→column)
- Product Dimensions → Width, Depth, Height, Diameter, Length, Weight
- Seat Dimensions → Seat Width, Seat Depth, Seat Height, Arm Height, Arm Width
- Packaging (dynamic key→column)

Usage: python zuo_step2_details.py
"""

import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─── CONFIG ───
INPUT_FILE = "zuo_Dining_Chair.xlsx"
OUTPUT_FILE = "zuo_Dining_Chair_Final.xlsx"

# Fixed columns order (always first, in this exact serial)
FIXED_COLUMNS = [
    "Manufacturer", "Source", "Image URL", "Product Name",
    "SKU", "Product Family Id", "Description", "Weight",
    "Width", "Depth", "Diameter", "Length", "Height",
]

# Dimension abbreviation → column name
DIM_MAP = {
    "W": "Width",
    "D": "Depth",
    "H": "Height",
    "L": "Length",
}

# Diameter variants
DIA_KEYS = ["Dia", "DIA", "Dia ", "Diam", "DIAM", "Diameter"]

# Seat abbreviation → column name
SEAT_MAP = {
    "SW": "Seat Width",
    "SD": "Seat Depth",
    "SH": "Seat Height",
    "AH": "Arm Height",
    "AW": "Arm Width",
}


# ─── SELENIUM SETUP ───
def build_driver():
    opts = Options()
    # opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    )
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(60)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver


# ─── DIMENSION PARSING ───
def parse_product_dimensions(dim_str):
    """
    Parse: '21.3" W x 22.2" D x 35" H' → {Width: 21.3, Depth: 22.2, Height: 35}
    Also handles: '18" Dia x 24" H', '18" DIA x 24" H'
    """
    result = {v: "" for v in DIM_MAP.values()}
    result["Diameter"] = ""

    if not dim_str or str(dim_str).strip() in ("", "nan"):
        return result

    dim_str = str(dim_str).strip()
    dim_str = dim_str.replace('"', '"').replace('"', '"').replace('″', '"').replace("''", '"')

    # Check for Diameter keywords
    for dia_key in DIA_KEYS:
        pattern = rf'([\d.]+)\s*["\']?\s*{re.escape(dia_key)}'
        m = re.search(pattern, dim_str, re.IGNORECASE)
        if m:
            result["Diameter"] = m.group(1)
            break

    # Standard W, D, H, L
    for abbr, col in DIM_MAP.items():
        # Pattern: number + optional " + space + abbreviation
        pattern = rf'([\d.]+)\s*["\']?\s*{abbr}(?:\b|(?=[^a-zA-Z]))'
        m = re.search(pattern, dim_str)
        if m:
            result[col] = m.group(1)

    return result


def parse_seat_dimensions(seat_str):
    """
    Parse: '19.3" SW x 18.9" SD x 17.5" SH' → {Seat Width: 19.3, ...}
    """
    result = {v: "" for v in SEAT_MAP.values()}

    if not seat_str or str(seat_str).strip() in ("", "nan"):
        return result

    seat_str = str(seat_str).strip()
    seat_str = seat_str.replace('"', '"').replace('"', '"').replace('″', '"').replace("''", '"')

    for abbr, col in SEAT_MAP.items():
        pattern = rf'([\d.]+)\s*["\']?\s*{abbr}(?:\b|(?=[^a-zA-Z]))'
        m = re.search(pattern, seat_str)
        if m:
            result[col] = m.group(1)

    return result


def parse_weight(text):
    """Extract weight from text like 'Product 1 Weight (lbs.): 24.5' or '24.5 lbs'"""
    if not text:
        return ""
    m = re.search(r'([\d.]+)\s*(?:lbs?\.?|pounds?)', str(text), re.IGNORECASE)
    if m:
        return m.group(1)
    # Just a number
    m = re.search(r'([\d.]+)', str(text))
    if m:
        return m.group(1)
    return ""


def product_family_from_name(name):
    """Extract family ID from product name (everything before last color/material word)."""
    if not name:
        return ""
    # Remove common color/finish suffixes
    colors = [
        "Black", "White", "Gray", "Grey", "Brown", "Beige", "Ivory",
        "Walnut", "Natural", "Gold", "Silver", "Bronze", "Chrome",
        "Blue", "Green", "Red", "Pink", "Orange", "Yellow", "Purple",
        "Teal", "Cream", "Taupe", "Rust", "Charcoal", "Navy",
        "Brass", "Copper", "Nickel", "Oak", "Ash", "Mahogany",
        "Espresso", "Cognac", "Caramel", "Tan", "Sand", "Sage",
        "Olive", "Moss", "Slate", "Smoke", "Pewter", "Antique",
        "Multicolor", "Clear", "Frosted", "Matte", "Glossy",
    ]
    parts = name.strip().split()
    # Walk backwards removing color words
    while parts and parts[-1] in colors:
        parts.pop()
    # Also remove "&" if trailing
    while parts and parts[-1] in ("&", "and", "-"):
        parts.pop()
    return " ".join(parts).strip()


def extract_product_details(driver):
    """
    Extract all product data from a Zuo product detail page.
    Returns dict with all fields.
    """
    data = {
        "Image URL": "",
        "SKU": "",
        "Product Family Id": "",
        "Description": "",
    }
    # Initialize dimension columns
    for col in ["Width", "Depth", "Height", "Diameter", "Length", "Weight",
                "Seat Width", "Seat Depth", "Seat Height", "Arm Height", "Arm Width"]:
        data[col] = ""

    dynamic_details = {}    # Product Details dynamic key-value
    dynamic_packaging = {}  # Packaging dynamic key-value

    page_text = driver.page_source

    # ─── IMAGE URL (high-res) ───
    try:
        img_els = driver.find_elements(By.CSS_SELECTOR,
            "img.gallery-placeholder__image, img.fotorama__img, "
            "img[data-role='image'], .product.media img, "
            ".MagicZoom img, img.product-image-photo"
        )
        for img in img_els:
            src = img.get_attribute("src") or img.get_attribute("data-src") or ""
            if src and "placeholder" not in src.lower() and src.startswith("http"):
                data["Image URL"] = src
                break
    except:
        pass

    # Fallback: get from og:image
    if not data["Image URL"]:
        try:
            og = driver.find_element(By.CSS_SELECTOR, "meta[property='og:image']")
            data["Image URL"] = og.get_attribute("content") or ""
        except:
            pass

    # ─── SKU ───
    try:
        # Try multiple selectors for SKU
        sku_selectors = [
            "div[itemprop='sku']",
            "span[itemprop='sku']",
            ".product-info-stock-sku .value",
            ".product.attribute.sku .value",
            "[data-th='SKU']",
        ]
        for sel in sku_selectors:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                sku = el.text.strip()
                if sku:
                    data["SKU"] = sku
                    break
            except:
                continue
    except:
        pass

    # Fallback: from page source
    if not data["SKU"]:
        m = re.search(r'"sku"\s*:\s*"([^"]+)"', page_text)
        if m:
            data["SKU"] = m.group(1)

    # ─── DESCRIPTION ───
    try:
        desc_selectors = [
            "#product-description",
            "div[itemprop='description']",
            ".product.attribute.description .value",
            ".product-info-description .value",
            ".product.info.description .value",
            "#description .value",
        ]
        for sel in desc_selectors:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                desc = el.text.strip()
                if desc:
                    data["Description"] = desc
                    break
            except:
                continue
    except:
        pass

    # ─── PRODUCT DETAILS / SPECIFICATIONS (Dynamic) ───
    # Zuo uses table or div-based key-value pairs for specs
    try:
        # Method 1: Table rows
        spec_tables = driver.find_elements(By.CSS_SELECTOR,
            "table.data.table.additional-attributes, "
            "#product-attribute-specs-table, "
            "table.product-attributes, "
            ".additional-attributes-wrapper table"
        )
        for table in spec_tables:
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                try:
                    th = row.find_element(By.TAG_NAME, "th")
                    td = row.find_element(By.TAG_NAME, "td")
                    key = th.text.strip()
                    value = td.text.strip()
                    if key and value:
                        dynamic_details[key] = value
                except:
                    continue
    except:
        pass

    # Method 2: div-based specs (common in Magento)
    try:
        spec_divs = driver.find_elements(By.CSS_SELECTOR,
            ".product-info-main .product.attribute, "
            ".product-details .attribute, "
            ".product-info .spec-row, "
            ".product-specs .row, "
            ".col-data, .data-table"
        )
        for div in spec_divs:
            try:
                label_el = div.find_element(By.CSS_SELECTOR, ".label, .title, th, dt, .spec-label")
                value_el = div.find_element(By.CSS_SELECTOR, ".value, .data, td, dd, .spec-value")
                key = label_el.text.strip().rstrip(":")
                value = value_el.text.strip()
                if key and value and key not in dynamic_details:
                    dynamic_details[key] = value
            except:
                continue
    except:
        pass

    # Method 3: JavaScript extraction - get all specification data
    try:
        js_specs = driver.execute_script("""
            var specs = {};
            // Try data tables
            document.querySelectorAll('table tr').forEach(function(row) {
                var th = row.querySelector('th, td:first-child');
                var td = row.querySelector('td:last-child');
                if (th && td && th !== td) {
                    var key = th.textContent.trim();
                    var val = td.textContent.trim();
                    if (key && val && key !== val) specs[key] = val;
                }
            });
            // Try dl/dt/dd
            document.querySelectorAll('dl').forEach(function(dl) {
                var dts = dl.querySelectorAll('dt');
                var dds = dl.querySelectorAll('dd');
                for (var i = 0; i < Math.min(dts.length, dds.length); i++) {
                    var key = dts[i].textContent.trim();
                    var val = dds[i].textContent.trim();
                    if (key && val) specs[key] = val;
                }
            });
            // Try label-value divs
            document.querySelectorAll('.product-info-main div, .product.info div').forEach(function(el) {
                var label = el.querySelector('.label, .title, strong');
                if (label) {
                    var val = el.textContent.replace(label.textContent, '').trim();
                    var key = label.textContent.trim().replace(':', '');
                    if (key && val && val.length < 500) specs[key] = val;
                }
            });
            return specs;
        """)
        if js_specs:
            for k, v in js_specs.items():
                if k not in dynamic_details:
                    dynamic_details[k] = v
    except:
        pass

    # ─── PROCESS DYNAMIC DETAILS ───
    # Extract known fields from dynamic_details
    dim_keys_found = []
    for key, value in list(dynamic_details.items()):
        key_lower = key.lower().strip()

        # General Dimensions
        if any(x in key_lower for x in ["general dimension", "product dimension", "dimension"]) and \
           any(c in value for c in ["W", "D", "H", "w", "d", "h"]):
            parsed = parse_product_dimensions(value)
            for col, val in parsed.items():
                if val:
                    data[col] = val
            dim_keys_found.append(key)

        # Seat dimensions (full string like "19.3" SW x 18.9" SD x 17.5" SH")
        elif "seat" in key_lower and any(x in value for x in ["SW", "SD", "SH", "sw", "sd", "sh"]):
            parsed = parse_seat_dimensions(value)
            for col, val in parsed.items():
                if val:
                    data[col] = val
            dim_keys_found.append(key)

        # Individual seat fields
        elif key_lower in ["seat height", "seat ht", "seat ht."]:
            data["Seat Height"] = re.sub(r'[^\d.]', '', str(value))
            dim_keys_found.append(key)
        elif key_lower in ["seat width", "seat wd", "seat wd."]:
            data["Seat Width"] = re.sub(r'[^\d.]', '', str(value))
            dim_keys_found.append(key)
        elif key_lower in ["seat depth", "seat dp", "seat dp."]:
            data["Seat Depth"] = re.sub(r'[^\d.]', '', str(value))
            dim_keys_found.append(key)
        elif key_lower in ["arm height"]:
            data["Arm Height"] = re.sub(r'[^\d.]', '', str(value))
            dim_keys_found.append(key)
        elif key_lower in ["arm width"]:
            data["Arm Width"] = re.sub(r'[^\d.]', '', str(value))
            dim_keys_found.append(key)

        # Weight — ONLY from Product Dimensions section
        # Key pattern: "Product 1 Weight (lbs.)" or "Product Weight (lbs.)"
        elif "product" in key_lower and "weight" in key_lower:
            data["Weight"] = parse_weight(value)
            dim_keys_found.append(key)

    # Remove processed dimension keys from dynamic_details
    for k in dim_keys_found:
        dynamic_details.pop(k, None)

    # ─── PACKAGING (Dynamic) ───
    try:
        # Look for packaging section
        packaging_section = None

        # Method 1: Find packaging tab/section
        tabs = driver.find_elements(By.CSS_SELECTOR,
            "#tab-label-packaging, #tab-label-additional, "
            "[data-role='content'], .product-info-tab, "
            "a[href='#packaging'], a[data-toggle='tab']"
        )
        for tab in tabs:
            if "packag" in tab.text.lower():
                try:
                    tab.click()
                    time.sleep(1)
                except:
                    pass
                break

        # Method 2: JS extraction for packaging
        pkg_data = driver.execute_script("""
            var pkg = {};
            var sections = document.querySelectorAll(
                '#packaging table tr, ' +
                '.packaging-info tr, ' +
                '[id*="packaging"] tr, ' +
                '[class*="packaging"] tr, ' +
                '[data-role="content"] table tr'
            );
            sections.forEach(function(row) {
                var cells = row.querySelectorAll('th, td');
                if (cells.length >= 2) {
                    var key = cells[0].textContent.trim();
                    var val = cells[1].textContent.trim();
                    if (key && val && key !== val) pkg[key] = val;
                }
            });
            // Also try div-based packaging
            document.querySelectorAll('[class*="packag"] div, [id*="packag"] div').forEach(function(el) {
                var label = el.querySelector('.label, strong, dt');
                if (label) {
                    var val = el.textContent.replace(label.textContent, '').trim();
                    var key = label.textContent.trim().replace(':', '');
                    if (key && val && val.length < 200) pkg[key] = val;
                }
            });
            return pkg;
        """)
        if pkg_data:
            for k, v in pkg_data.items():
                pkg_key = f"Pkg: {k}" if not k.lower().startswith("pkg") else k
                dynamic_packaging[pkg_key] = v
    except:
        pass

    # ─── PRODUCT FAMILY ID ───
    product_name = data.get("_product_name", "")
    if product_name:
        data["Product Family Id"] = product_family_from_name(product_name)

    return data, dynamic_details, dynamic_packaging


def main():
    print("🟢 Zuo Modern — Step 2: Detail Scraper")
    print("=" * 50)

    # Read Step 1 output
    df = pd.read_excel(INPUT_FILE)
    total = len(df)
    print(f"📋 Input: {total} products from {INPUT_FILE}\n")

    driver = build_driver()

    all_rows = []
    all_dynamic_keys = set()
    all_pkg_keys = set()

    try:
        for idx, row in df.iterrows():
            product_url = str(row.get("Source", "")).strip()
            product_name = str(row.get("Product Name", "")).strip()
            step1_image = str(row.get("Image URL", "")).strip()
            manufacturer = str(row.get("Manufacturer", "Zuo")).strip()

            if not product_url or product_url == "nan":
                print(f"[{idx+1}/{total}] ⚠ Skipping — no URL")
                continue

            print(f"[{idx+1}/{total}] 🔍 {product_name[:50]}...")

            try:
                driver.get(product_url)
                time.sleep(4)

                # Wait for page to load
                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR,
                            ".product-info-main, .product.info, "
                            "h1.page-title, [itemprop='name']"
                        ))
                    )
                except:
                    time.sleep(3)

                # Extract details
                details, dynamic, packaging = extract_product_details(driver)

                # Update product name from page if available
                try:
                    h1 = driver.find_element(By.CSS_SELECTOR, "h1.page-title span, h1 span[itemprop='name'], h1")
                    page_name = h1.text.strip()
                    if page_name:
                        product_name = page_name
                except:
                    pass

                # Build row — fixed columns first
                row_data = {
                    "Manufacturer": manufacturer,
                    "Source": product_url,
                    "Image URL": details.get("Image URL") or step1_image,
                    "Product Name": product_name,
                    "SKU": details.get("SKU", ""),
                    "Product Family Id": product_family_from_name(product_name),
                    "Description": details.get("Description", ""),
                    "Weight": details.get("Weight", ""),
                    "Width": details.get("Width", ""),
                    "Depth": details.get("Depth", ""),
                    "Diameter": details.get("Diameter", ""),
                    "Length": details.get("Length", ""),
                    "Height": details.get("Height", ""),
                }

                # Seat/Arm fields → dynamic (after fixed columns)
                seat_fields = {
                    "Seat Width": details.get("Seat Width", ""),
                    "Seat Depth": details.get("Seat Depth", ""),
                    "Seat Height": details.get("Seat Height", ""),
                    "Arm Height": details.get("Arm Height", ""),
                    "Arm Width": details.get("Arm Width", ""),
                }
                for k, v in seat_fields.items():
                    if v:
                        row_data[k] = v
                        all_dynamic_keys.add(k)

                # Add dynamic product details
                for k, v in dynamic.items():
                    col_name = f"Detail: {k}"
                    row_data[col_name] = v
                    all_dynamic_keys.add(col_name)

                # Add packaging
                for k, v in packaging.items():
                    row_data[k] = v
                    all_pkg_keys.add(k)

                all_rows.append(row_data)

                # Print summary
                sku_info = f"SKU: {row_data['SKU']}" if row_data['SKU'] else "SKU: -"
                dim_info = f"W:{row_data['Width']} D:{row_data['Depth']} H:{row_data['Height']}"
                print(f"         ✅ {sku_info} | {dim_info}")

            except Exception as e:
                print(f"         ❌ Error: {e}")
                all_rows.append({
                    "Manufacturer": manufacturer,
                    "Source": product_url,
                    "Image URL": step1_image,
                    "Product Name": product_name,
                })
                continue

            # Save progress every 20 products
            if (idx + 1) % 20 == 0:
                _save_progress(all_rows, FIXED_COLUMNS, all_dynamic_keys, all_pkg_keys)
                print(f"    💾 Progress saved ({idx+1}/{total})")

    finally:
        driver.quit()

    # Final save
    _save_progress(all_rows, FIXED_COLUMNS, all_dynamic_keys, all_pkg_keys)
    print(f"\n{'=' * 50}")
    print(f"✅ Done! {len(all_rows)} products saved to {OUTPUT_FILE}")


def _save_progress(rows, fixed_cols, dynamic_keys, pkg_keys):
    """Save current progress to Excel."""
    if not rows:
        return

    df = pd.DataFrame(rows)

    # Build column order: fixed → dynamic details (sorted) → packaging (sorted)
    all_cols = list(fixed_cols)
    for col in sorted(dynamic_keys):
        if col not in all_cols:
            all_cols.append(col)
    for col in sorted(pkg_keys):
        if col not in all_cols:
            all_cols.append(col)

    # Add any remaining columns not yet in list
    for col in df.columns:
        if col not in all_cols:
            all_cols.append(col)

    # Reorder, only keep columns that exist
    final_cols = [c for c in all_cols if c in df.columns]
    df = df[final_cols]

    df.to_excel(OUTPUT_FILE, index=False)


if __name__ == "__main__":
    main()