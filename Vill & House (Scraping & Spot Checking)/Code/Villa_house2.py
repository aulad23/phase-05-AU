import os
import re
import time
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ========== PATH SETTINGS ==========
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

input_file  = os.path.join(SCRIPT_DIR, "Vila_desks.xlsx")
output_file = os.path.join(SCRIPT_DIR, "Vila_desks_Final.xlsx")


# ========== BROWSER SETUP (VISIBLE WINDOW) ==========
def make_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(5)
    return driver


# ========== HELPER FUNCTIONS ==========
def get_text(tag):
    return tag.get_text(strip=True) if tag else ""


def extract_numeric_or_raw(value: str) -> str:
    """
    Value theke first numeric part tule dei (e.g. '23 inches' -> '23').
    Jodi number na thake tahole original value return.
    """
    m = re.search(r"(\d+(\.\d+)?)", value)
    if m:
        return m.group(1)
    return value.strip()


def parse_dimensions_fields(dimension_text: str):
    """
    Dimension theke:
    - 12W x 11D x 21H  -> Width, Depth, Height
    - Dia/DIA/DIAM/Diameter -> Diameter
    - L                -> Length
    - Cushion ...      -> Cushion (label chara)
    - 'Socket ...'     -> Socket (full line/part)
    - '...Watt...'     -> Wattage (full line/part)
    """
    width = depth = height = diameter = length = cushion = ""
    socket_val = ""
    wattage_val = ""

    if not dimension_text:
        return {
            "Width": width,
            "Depth": depth,
            "Height": height,
            "Diameter": diameter,
            "Length": length,
            "Cushion": cushion,
            "Socket": socket_val,
            "Wattage": wattage_val,
        }

    upper = dimension_text.upper()

    # numeric dimension pattern (e.g. 12W, 11 D, 21H, 45DIA, 20L)
    for match in re.finditer(r"(\d+(\.\d+)?)[ ]*(W|D|H|DIA|DIAM|DIAMETER|L)\b", upper):
        value = match.group(1)
        unit = match.group(3)
        if unit == "W":
            width = value
        elif unit == "D":
            depth = value
        elif unit == "H":
            height = value
        elif unit in ("DIA", "DIAM", "DIAMETER"):
            diameter = value
        elif unit == "L":
            length = value

    # Cushion (full phrase from 'Cushion...' but label chara)
    m_cush = re.search(r"(Cushion[^|]+)", dimension_text, flags=re.IGNORECASE)
    if m_cush:
        cushion_line = m_cush.group(1).strip()
        # Remove 'Cushion:' or 'Cushion Fill:' prefix
        cushion = re.sub(r"(?i)^Cushion(?:\s*Fill)?\s*:\s*", "", cushion_line).strip()

    # Socket & Wattage line gulo alada dhorbo (dimension_text split by "|")
    parts = [p.strip() for p in dimension_text.split("|") if p.strip()]
    for part in parts:
        low = part.lower()
        if "socket" in low and not socket_val:
            # e.g. "Socket Type A-E26 (3-Way)"
            socket_val = part.strip()
        if "watt" in low and not wattage_val:
            # e.g. "150 Watt Max"
            wattage_val = part.strip()

    return {
        "Width": width,
        "Depth": depth,
        "Height": height,
        "Diameter": diameter,
        "Length": length,
        "Cushion": cushion,
        "Socket": socket_val,
        "Wattage": wattage_val,
    }


def parse_specifications_fields(specs_text: str):
    """
    Specifications theke:
    - Weight (only from 'Item Weight' line → numeric part)
    - Seat Depth/Width/Height  (only numeric)
    - Seat Arm Width/Length/Height -> Arm Width/Length/Height (only numeric)
    - Com  (from 'Number of Yards to Recover' line mainly, only numeric)
    - Shade Details (from 'Shade Included?' line)
    (Wattage & Socket ekhane theke ashbe na)

    return: (fields_dict, weight_val)
    """
    seat_depth = ""
    seat_width = ""
    seat_height = ""
    arm_width = ""
    arm_length = ""
    arm_height = ""
    com_val = ""
    shade_details = ""
    weight_val = ""  # ONLY from "Item Weight" line

    if not specs_text:
        fields = {
            "Seat Depth": seat_depth,
            "Seat Width": seat_width,
            "Seat Height": seat_height,
            "Arm Width": arm_width,
            "Arm Length": arm_length,
            "Arm Height": arm_height,
            "Com": com_val,
            "Shade Details": shade_details,
        }
        return fields, weight_val

    parts = [p.strip() for p in specs_text.split("|") if p.strip()]
    for part in parts:
        if ":" not in part:
            continue
        label, value = part.split(":", 1)
        label_clean = label.strip().upper()
        value_clean = value.strip()

        # Weight → only from "Item Weight" line (numeric only)
        if "ITEM WEIGHT" in label_clean:
            num = extract_numeric_or_raw(value_clean)
            weight_val = num

        # Seat depth/width/height (numeric only)
        elif label_clean.startswith("SEAT DEPTH"):
            seat_depth = extract_numeric_or_raw(value_clean)
        elif label_clean.startswith("SEAT WIDTH"):
            seat_width = extract_numeric_or_raw(value_clean)
        elif label_clean.startswith("SEAT HEIGHT"):
            seat_height = extract_numeric_or_raw(value_clean)

        # Seat Arm Width/Length/Height → Arm Width/Length/Height (numeric only)
        elif "SEAT ARM WIDTH" in label_clean:
            arm_width = extract_numeric_or_raw(value_clean)
        elif "SEAT ARM LENGTH" in label_clean:
            arm_length = extract_numeric_or_raw(value_clean)
        elif "SEAT ARM HEIGHT" in label_clean:
            arm_height = extract_numeric_or_raw(value_clean)

        # (fallback jodi kono site e sudhu ARM WIDTH/HEIGHT thake)
        elif label_clean.startswith("ARM WIDTH") and not arm_width:
            arm_width = extract_numeric_or_raw(value_clean)
        elif label_clean.startswith("ARM LENGTH") and not arm_length:
            arm_length = extract_numeric_or_raw(value_clean)
        elif label_clean.startswith("ARM HEIGHT") and not arm_height:
            arm_height = extract_numeric_or_raw(value_clean)

        # Com: Number of Yards to Recover (numeric only)
        elif "NUMBER OF YARDS TO RECOVER" in label_clean:
            com_val = extract_numeric_or_raw(value_clean)
        elif label_clean.startswith("COM") and not com_val:
            com_val = extract_numeric_or_raw(value_clean)

        # Shade Details: Shade Included?
        elif "SHADE INCLUDED" in label_clean:
            shade_details = value_clean.strip()

    fields = {
        "Seat Depth": seat_depth,
        "Seat Width": seat_width,
        "Seat Height": seat_height,
        "Arm Width": arm_width,
        "Arm Length": arm_length,
        "Arm Height": arm_height,
        "Com": com_val,
        "Shade Details": shade_details,
    }
    return fields, weight_val


def parse_page(html, current_url):
    """page_source theke sob data parse kore ekta dict return kore"""
    soup = BeautifulSoup(html, "html.parser")

    # --- PRODUCT NAME / FAMILY ID ---
    name_tag = soup.select_one("h1.productView-title") or soup.select_one("h1")
    product_name = get_text(name_tag)
    product_family_id = product_name

    # --- SKU ---
    sku = ""
    sku_tag = (
        soup.select_one("span.sku")
        or soup.select_one("span.productView-info-value--sku")
        or soup.select_one("span#sku")
    )
    if sku_tag:
        sku = get_text(sku_tag)

    if not sku:
        for dl in soup.select("dl, div.productView-info"):
            dts = dl.select("dt")
            dds = dl.select("dd")
            for dt, dd in zip(dts, dds):
                label = get_text(dt)
                if "SKU" in label.upper():
                    value = get_text(dd)
                    if value:
                        sku = value
                        break
            if sku:
                break

    if not sku:
        for row in soup.select("tr"):
            th = row.find("th")
            td = row.find("td")
            if th and td and "SKU" in get_text(th).upper():
                val = get_text(td)
                if val:
                    sku = val
                    break

    # --- DESCRIPTION ---
    desc_blocks = soup.select("div.productView-description-text")
    description = ""
    if desc_blocks:
        all_texts = []
        for block in desc_blocks:
            ps = block.find_all("p")
            if ps:
                block_text = " ".join(p.get_text(" ", strip=True) for p in ps)
                all_texts.append(block_text)
            else:
                all_texts.append(block.get_text(" ", strip=True))
        description = " ".join(all_texts).strip()

    description = re.sub(r"\s+", " ", description)

    # --- DIMENSION (bullets-list theke raw text) ---
    dimension_text = ""
    bullets = soup.select("ul.bullets-list li")
    if bullets:
        dimension_text = " | ".join(get_text(li) for li in bullets if get_text(li))

    # --- SPECIFICATIONS (dl#additional-measurements theke raw text) ---
    specs_text = ""
    spec_dl = soup.select_one("dl#additional-measurements")
    if spec_dl:
        parts = []
        dts = spec_dl.select("dt")
        dds = spec_dl.select("dd")
        for dt, dd in zip(dts, dds):
            label = get_text(dt)
            value = get_text(dd)
            if label or value:
                parts.append(f"{label}: {value}")
        specs_text = " | ".join(parts)

    # --- COLOR ---
    color = ""
    color_span = soup.select_one("span.color-under-bullet")
    if color_span:
        txt = color_span.get_text(" ", strip=True)
        txt = re.sub(r"(?i)^color:\s*", "", txt).strip()
        color = txt

    if not color:
        main_swatch = soup.select_one(
            "ul.productAvailableColors li.availableColorsMain a.swatchColor"
        )
        if main_swatch:
            color = (
                main_swatch.get("name")
                or main_swatch.get("title")
                or ""
            )
            if not color:
                img = main_swatch.select_one("img")
                if img:
                    color = img.get("alt", "")

    # --- MAIN IMAGE URL (variation-wise) ---
    image_url = ""
    img_tag = (
        soup.select_one("figure.productView-image img")
        or soup.select_one("img#productDefaultImage")
        or soup.select_one("div.productView-image img")
    )
    if img_tag:
        src = img_tag.get("src") or ""
        if src:
            image_url = urljoin(current_url, src)

    # --- PARSED FIELDS FROM DIMENSION & SPECS ---
    dim_fields = parse_dimensions_fields(dimension_text)
    spec_fields, weight_val = parse_specifications_fields(specs_text)

    print(
        f"      🏷️ Name: {'✅' if product_name else '❌'} "
        f"| SKU: {sku or '-'} "
        f"| Color: {color or '-'} "
        f"| Desc: {'✅' if description else '❌'} "
        f"| Dim: {'✅' if dimension_text else '❌'} "
        f"| Specs: {'✅' if specs_text else '❌'} "
        f"| Img: {'✅' if image_url else '❌'}"
    )

    base = {
        "Product URL": current_url,
        "Image URL": image_url,
        "Product Name": product_name,
        "SKU": sku,
        "Product Family Id": product_family_id,
        "Description": description,
        "Weight": weight_val,          # numeric from "Item Weight"
        "Dimension": dimension_text,
        "Specifications": specs_text,
        "Color": color,
    }

    # dim_fields: Width, Depth, Height, Diameter, Length, Cushion, Socket, Wattage
    base.update(dim_fields)
    # spec_fields: Seat..., Arm..., Com, Shade Details
    base.update(spec_fields)
    return base


def get_variation_urls_from_html(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    urls = []

    finish_ul = (
        soup.select_one("div.product-view-finish-container ul.productAvailableColors")
        or soup.select_one("ul.productAvailableColors")
    )
    if not finish_ul:
        return []

    for a in finish_ul.select("a.swatchColor"):
        href = a.get("href")
        if not href:
            continue
        full = urljoin(base_url, href)
        if full.rstrip("/") != base_url.rstrip("/"):
            urls.append(full)

    urls = list(dict.fromkeys(urls))
    return urls


def scrape_product_with_variations(driver, base_url):
    rows = []

    print(f"   🌐 Fetching base page in browser...")
    driver.get(base_url)

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.productView"))
        )
    except Exception:
        print("   ⚠️ main product container time-out. Still trying to parse...")
        time.sleep(3)

    # main selected (manually selected finish)
    print(f"   🧾 Scraping main selected finish (current page)...")
    html = driver.page_source
    main_row = parse_page(html, driver.current_url)
    rows.append(main_row)

    # variations
    variation_urls = get_variation_urls_from_html(html, base_url)

    if variation_urls:
        print(f"   🎨 Found {len(variation_urls)} extra variation URLs.")
    else:
        print(f"   🎨 No extra variations found from HTML.")
        return rows

    for vidx, vurl in enumerate(variation_urls, start=1):
        print(f"   ➜ [{vidx}/{len(variation_urls)}] Variation URL: {vurl}")
        driver.get(vurl)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.productView"))
            )
        except Exception:
            print("      ⚠️ productView not found quickly, parsing anyway...")
            time.sleep(3)

        v_html = driver.page_source
        v_row = parse_page(v_html, driver.current_url)
        rows.append(v_row)
        time.sleep(1.5)

    return rows


# ========== MAIN SCRIPT ==========
if not os.path.exists(input_file):
    print(f"❌  Input file not found at: {input_file}")
    raise SystemExit()

df = pd.read_excel(input_file)

if "Product URL" in df.columns:
    url_col = "Product URL"
elif "Product Url" in df.columns:
    url_col = "Product Url"
else:
    print("❌  Excel e 'Product URL' or 'Product Url' column paoa jai ni.")
    print(f"Columns found: {list(df.columns)}")
    raise SystemExit()

urls = df[url_col].dropna().tolist()

print("\n🚀 Starting detailed product + variation scraping (VISIBLE Chrome)...")
print("=================================================================\n")

driver = make_driver()
results = []

# final column order
desired_cols = [
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
    "Color",
    "Cushion",
    "Seat Depth",
    "Seat Width",
    "Seat Height",
    "Arm Width",
    "Arm Length",
    "Arm Height",
    "Shade Details",
    "Com",
    "Wattage",
    "Socket",
]

BATCH_SIZE = 5  # 5 ta base product por por save


def save_batch(results_list, upto_products):
    if not results_list:
        return
    df_tmp = pd.DataFrame(results_list)
    for col in desired_cols:
        if col not in df_tmp.columns:
            df_tmp[col] = ""
    df_tmp = df_tmp[desired_cols]
    df_tmp.to_excel(output_file, index=False)
    print(f"💾 Batch saved after {upto_products} products. Total rows: {len(df_tmp)}")


try:
    total = len(urls)
    for i, base_url in enumerate(urls, start=1):
        print(f"[{i}/{total}] 🔗 Base Product: {base_url}")
        rows = scrape_product_with_variations(driver, base_url)
        results.extend(rows)
        print(f"   ✅ Rows added from this product (including variations): {len(rows)}")
        print("-----------------------------------------------------------------\n")
        # batch save every 5 products, and also last batch
        if (i % BATCH_SIZE == 0) or (i == total):
            save_batch(results, i)
        time.sleep(2)
finally:
    driver.quit()

print("=================================================================")
print("🎯 Scraping completed successfully!")
print(f"📦 Total Rows (all variations): {len(results)}")
print(f"💾 Final Excel saved to: {output_file}")
print("=================================================================")
