import logging
import os
import re
os.environ["WDM_LOG"] = "0"

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time

logging.getLogger("selenium").setLevel(logging.CRITICAL)

BASE_URL = "https://jonathanbrowninginc.com"

def setup_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    options.add_argument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(log_path=os.devnull)
    return webdriver.Chrome(options=options, service=service)


def parse_weight(weight_str):
    """Extract only the numeric value from weight, e.g. '15 lbs' → '15'"""
    if not weight_str:
        return ''
    match = re.search(r'[\d.]+', str(weight_str))
    return match.group() if match else ''


def parse_electrical(electrical_str):
    """
    Parse electrical strings like:
      1 x E26 12W LED bulb only       → Socket: E26,       Wattage: 12W
      1 x E26 A19 12W LED bulb only   → Socket: E26 A19,   Wattage: 12W
      1 x GU10 MR16 6W LED bulb only  → Socket: GU10 MR16, Wattage: 6W
      1 x E26 40W - 75W Max - LED...  → Socket: E26,       Wattage: 40W - 75W Max
      1 x 6W LED Only                 → Socket: (empty),   Wattage: 6W
    """
    socket = wattage = ""
    if not electrical_str:
        return socket, wattage

    # Remove leading 'N x ' prefix (e.g., '1 x ', '2 x ')
    text = re.sub(r'^\d+\s*x\s*', '', str(electrical_str).strip(), flags=re.IGNORECASE)

    # Find position of first wattage (e.g., 12W, 3.5W, 40W)
    watt_match = re.search(r'(\d+\.?\d*W)', text, re.IGNORECASE)
    if not watt_match:
        return socket, wattage

    # Socket = everything before the wattage
    socket = text[:watt_match.start()].strip()

    # Wattage = from first watt match; capture range if present (e.g., '40W - 75W Max')
    # Stop before ' - LED' or ' LED'
    watt_section = text[watt_match.start():]
    watt_end = re.search(r'\s*-?\s*LED\b', watt_section, re.IGNORECASE)
    if watt_end:
        wattage = watt_section[:watt_end.start()].strip()
    else:
        wattage = watt_match.group(1)

    return socket, wattage


def parse_dimension(dimension_str):
    """
    Parse dimension strings like:
      20 DIA x 79.5 H
      10 W x 40 D x 54 OAH
      10 Dia x 60 H
      27.75" W x 8" D x 44" H
      5" Diam. x 25 H
      60 L x 16.75 W x 60 OAH
    Returns: width, depth, diameter, length, height (empty string if not found)
    """
    width = depth = height = diameter = length = ""

    if not dimension_str:
        return width, depth, diameter, length, height

    # Remove ASCII and ALL Unicode quote/apostrophe variants
    dim = re.sub(r'[\u2018\u2019\u201a\u201b\u201c\u201d\u201e\u201f\u0022\u0027]', '', str(dimension_str)).strip()
    # Remove non-breaking spaces
    dim = dim.replace('\xa0', ' ').strip()

    # Find all (number)(unit) pairs — unit may include trailing dot e.g. "Diam."
    parts = re.findall(r'([\d.]+)\s*([A-Za-z.]+)', dim)

    for value, unit in parts:
        u = unit.upper().rstrip('.')   # strip trailing dot: "DIAM." → "DIAM"
        if u in ("DIA", "DIAM", "DIAMETER"):
            diameter = value
        elif u == "W":
            width = value
        elif u == "D":
            depth = value
        elif u == "L":
            length = value
        elif u in ("H", "OAH"):
            height = value

    return width, depth, diameter, length, height


def parse_sub_content(sub_content_html, debug=False):
    soup = BeautifulSoup(sub_content_html, "html.parser")
    full_text = soup.get_text(separator="\n")
    lines = [l.strip() for l in full_text.split("\n") if l.strip()]

    if debug:
        print(f"    [DEBUG] All lines found: {lines}")

    description = sku = dimension = weight = finish_text = shade = electrical = ""

    bold_labels = {
        "finishes", "model #", "dimensions", "net wt.",
        "shade", "electrical", "ceiling plate / j-box"
    }

    desc_lines = []
    for line in lines:
        if line.lower() in bold_labels:
            break
        desc_lines.append(line)
    description = " ".join(desc_lines).strip()

    i = 0
    while i < len(lines):
        label = lines[i].lower().strip()
        next_val = lines[i + 1].strip() if i + 1 < len(lines) else ""

        if label == "model #":
            sku = next_val
        elif label == "dimensions":
            dimension = next_val
        elif label == "net wt.":
            weight = next_val
        elif label == "finishes":
            if next_val.lower() != "view all finishes":
                finish_text = next_val
            elif i + 2 < len(lines):
                finish_text = lines[i + 2].strip()
        elif label == "shade":
            shade = next_val
        elif label == "electrical":
            electrical = next_val
        i += 1

    return description, sku, dimension, weight, finish_text, shade, electrical


def get_finishes_from_popup(driver):
    try:
        btns = driver.find_elements(By.CLASS_NAME, "view-finishes")
        if not btns:
            return ""
        driver.execute_script("arguments[0].click();", btns[0])
        time.sleep(2)
        finish_wrapper = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "finish-wrapper"))
        )
        finish_names = [
            el.text.strip()
            for el in finish_wrapper.find_elements(By.CLASS_NAME, "name")
            if el.text.strip()
        ]
        return ", ".join(finish_names)
    except Exception:
        return ""


def scrape_product(driver, url, product_name, image_url="", debug=False):
    print(f"  Scraping: {product_name}")
    try:
        driver.get(url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "product-name"))
        )
        time.sleep(1.5)
    except Exception as e:
        print(f"  Page load error: {e}")
        return None

    soup = BeautifulSoup(driver.page_source, "html.parser")
    product_family_id = url.split("?deep=")[1] if "?deep=" in url else ""

    description = sku = dimension = weight = finish = shade = electrical = ""
    info_li = soup.find("li", class_="information")
    if info_li:
        sub = info_li.find("div", class_="sub-content")
        if sub:
            description, sku, dimension, weight, finish, shade, electrical = parse_sub_content(str(sub), debug=debug)

    finish_popup = get_finishes_from_popup(driver)
    if finish_popup:
        finish = finish_popup
        print(f"    Finish (popup): {finish[:80]}")
    else:
        print(f"    Finish (text): {finish[:80]}")

    if shade:
        print(f"    Shade: {shade}")
    if electrical:
        print(f"    Electrical: {electrical}")

    # ── Parse weight (remove unit) ─────────────────────────────────────────────
    weight = parse_weight(weight)

    # ── Parse electrical into separate columns ─────────────────────────────────
    socket, wattage = parse_electrical(electrical)
    print(f"    Electrical: '{electrical}' → Socket={socket}, Wattage={wattage}")

    # ── Parse dimension into separate columns ──────────────────────────────────
    width, depth, diameter, length, height = parse_dimension(dimension)
    print(f"    Dimension: '{dimension}' → W={width}, D={depth}, H={height}, DIA={diameter}, L={length}")

    # ── Image URL comes from input Excel ──────────────────────────────────────
    print(f"    Image URL: {image_url[:80] if image_url else 'NOT FOUND'}")

    tearsheet_link = ""
    tearsheet_li = soup.find("li", class_="tearsheet")
    if tearsheet_li:
        a_tag = tearsheet_li.find("a")
        if a_tag:
            href = a_tag.get("href", "")
            tearsheet_link = BASE_URL + href if href.startswith("/") else href

    return {
        "Product URL": url,
        "Image URL": image_url,
        "Product Name": product_name,
        "SKU": sku,
        "Product Family Id": product_family_id,
        "Description": description,
        "Weight": weight,
        "Dimension": dimension,
        "Width": width,
        "Depth": depth,
        "Diameter": diameter,
        "Length": length,
        "Height": height,
        "Shade Details": shade,
        "Electrical": electrical,
        "Socket": socket,
        "Wattage": wattage,
        "Tearsheet Link": tearsheet_link,
    }


# ── MAIN ──────────────────────────────────────────────────────────────────────

input_file = "jonathan_browning_flush-mounts.xlsx"
output_file = "jonathan_browning_flush-mounts_details.xlsx"

DEBUG_FIRST_PRODUCT = True

print(f"Reading {input_file}...")
df_input = pd.read_excel(input_file)
print(f"Total products: {len(df_input)}\n")

driver = setup_driver()

results = []
for idx, row in df_input.iterrows():
    debug = DEBUG_FIRST_PRODUCT and idx == 0
    result = scrape_product(
        driver,
        row["Product URL"],
        row["Product Name"],
        image_url=str(row["Image URL"]) if "Image URL" in df_input.columns and pd.notna(row["Image URL"]) else "",
        debug=debug
    )
    if result:
        results.append(result)
    time.sleep(1)

driver.quit()

columns = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family Id",
    "Description", "Weight", "Dimension", "Width", "Depth", "Diameter", "Length", "Height",
    "Shade Details", "Electrical", "Socket", "Wattage", "Tearsheet Link"
]
df_output = pd.DataFrame(results, columns=columns)
df_output = df_output.fillna('')
df_output.to_excel(output_file, index=False)

print(f"\nDone! {len(df_output)} products saved to '{output_file}'")
print(df_output.head())