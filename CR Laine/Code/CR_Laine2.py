import os
import re
import time
import pandas as pd
from urllib.parse import urljoin
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# =========== CONFIG ===========

# 🔹 Script er folder (jeikhane .py file ache)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 🔹 Input/Output same path e save hobe
INPUT_FILE  = os.path.join(SCRIPT_DIR, "products_Ottomans.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "products_Ottomans_Details.xlsx")

BASE_URL = "https://www.crlaine.com"
HEADLESS = False        # False = visible Chrome
PAGE_WAIT = 2.0
BATCH_SIZE = 5          # save every 5 products
SCROLL_PAUSE_SECONDS = 1.0
# ==============================

def setup_driver(headless=False):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,1000")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def parse_dim_text(dim_text):
    width = depth = diameter = height = ""
    if not dim_text or not isinstance(dim_text, str):
        return width, depth, diameter, height

    s = dim_text.upper().replace("\u00A0", " ").strip()
    matches = re.findall(r'([\d\.]+)\s*(W|DIA|D|H)\b', s)
    for val, unit in matches:
        if unit == 'W' and not width:
            width = val
        elif unit == 'D' and not depth:
            depth = val
        elif unit == 'H' and not height:
            height = val
        elif unit == 'DIA' and not diameter:
            diameter = val

    if not diameter:
        m = re.search(r'([\d\.]+)\s*DIA\b', s)
        if m:
            diameter = m.group(1)

    clean = lambda x: re.search(r'[\d\.]+', x).group(0) if x and re.search(r'[\d\.]+', x) else ""
    return clean(width), clean(depth), clean(diameter), clean(height)

def extract_detail_fields_from_dimtable(dimtable_div):
    out = {
        "Description": "",
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Height": "",
        "Seat Height": "",
        "Arm Height": "",
        "Cushion": "",
        "Com": ""
    }
    if not dimtable_div:
        return out

    for d in dimtable_div.find_all("div"):
        label = d.find("span", class_="detailInfoLabel")
        if not label:
            continue

        label_text = label.get_text(strip=True).upper().rstrip(":")
        spans = d.find_all("span")
        val = spans[1].get_text(" ", strip=True) if len(spans) > 1 else ""
        if not val:
            continue

        if "OUTSIDE" in label_text:
            w, dep, dia, h = parse_dim_text(val)
            out["Width"], out["Depth"], out["Diameter"], out["Height"] = w, dep, dia, h

        elif "DESCRIPTION" in label_text:
            out["Description"] = val

        elif "WEIGHT" in label_text:
            m = re.search(r'([\d\.,]+)', val)
            out["Weight"] = m.group(1).replace(",", "") if m else val

        elif label_text == "SEAT":
            m = re.search(r'([\d\.]+)', val)
            out["Seat Height"] = m.group(1) if m else val

        elif label_text == "ARM":
            m = re.search(r'([\d\.]+)', val)
            out["Arm Height"] = m.group(1) if m else val

        elif "SEAT CUSHION" in label_text:
            out["Cushion"] = re.sub(r'\s+', ' ', val).strip()

        elif label_text == "COM":
            out["Com"] = re.sub(r'\s+', ' ', val).strip()

    return out

def scrape_product_detail(driver, url):
    data = {
        "Product URL": url,
        "Description": "",
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Height": "",
        "Seat Height": "",
        "Arm Height": "",
        "Cushion": "",
        "Com": ""
    }
    try:
        full_url = urljoin(BASE_URL, url) if url.startswith("/") else url
        driver.get(full_url)
        time.sleep(PAGE_WAIT)

        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/4);")
            time.sleep(SCROLL_PAUSE_SECONDS)
        except Exception:
            pass

        soup = BeautifulSoup(driver.page_source, "html.parser")
        dimtable = soup.find("div", class_="dimtable")
        details = extract_detail_fields_from_dimtable(dimtable)
        data.update(details)

    except Exception as e:
        print(f"⚠️ Failed: {url} -> {e}")
    return data

def save_partial(results, cols_to_keep):
    if not results:
        return

    base_cols = cols_to_keep + ["Product Family Id"]
    detail_cols = [
        "Description", "Weight", "Width", "Depth", "Diameter", "Height",
        "Seat Height", "Arm Height", "Cushion", "Com"
    ]
    all_cols = base_cols + detail_cols

    df_partial = pd.DataFrame(results)

    for col in all_cols:
        if col not in df_partial.columns:
            df_partial[col] = ""

    df_partial = df_partial[all_cols]
    df_partial.to_excel(OUTPUT_FILE, index=False)
    print(f"💾 Saved progress ({len(results)} rows) → {OUTPUT_FILE}")

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input file not found: {INPUT_FILE}")
        return

    df_input = pd.read_excel(INPUT_FILE)
    if "Product URL" not in df_input.columns:
        print("❌ Missing 'Product URL' column.")
        return

    cols_to_keep = [c for c in ["Product URL", "Image URL", "Product Name", "SKU"] if c in df_input.columns]

    results = []
    if os.path.exists(OUTPUT_FILE):
        try:
            prev = pd.read_excel(OUTPUT_FILE)
            if "Product URL" in prev.columns:
                done_urls = set(prev["Product URL"].astype(str))
                print(f"Resuming... {len(done_urls)} already done.")
                results = prev.to_dict(orient="records")
            else:
                done_urls = set()
        except Exception:
            done_urls = set()
    else:
        done_urls = set()

    driver = setup_driver(headless=HEADLESS)

    try:
        total = len(df_input)
        for idx, row in df_input.iterrows():
            url = str(row["Product URL"])
            if url in done_urls:
                print(f"[{idx+1}/{total}] Skipping (already done): {url}")
                continue

            print(f"[{idx+1}/{total}] Scraping: {url}")
            detail = scrape_product_detail(driver, url)

            out_row = {c: row.get(c, "") for c in cols_to_keep}

            # 🔴 ekhanei: Product Family Id = Product Name
            out_row["Product Family Id"] = row.get("Product Name", "")

            out_row.update(detail)
            results.append(out_row)

            if len(results) % BATCH_SIZE == 0:
                save_partial(results, cols_to_keep)
                print("✅ Batch saved, continuing...\n")

            time.sleep(0.8)
    finally:
        driver.quit()

    save_partial(results, cols_to_keep)
    print(f"\n🎯 Finished! Total {len(results)} products saved → {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
