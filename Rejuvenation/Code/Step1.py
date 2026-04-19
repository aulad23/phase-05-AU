# -*- coding: utf-8 -*-
# Rejuvenation "Cabinets" Step-1 Scraper (MANUAL MODE)
# You scroll manually, Selenium only extracts loaded products
# Output: Excel only

import re
import time
import pandas as pd
from urllib.parse import urljoin

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException

# =======================
# CONFIG
# =======================
BASE_URL = "https://www.rejuvenation.com"
OUT_XLSX = "rejuvenation_Knobs.xlsx"

CATEGORY_LINKS = {
    "Knobs": [
        "https://www.rejuvenation.com/shop/hardware/cabinet-knobs/",
        #"https://www.rejuvenation.com/shop/hardware/bin-pulls-hardware/"
    ]
}

# =======================
# HELPERS
# =======================
def clean_text(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"[\x00-\x1f\x7f]", " ", str(s))
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_price(txt: str) -> str:
    if not txt:
        return ""
    txt = clean_text(txt)
    m = re.findall(r"\d[\d,]*", txt)
    return m[0] if m else ""

def uniq_preserve(seq):
    seen = set()
    out = []
    for x in seq:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out

# =======================
# DRIVER (DEBUG / MANUAL MODE)
# =======================
def build_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")

    # ✅ MANUAL / DEBUG MODE
    opts.add_experimental_option("detach", True)  # Chrome stays open
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    return webdriver.Chrome(options=opts)

def safe_find_text(el, by, sel):
    try:
        return clean_text(el.find_element(by, sel).text)
    except Exception:
        return ""

def safe_find_attr(el, by, sel, attr):
    try:
        return clean_text(el.find_element(by, sel).get_attribute(attr))
    except Exception:
        return ""

# =======================
# EXTRACTION
# =======================
def extract_product_cells(driver):
    rows = []
    cells = driver.find_elements(By.CSS_SELECTOR, "[data-component='Shop-GridItem'], .grid-item")

    for cell in cells:
        try:
            href = safe_find_attr(cell, By.CSS_SELECTOR, "a.product-image-link", "href")
            if not href:
                href = safe_find_attr(cell, By.CSS_SELECTOR, ".product-name a", "href")
            product_url = urljoin(BASE_URL, href) if href else ""

            img_url = safe_find_attr(cell, By.CSS_SELECTOR, "img[src]", "src")

            name = (
                safe_find_text(cell, By.CSS_SELECTOR, ".product-name a span")
                or safe_find_text(cell, By.CSS_SELECTOR, ".product-name")
            )

            prices = []
            for n in cell.find_elements(By.CSS_SELECTOR, ".suggested-price .amount, .product-price .amount"):
                p = normalize_price(n.text)
                if p:
                    prices.append(p)

            prices = uniq_preserve(prices)
            list_price = " - ".join(prices)

            if product_url and name:
                rows.append({
                    "Product URL": product_url,
                    "Image URL": img_url,
                    "Product Name": name,
                    "List Price": list_price
                })

        except (StaleElementReferenceException, Exception):
            continue

    return rows

# =======================
# MAIN
# =======================
def main():
    driver = build_driver()
    all_rows = []

    for cat, urls in CATEGORY_LINKS.items():
        for url in urls:
            print(f"\n[OPEN] {url}")
            driver.get(url)
            time.sleep(2)


            print("\n🟡 MANUAL MODE ENABLED")
            print("➡ Scroll the page MANUALLY until all products are loaded")
            print("➡ Lazy-load / Load More nijer moto korben")
            input("➡ Scroll shesh hole ENTER press korun...")

            rows = extract_product_cells(driver)
            print(f"[FOUND] {len(rows)} products")

            for r in rows:
                r["Category"] = cat
                r["Category URL"] = url
            all_rows.extend(rows)

    df = pd.DataFrame(all_rows)
    if not df.empty:
        df = df.drop_duplicates(subset=["Product URL"], keep="first")
        df = df[["Product URL", "Image URL", "Product Name", "List Price"]]

    df.to_excel(OUT_XLSX, index=False)
    print(f"\n✅ SAVED: {OUT_XLSX} | Rows: {len(df)}")

if __name__ == "__main__":
    main()
