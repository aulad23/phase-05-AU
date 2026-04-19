# -*- coding: utf-8 -*-
# ==============================================================
# Chaddock.com — Step 1 (List pages → Excel)
# --------------------------------------------------------------
# Collects:
#   - Product URL
#   - Image URL
#   - Product Name
#   - SKU
#   - Category
#
# Requirements:
#   - Python 3.9+
#   - Chrome (must be installed on system)
#   - ChromeDriver (auto-handled via chromedriver-autoinstaller)
#
# Usage:
#   1) Save as: chaddock_step1.py
#   2) Run: python chaddock_step1.py
#   3) Output Excel: chaddock_products_step1.xlsx
# ==============================================================

import os
import sys
import subprocess
import time
import pandas as pd
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import chromedriver_autoinstaller


# ----------------- AUTO-INSTALL LIBS -----------------
required = ["selenium", "pandas", "openpyxl", "chromedriver-autoinstaller"]
for pkg in required:
    try:
        __import__(pkg)
    except ImportError:
        print(f"Installing missing package: {pkg} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
# -----------------------------------------------------


# ----------------- CONFIG -----------------
LIST_URL = (
    "https://chaddock.com/styles?ProdType=Nightstands|Tables%3aAccent+Tables|"
    "Tables%3aDrink+Tables|Tables%3aend+tables|Tables%3aSide+Tables&PageIndex=1"
)
CATEGORY = "Nightstands & Tables"
OUTPUT_XLSX = "chaddock_products_step1.xlsx"

# Scroll settings
SCROLL_PAUSE = 1.5
SCROLL_STAGNATION_LIMIT = 5
MAX_SCROLL_CYCLES = 500

# ================== IMPORTANT ==================
# Update selectors using browser Inspect tool
PRODUCT_CARD_SELECTOR = "div.product-card"
PRODUCT_URL_SELECTOR = "a.product-link"
IMAGE_SELECTOR = "img.product-image"
NAME_SELECTOR = "h3.product-name"
SKU_SELECTOR = "span.product-sku"
# ===============================================
# ------------------------------------------------


def connect_driver() -> webdriver.Chrome:
    """Initialize and return Chrome WebDriver (auto handles chromedriver)."""
    chromedriver_autoinstaller.install()
    opts = Options()
    opts.add_argument("--disable-gpu")
    opts.add_argument("--start-maximized")
    # opts.add_argument("--headless=new")  # Run headless if needed
    return webdriver.Chrome(options=opts)


def safe_text(el) -> str:
    """Get textContent safely from element."""
    if not el:
        return ""
    try:
        return (el.get_attribute("textContent") or "").strip()
    except StaleElementReferenceException:
        return ""


def safe_attribute(el, attr) -> str:
    """Get attribute safely from element."""
    if not el:
        return ""
    try:
        return (el.get_attribute(attr) or "").strip()
    except StaleElementReferenceException:
        return ""


def extract_item(card) -> dict:
    """Extract product info from one product card."""
    product_url, img_url, name_text, sku_text = "", "", "", ""

    try:
        a = card.find_element(By.CSS_SELECTOR, PRODUCT_URL_SELECTOR)
        product_url = a.get_attribute("href")
    except Exception:
        pass

    try:
        img = card.find_element(By.CSS_SELECTOR, IMAGE_SELECTOR)
        img_url = safe_attribute(img, "src") or safe_attribute(img, "data-src")
    except Exception:
        pass

    try:
        name_el = card.find_element(By.CSS_SELECTOR, NAME_SELECTOR)
        name_text = safe_text(name_el)
    except Exception:
        pass

    try:
        sku_el = card.find_element(By.CSS_SELECTOR, SKU_SELECTOR)
        sku_text = safe_text(sku_el)
    except Exception:
        pass

    return {
        "Product URL": product_url,
        "Image URL": img_url,
        "Product Name": name_text,
        "SKU": sku_text,
    }


def scroll_and_collect(driver):
    """Scroll through page, collect all product cards until stagnation."""
    collected = []
    seen_keys = set()
    last_len = 0
    stagnation = 0
    cycles = 0

    while stagnation < SCROLL_STAGNATION_LIMIT and cycles < MAX_SCROLL_CYCLES:
        cards = driver.find_elements(By.CSS_SELECTOR, PRODUCT_CARD_SELECTOR)

        new_items_count = 0
        for card in cards:
            data = extract_item(card)
            key = data.get("Product URL") or (data.get("Product Name"), data.get("SKU"))
            if key and key not in seen_keys:
                seen_keys.add(key)
                collected.append(data)
                new_items_count += 1

        if len(collected) == last_len:
            stagnation += 1
        else:
            stagnation = 0
            last_len = len(collected)

        cycles += 1
        print(f"Cycle {cycles}: Collected {new_items_count} new items. Total: {len(collected)}")

        # Scroll to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE)

    return collected


def main():
    driver = None
    try:
        driver = connect_driver()
        print(f"[{CATEGORY}] Harvesting from: {LIST_URL}")
        driver.get(LIST_URL)

        # Wait for product container
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, PRODUCT_CARD_SELECTOR))
        )

        all_rows = scroll_and_collect(driver)

        # Build DataFrame
        cols = ["Product URL", "Image URL", "Product Name", "SKU", "Category"]
        df = pd.DataFrame(all_rows, columns=cols)
        df["Category"] = CATEGORY

        # Remove duplicates
        df.drop_duplicates(subset=["Product URL"], inplace=True)

        # Save Excel
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"✅ Saved {len(df)} items to {OUTPUT_XLSX}")

    except Exception as e:
        print(f"❌ Error occurred: {e}")
    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()
