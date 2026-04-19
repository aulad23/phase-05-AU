# LIST PAGE SCRAPER — PAGE RANGE VERSION
# Scrapes ONLY the page range you set (no auto-detect of last page).
# Safe getters prevent 'NoneType' .text errors.
# Output always saved as: visualcomfort_chandelier.xlsx

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse, urljoin

# ----------------- CONFIG -----------------
CHROMEDRIVER_PATH = r"C:/chromedriver.exe"
BASE_URL    = "https://www.visualcomfort.com"
LISTING_URL = f"{BASE_URL}/us/c/fans"
PAGE_PARAM  = "p"

START_PAGE = 1
END_PAGE   = 4

HEADLESS = False
WAIT_CARD_SEC = 25
SCROLL_STEPS = 14
SCROLL_PAUSE = 0.6
# ------------------------------------------

def build_driver():
    opts = webdriver.ChromeOptions()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--start-maximized")
    return webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=opts)

driver = build_driver()

def add_or_replace_query_param(url: str, key: str, value: str) -> str:
    parts = list(urlparse(url))
    query = dict(parse_qsl(parts[4], keep_blank_values=True))
    query[key] = str(value)
    parts[4] = urlencode(query, doseq=True)
    return urlunparse(parts)

def wait_for_any_selectors(selectors, timeout=WAIT_CARD_SEC):
    def _cond(drv):
        for sel in selectors:
            if drv.find_elements(By.CSS_SELECTOR, sel):
                return True
        return False
    WebDriverWait(driver, timeout).until(_cond)

def smooth_scroll(step_pause=SCROLL_PAUSE, steps=SCROLL_STEPS):
    for _ in range(steps):
        driver.execute_script("window.scrollBy(0, 1200);")
        time.sleep(step_pause)

def get_soup(url: str) -> BeautifulSoup:
    driver.get(url)
    wait_for_any_selectors(
        [".product-card", "li.product-item", "ol.products li", ".products-grid"]
    )
    time.sleep(0.8)
    smooth_scroll()
    time.sleep(0.6)
    return BeautifulSoup(driver.page_source, "html.parser")

def select_cards(soup: BeautifulSoup):
    for sel in [
        ".products ol > li.product-card",
        "li.product-card",
        "ol.products li.product-item",
        "li.product-item",
        "li.item.product.product-item"
    ]:
        cards = soup.select(sel)
        if cards:
            return cards
    return []

def first_el(root, selectors):
    for sel in selectors:
        el = root.select_one(sel)
        if el:
            return el
    return None

def el_text(el):
    return el.get_text(strip=True) if el else ""

def el_attr(el, attrs):
    if not el:
        return ""
    for a in attrs:
        v = el.get(a)
        if v:
            return v
    return ""

def parse_card(card):
    a_link = first_el(card, ["a.product-item-link", ".name a", "a"])
    href   = urljoin(BASE_URL, (a_link.get("href") or "").split("?")[0]) if a_link else ""
    name   = el_text(first_el(card, [".product-item-link", ".name a", ".name", "a"]))

    sku    = el_text(first_el(card, [".sku p", ".sku", "[data-sku]"]))

    img_el = first_el(card, ["img.product-image-photo", ".product-image img", "img"])
    img_src = el_attr(img_el, ["src", "data-src", "data-original", "data-srcset", "srcset"])
    img_src = urljoin(BASE_URL, img_src) if img_src else ""

    price_el = first_el(card, [".price .price-final", ".price .price-wrapper", ".price", "[data-price-type='finalPrice']"])
    price = el_text(price_el)

    return {
        "Product URL": href,
        "Image URL": img_src,
        "Product Name": name,
        "SKU": sku,
        "List Price": price
    }

def scrape_page_range(start_page: int, end_page: int):
    results = []
    for page in range(start_page, end_page + 1):
        page_url = add_or_replace_query_param(LISTING_URL, PAGE_PARAM, page)
        print(f"\n=== Scraping listing page {page}/{end_page} ===")
        print(page_url)

        try:
            soup = get_soup(page_url)
        except Exception as e:
            print(f"[ERROR] Failed to load page {page}: {e}")
            continue

        cards = select_cards(soup)
        count = len(cards)
        print(f"[INFO] Found {count} product cards on page {page}")

        for idx, card in enumerate(cards, 1):
            try:
                data = parse_card(card)
                print(f"  - Product {idx}/{count}")
                print(f"    Name: {data['Product Name'] or '[missing]'}")
                print(f"    SKU: {data['SKU'] or '[missing]'}")
                print(f"    Link: {data['Product URL'] or '[missing]'}")
                results.append(data)
            except Exception as e:
                print(f"    [WARN] Card parse error: {e}")
                continue

    return results

try:
    rows = scrape_page_range(START_PAGE, END_PAGE)

    # Save in same folder as script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(script_dir, "visualcomfort_fans.xlsx")

    # *** FIX: Actually save the file ***
    pd.DataFrame(rows).to_excel(out_path, index=False)

    print(f"\n✅ Saved {len(rows)} rows to {out_path}")

finally:
    driver.quit()
