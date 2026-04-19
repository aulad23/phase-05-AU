import os
import time
import math
import sys
import traceback
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException, JavascriptException

# ====================== CONFIG ======================
CHROMEDRIVER_PATH = r"C:/chromedriver.exe"
BASE_URL = "https://www.jaipurliving.com/rugs.html"
TOTAL_PAGES = 43  # as specified
OUTPUT_PATH = "jaipur_list_step1.xlsx"

SCROLL_PAUSE_SECS = 0.8          # pause between scrolls
SCROLL_CHUNK_PX = 900            # scroll amount per chunk
STABLE_CHECKS_REQUIRED = 3       # times content height must stay unchanged to consider page fully loaded
IMG_RETRY_PER_ITEM = 3           # tries to resolve image URL per product
PAGE_LOAD_TIMEOUT = 45           # seconds
CONTENT_WAIT_TIMEOUT = 30        # seconds

# ================== DRIVER SETUP ====================
def build_driver():
    opts = Options()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option('useAutomationExtension', False)

    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver

# ================== HELPERS =========================
def wait_for_list(driver):
    """
    Waits for the product list to appear, with retries and fallback selectors.
    """
    found = False
    for attempt in range(3):
        try:
            print(f"[INFO] Waiting for product list (attempt {attempt + 1})...")
            WebDriverWait(driver, CONTENT_WAIT_TIMEOUT).until(
                EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    "ol.products.list.items.product-items, div.products.wrapper.grid.products-grid"
                ))
            )
            found = True
            break
        except TimeoutException:
            print(f"[WARN] Product list not found, retrying after refresh (attempt {attempt + 1})...")
            driver.refresh()
            time.sleep(5)
    if not found:
        raise TimeoutException("Product list not found after multiple retries.")

def get_current_height(driver):
    try:
        return driver.execute_script("return document.body.scrollHeight || document.documentElement.scrollHeight;")
    except JavascriptException:
        return None

def lazy_scroll_to_bottom(driver):
    """
    Scrolls in chunks until page height is stable for STABLE_CHECKS_REQUIRED iterations.
    """
    stable_count = 0
    last_height = get_current_height(driver)

    while True:
        driver.execute_script(f"window.scrollBy(0, {SCROLL_CHUNK_PX});")
        time.sleep(SCROLL_PAUSE_SECS)

        new_height = get_current_height(driver)
        if not new_height:
            break

        if new_height == last_height:
            stable_count += 1
        else:
            stable_count = 0

        if stable_count >= STABLE_CHECKS_REQUIRED:
            break

        last_height = new_height

    time.sleep(1.2)

def scroll_into_view(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element)
        time.sleep(0.3)
    except JavascriptException:
        pass

def parse_srcset(srcset_str):
    try:
        candidates = [part.strip() for part in srcset_str.split(",")]
        if not candidates:
            return None
        last = candidates[-1]
        pieces = last.split()
        if pieces:
            return pieces[0]
        return None
    except Exception:
        return None

def extract_image_url(img_el, driver):
    scroll_into_view(driver, img_el)
    time.sleep(0.25)

    for _ in range(IMG_RETRY_PER_ITEM):
        try:
            try:
                current_src = driver.execute_script("return arguments[0].currentSrc || '';", img_el)
            except JavascriptException:
                current_src = ""

            if current_src and not current_src.strip().lower().startswith("data:"):
                return current_src.strip()

            srcset = img_el.get_attribute("srcset") or img_el.get_attribute("data-srcset")
            if srcset:
                parsed = parse_srcset(srcset)
                if parsed and not parsed.lower().startswith("data:"):
                    return parsed

            src = img_el.get_attribute("src") or img_el.get_attribute("data-src")
            if src and not src.strip().lower().startswith("data:"):
                return src.strip()

            driver.execute_script("window.scrollBy(0, 150);")
            time.sleep(0.35)
        except StaleElementReferenceException:
            break
    return None

def collect_page_items(driver, page_url):
    items = []
    driver.get(page_url)
    wait_for_list(driver)
    lazy_scroll_to_bottom(driver)

    li_items = driver.find_elements(By.CSS_SELECTOR, "ol.products.list.items.product-items > li.item.product.product-item, div.products.wrapper.grid.products-grid li.item.product.product-item")

    for idx, li in enumerate(li_items, start=1):
        try:
            name_a = None
            try:
                name_a = li.find_element(By.CSS_SELECTOR, "strong.product.name.product-item-name a.product-item-link")
            except NoSuchElementException:
                try:
                    name_a = li.find_element(By.CSS_SELECTOR, "a.product-item-link")
                except NoSuchElementException:
                    name_a = None

            product_url = name_a.get_attribute("href").strip() if name_a else ""
            product_name = (name_a.text or "").strip() if name_a else ""

            img_el = None
            try:
                img_el = li.find_element(By.CSS_SELECTOR, "img.product-image-photo")
            except NoSuchElementException:
                img_el = None

            image_url = extract_image_url(img_el, driver) if img_el else None

            if not image_url:
                try:
                    a_main = li.find_element(By.CSS_SELECTOR, "a.product-item-link.mainImage, a.mainImage")
                    img2 = a_main.find_element(By.TAG_NAME, "img")
                    image_url = extract_image_url(img2, driver)
                except NoSuchElementException:
                    pass

            items.append({
                "Product URL": product_url,
                "Image URL": image_url or "",
                "Product Name": product_name
            })
        except Exception:
            traceback.print_exc(file=sys.stderr)
            continue

    return items

# ================== MAIN FLOW =======================
def main():
    driver = build_driver()
    all_rows = []

    try:
        for p in range(1, TOTAL_PAGES + 1):
            page_url = BASE_URL if p == 1 else f"{BASE_URL}?p={p}"
            print(f"\n[INFO] Scraping page {p}/{TOTAL_PAGES}: {page_url}")
            try:
                page_rows = collect_page_items(driver, page_url)
                print(f"[INFO] Collected {len(page_rows)} products on page {p}")
                all_rows.extend(page_rows)
            except TimeoutException as e:
                print(f"[ERROR] Skipping page {p}: Timeout waiting for products")
                continue
            time.sleep(1.0)
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    df = pd.DataFrame(all_rows, columns=["Product URL", "Image URL", "Product Name"])
    df = df[~(df["Product URL"].eq("") & df["Product Name"].eq(""))].reset_index(drop=True)

    os.makedirs(os.path.dirname(OUTPUT_PATH) or ".", exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False)
    print(f"\n[DONE] Wrote {len(df)} rows to: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
