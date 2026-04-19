# -*- coding: utf-8 -*-
"""
Loloi Rugs - Step 1 (List Page Scraper) — robust image URL handling
Collects: Product URL, Image URL, Product Name, SKU
Page: https://www.loloirugs.com/collections/rugs-all
"""

import time
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException
)

BASE_URL = "https://www.loloirugs.com"
LIST_URL = "https://www.loloirugs.com/collections/rugs-all"

# ======= Configs you can tweak =======
HEADLESS = False               # set True for headless
CLICK_PAUSE = 1.2              # seconds between load-more clicks
SCROLL_PAUSE = 0.8             # seconds between scrolls to help lazy-load images
MAX_TRIES_NO_CHANGE = 5        # stop if product count doesn't grow after these tries
EXPECTED_TOTAL = 3234          # as provided
OUTPUT_PATH = Path(__file__).parent / "loloi_rugs_list.xlsx"
IMG_WAIT_PER_CARD = 0.25       # small wait after card comes into view
IMG_RETRIES = 4                # retries per card to replace data: placeholder
# =====================================

def make_driver(headless=HEADLESS):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--window-size=1400,1000")
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(120)
    return driver

def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    time.sleep(0.25)
    try:
        element.click()
        return True
    except (ElementClickInterceptedException, StaleElementReferenceException):
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            return False

def close_possible_popups(driver):
    selectors = [
        "button[aria-label='Close']",
        "button.cookie-accept, button#onetrust-accept-btn-handler",
        "div#klaviyo-bis-close, button.klaviyo-close-form",
    ]
    for sel in selectors:
        try:
            for e in driver.find_elements(By.CSS_SELECTOR, sel):
                if e.is_displayed():
                    safe_click(driver, e)
                    time.sleep(0.2)
        except Exception:
            pass

def load_all_products(driver):
    wait = WebDriverWait(driver, 20)
    driver.get(LIST_URL)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.product-showcase__container > div#searchResults")))

    no_change_strikes, last_count = 0, 0

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE)

        load_more = None
        try:
            load_more = driver.find_element(By.CSS_SELECTOR, "button#load-more")
            if load_more.is_displayed() and load_more.is_enabled():
                close_possible_popups(driver)
                if safe_click(driver, load_more):
                    time.sleep(CLICK_PAUSE)
                else:
                    break
            else:
                load_more = None
        except NoSuchElementException:
            load_more = None

        time.sleep(0.6)
        cards = driver.find_elements(By.CSS_SELECTOR, "div.product-card.relative.card-block__item")
        count = len(cards)
        print(f"Loaded items: {count}")

        if EXPECTED_TOTAL and count >= EXPECTED_TOTAL:
            break

        if load_more is None:
            if count == last_count:
                no_change_strikes += 1
            else:
                no_change_strikes = 0
            if no_change_strikes >= MAX_TRIES_NO_CHANGE:
                break

        last_count = count

    driver.execute_script("window.scrollTo(0, 0);")

def parse_srcset_for_best(srcset: str) -> str:
    """
    From a srcset string, return the URL with the largest width descriptor.
    """
    best_url, best_w = "", -1
    for part in srcset.split(","):
        part = part.strip()
        if not part:
            continue
        pieces = part.split()
        url = pieces[0]
        w = -1
        if len(pieces) > 1 and pieces[1].endswith("w"):
            try:
                w = int(pieces[1][:-1])
            except ValueError:
                w = -1
        if w > best_w:
            best_w, best_url = w, url
    return best_url or srcset.split(",")[0].strip().split()[0]

def get_image_url_from_card(driver, card) -> str:
    """
    Robustly returns a real image URL (not data:) from a product card.
    Priority:
      1) JS: img.currentSrc
      2) img.src (if not data:)
      3) img[data-src]
      4) <source> srcset / data-srcset (pick largest)
    """
    # Ensure the card is in view to trigger lazy load
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
    time.sleep(IMG_WAIT_PER_CARD)

    # Try several times in case the image swaps after we scroll
    for _ in range(IMG_RETRIES):
        try:
            img = card.find_element(By.CSS_SELECTOR, "img.product-card__image")
        except NoSuchElementException:
            try:
                img = card.find_element(By.CSS_SELECTOR, "picture img")
            except NoSuchElementException:
                img = None

        current_src = ""
        if img:
            try:
                current_src = driver.execute_script("return arguments[0].currentSrc || '';", img) or ""
            except Exception:
                current_src = ""

            if current_src and not current_src.startswith("data:"):
                return current_src.strip()

            # fallback to src / data-src
            src = (img.get_attribute("src") or "").strip()
            if src and not src.startswith("data:"):
                return src

            data_src = (img.get_attribute("data-src") or "").strip()
            if data_src and not data_src.startswith("data:"):
                return data_src

        # try <source> srcset
        try:
            source = card.find_element(By.CSS_SELECTOR, "picture source")
            srcset = (source.get_attribute("srcset") or source.get_attribute("data-srcset") or "").strip()
            if srcset:
                best = parse_srcset_for_best(srcset)
                if best and not best.startswith("data:"):
                    return best
        except NoSuchElementException:
            pass

        # give it a moment to swap from placeholder
        time.sleep(0.25)

    # If still nothing usable, return empty string
    return ""

def parse_cards(driver):
    cards = driver.find_elements(By.CSS_SELECTOR, "div.product-card.relative.card-block__item")
    rows, seen_urls = [], set()

    for idx, card in enumerate(cards):
        try:
            # Product anchor (name/url)
            name_a = card.find_element(By.CSS_SELECTOR, "a.js-product-name.product-card__name")
            href = name_a.get_attribute("href") or name_a.get_attribute("data-href")
            if href and href.startswith("/"):
                href = urljoin(BASE_URL, href)
            if not href:
                handle = name_a.get_attribute("data-product-handle")
                if handle:
                    href = urljoin(BASE_URL, f"/products/{handle}")

            # ✅ Full Product Name from `title` attribute → "Abi-01 Mh Stone / Multi"
            # Fallback to `data-label` with title-case if title attribute is missing
            product_name = (name_a.get_attribute("title") or "").strip()
            if not product_name:
                product_name = (name_a.get_attribute("data-label") or "").strip().title()

            # ✅ SKU from <span> inside the anchor → "Abi-01"
            sku = ""
            try:
                sku_span = name_a.find_element(By.CSS_SELECTOR, "span")
                sku = (sku_span.text or "").strip()
            except NoSuchElementException:
                # fallback: first word of data-label
                label = (name_a.get_attribute("data-label") or "").strip()
                sku = label.split()[0] if label else ""

            # Robust image URL
            img_url = get_image_url_from_card(driver, card)

            # Dedup on Product URL
            if href and href in seen_urls:
                continue
            seen_urls.add(href)

            rows.append({
                "Product URL": href,
                "Image URL": img_url,
                "Product Name": product_name,
                "SKU": sku,
            })

            if (idx + 1) % 100 == 0:
                print(f"Parsed {idx + 1} cards...")

        except StaleElementReferenceException:
            continue
        except Exception as e:
            print(f"Parse error on card #{idx}: {e}")
            continue

    return rows

def save_to_excel(rows, output_path: Path):
    df = pd.DataFrame(rows, columns=["Product URL", "Image URL", "Product Name", "SKU"])
    df = df.dropna(subset=["Product URL"]).drop_duplicates(subset=["Product URL"]).reset_index(drop=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Step1_List")
    print(f"Saved {len(df)} rows to: {output_path.resolve()}")

def main():
    driver = make_driver()
    try:
        load_all_products(driver)
        rows = parse_cards(driver)
        print(f"Total rows parsed: {len(rows)}")
        save_to_excel(rows, OUTPUT_PATH)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()