# -*- coding: utf-8 -*-
"""
Zuo Modern — Step 1: Category/Listing Page Scraper
===================================================
Collects: Manufacturer, Source (Product URL), Product Name, Image URL
from all given category pages on zuomod.com

Usage: python zuo_step1_listpage.py
Output: zuo_products_step1.xlsx
"""

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─── CONFIG ───
MANUFACTURER = "Zuo"

CATEGORIES = {
    "Dining Chairs": [
        "https://www.zuomod.com/indoor/dining/chairs",
        #"https://www.zuomod.com/indoor/bedroom/beds",
        #"https://www.zuomod.com/indoor/gaming/tables",
        #"https://www.zuomod.com/indoor/living/tables"
    ],
}

OUTPUT_FILE = "zuo_Dining_Chairs.xlsx"

# ─── SELENIUM SETUP ───
def build_driver():
    opts = Options()
    # opts.add_argument("--headless=new")  # uncomment for headless
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


def scroll_to_bottom(driver, pause=2, max_scrolls=50):
    """Scroll down to load all lazy-loaded products."""
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


def click_all_pages(driver):
    """
    Zuo uses pagination (20/40/80/160 per page).
    Try to set 'show all' or click through pages.
    Returns list of all product elements across pages.
    """
    all_products = []

    # Try to set items per page to maximum (160)
    try:
        limiter_links = driver.find_elements(By.CSS_SELECTOR, "a.limiter-options, select.limiter-options option, a[data-limit]")
        for link in limiter_links:
            text = link.text.strip()
            if "160" in text:
                link.click()
                time.sleep(5)
                break
    except:
        pass

    # Try select dropdown for per-page
    try:
        selects = driver.find_elements(By.CSS_SELECTOR, "select.limiter-options, select#limiter")
        for sel in selects:
            for option in sel.find_elements(By.TAG_NAME, "option"):
                if "160" in option.text:
                    option.click()
                    time.sleep(5)
                    break
    except:
        pass

    while True:
        scroll_to_bottom(driver, pause=2)
        time.sleep(2)

        # Extract products from current page
        products = driver.find_elements(By.CSS_SELECTOR,
            "li.product-item, div.product-item, .product-item-info, "
            "li.item.product, div.item.product"
        )

        if not products:
            # Fallback: try broader selectors
            products = driver.find_elements(By.CSS_SELECTOR,
                "ol.products li, ul.products li, div.products-grid li, "
                ".product-items li"
            )

        for p in products:
            try:
                # Product URL
                link_el = p.find_element(By.CSS_SELECTOR, "a.product-item-link, a.product-item-photo, a[href*='zuomod.com']")
                product_url = link_el.get_attribute("href") or ""

                # Product Name
                name_el = None
                try:
                    name_el = p.find_element(By.CSS_SELECTOR, "a.product-item-link")
                except:
                    try:
                        name_el = p.find_element(By.CSS_SELECTOR, ".product-item-name a, .product-name a, .product-item-link")
                    except:
                        pass
                product_name = name_el.text.strip() if name_el else ""

                # Image URL
                img_el = None
                try:
                    img_el = p.find_element(By.CSS_SELECTOR, "img.product-image-photo, img.photo")
                except:
                    try:
                        img_el = p.find_element(By.TAG_NAME, "img")
                    except:
                        pass

                image_url = ""
                if img_el:
                    image_url = (
                        img_el.get_attribute("data-src")
                        or img_el.get_attribute("src")
                        or ""
                    )

                if product_url:
                    all_products.append({
                        "product_url": product_url.strip(),
                        "product_name": product_name.strip(),
                        "image_url": image_url.strip(),
                    })
            except Exception as e:
                print(f"    ⚠ Error extracting product: {e}")
                continue

        # Check for next page
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR,
                "a.action.next, li.pages-item-next a, a.next"
            )
            next_url = next_btn.get_attribute("href")
            if next_url:
                print(f"    → Next page: {next_url}")
                driver.get(next_url)
                time.sleep(5)
            else:
                break
        except:
            break

    return all_products


def scrape_category(driver, category_name, urls):
    """Scrape all URLs for a single category."""
    category_products = []
    seen_urls = set()

    for url in urls:
        print(f"\n  📄 Loading: {url}")
        try:
            driver.get(url)
            time.sleep(5)

            # Wait for products to appear
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR,
                        "li.product-item, div.product-item, .product-item-info, "
                        "ol.products li, .product-items li"
                    ))
                )
            except:
                print(f"    ⚠ No products found on page, trying scroll...")

            products = click_all_pages(driver)
            print(f"    ✅ Found {len(products)} products on this page")

            for p in products:
                if p["product_url"] not in seen_urls:
                    seen_urls.add(p["product_url"])
                    p["category"] = category_name
                    category_products.append(p)

        except Exception as e:
            print(f"    ❌ Error loading {url}: {e}")
            continue

    return category_products


def main():
    print("🟢 Zuo Modern — Step 1: Category Scraper")
    print("=" * 50)

    driver = build_driver()
    all_products = []

    try:
        for cat_name, cat_urls in CATEGORIES.items():
            print(f"\n📁 Category: {cat_name}")
            products = scrape_category(driver, cat_name, cat_urls)
            all_products.extend(products)
            print(f"  📊 Total unique for '{cat_name}': {len(products)}")

    finally:
        driver.quit()

    # Remove duplicates by URL
    seen = set()
    unique = []
    for p in all_products:
        if p["product_url"] not in seen:
            seen.add(p["product_url"])
            unique.append(p)

    print(f"\n{'=' * 50}")
    print(f"📊 Total unique products: {len(unique)}")

    # Build DataFrame
    rows = []
    for p in unique:
        rows.append({
            "Manufacturer": MANUFACTURER,
            "Source": p["product_url"],
            "Product Name": p["product_name"],
            "Image URL": p["image_url"],
            "Category": p["category"],
        })

    df = pd.DataFrame(rows)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Saved to {OUTPUT_FILE}")
    print(f"\nSample (first 5):")
    print(df.head().to_string(index=False))


if __name__ == "__main__":
    main()