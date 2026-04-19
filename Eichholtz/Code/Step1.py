"""
Eichholtz Product Scraper
Usage: python eichholtz_scraper.py
Input URL hardcoded (can be changed below)
Output: eichholtz_products.xlsx
"""

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===================== CONFIG =====================
# Single URL or multiple URLs — both work
INPUT_URL = [
    "https://www.eichholtz.com/en/collection/furniture/rugs-carpets.html",
    #"https://www.eichholtz.com/en/collection/furniture/rugs-carpets.html",
    #"https://www.eichholtz.com/en/collection/furniture/tables/bars-butler-trays.html"
]
OUTPUT_FILE = "eichholtz_Rugs.xlsx"
WAIT_SECONDS = 5
# ==================================================


def get_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)


def get_total_pages(driver):
    try:
        pagination = driver.find_elements(By.CSS_SELECTOR, "li.ais-Pagination-item--page a")
        if not pagination:
            return 1
        pages = []
        for a in pagination:
            try:
                pages.append(int(a.text.strip()))
            except ValueError:
                pass
        return max(pages) if pages else 1
    except Exception:
        return 1


def scrape_page(driver):
    products = []
    wait = WebDriverWait(driver, 15)

    wait.until(EC.presence_of_all_elements_located(
        (By.CSS_SELECTOR, "div[itemprop='itemListElement']")
    ))
    time.sleep(WAIT_SECONDS)

    items = driver.find_elements(By.CSS_SELECTOR, "div[itemprop='itemListElement']")
    for item in items:
        try:
            try:
                product_url = item.find_element(By.CSS_SELECTOR, "a.result").get_attribute("href")
                product_url = product_url.split("?")[0]
            except Exception:
                product_url = ""

            try:
                image_url = item.find_element(By.CSS_SELECTOR, "img.hover-image").get_attribute("src")
            except Exception:
                image_url = ""

            try:
                product_name = item.find_element(By.CSS_SELECTOR, "h3.product-item-link").text.strip()
            except Exception:
                product_name = ""

            try:
                sku = item.find_element(By.CSS_SELECTOR, "span.text-xs").text.strip()
            except Exception:
                sku = ""

            if product_name or sku:
                products.append({
                    "Product URL": product_url,
                    "Image URL": image_url,
                    "Product Name": product_name,
                    "SKU": sku,
                })
        except Exception as e:
            print(f"  [!] Item parse error: {e}")
            continue

    return products


def build_page_url(base_url, page_num):
    if page_num == 1:
        return base_url
    base = base_url.rstrip("/")
    return f"{base}?page={page_num}"


def scrape_url(driver, base_url):
    all_products = []
    driver.get(base_url)
    time.sleep(WAIT_SECONDS)

    total_pages = get_total_pages(driver)
    print(f"  Total pages: {total_pages}")

    for page in range(1, total_pages + 1):
        url = build_page_url(base_url, page)
        print(f"  Scraping page {page}/{total_pages}: {url}")
        if page > 1:
            driver.get(url)
            time.sleep(WAIT_SECONDS)

        page_products = scrape_page(driver)
        print(f"    Found {len(page_products)} products")
        all_products.extend(page_products)

    return all_products


def main():
    # Support both single string and list of URLs
    urls = INPUT_URL if isinstance(INPUT_URL, list) else [INPUT_URL]

    print(f"Total URLs to scrape: {len(urls)}")
    driver = get_driver()
    all_products = []

    try:
        for i, url in enumerate(urls, 1):
            print(f"\n[URL {i}/{len(urls)}] {url}")
            try:
                products = scrape_url(driver, url)
                all_products.extend(products)
            except Exception as e:
                print(f"  [!] Failed to scrape {url}: {e}")
                continue
    finally:
        driver.quit()

    if all_products:
        df = pd.DataFrame(all_products, columns=["Product URL", "Image URL", "Product Name", "SKU"])
        # Remove duplicate products (same SKU)
        df.drop_duplicates(subset=["SKU"], keep="first", inplace=True)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nDone! {len(df)} unique products saved to '{OUTPUT_FILE}'")
    else:
        print("\nNo products found!")


if __name__ == "__main__":
    main()