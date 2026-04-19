"""
crlaine_scraper.py
Scrapes product cards from CR Laine Sofas / Loveseats / Sectionals.

Outputs Excel file in the SAME FOLDER as this script.

Columns:
- Product URL
- Image URL
- Product Name (combined stylename + last stylenumber)
- SKU
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import pandas as pd
import os

# === CONFIG ===
BASE_URL = "https://www.crlaine.com"
START_URLS = [
    "https://www.crlaine.com/products/CRL/cat/11/category/Ottomans",
    #"https://www.crlaine.com/products/CRL/cat/5/category/Loveseats_Settees",
    #"https://www.crlaine.com/products/CRL/cat/6/category/Sectionals"
]

# 🔹 Save Excel in the SAME folder as this .py file
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "products_Ottomans.xlsx")

SCROLL_PAUSE_SECONDS = 4  # lazy load wait
MAX_SCROLLS = 60          # safety cap
HEADLESS = False          # ❗ False = Chrome window visible
# ==============

def setup_driver(headless=True):
    options = Options()
    if headless:
        options.add_argument("--headless=new")  # for invisible mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1200")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def scroll_to_bottom(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

def collect_products_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    products = []

    cards = soup.select("div.style_thumbs, div[class*='style_thumbs']")
    if not cards:
        cards = soup.find_all("div", attrs={"stylename": True})

    for card in cards:
        try:
            # Product URL
            a = card.find("a", class_="pageLoc")
            href = a.get("href").strip() if a and a.get("href") else None
            product_url = urljoin(BASE_URL, href) if href else None

            # Image URL
            img = card.find("img")
            img_url = None
            if img:
                src = img.get("src") or img.get("lazyload") or img.get("data-src")
                if src:
                    img_url = urljoin(BASE_URL, src.strip())

            # Product Name
            name_div = card.find("div", class_="stylename")
            product_name = name_div.get_text(strip=True) if name_div else ""

            # SKU and additional info
            stylenumber_divs = card.find_all("div", class_="stylenumber")
            sku = None
            extra_text = ""
            if stylenumber_divs:
                # last stylenumber is extra text like "Queen Bed (65W)"
                extra_text = stylenumber_divs[-1].get_text(strip=True)

                # first stylenumber (SKU)
                def looks_like_sku(text):
                    if not text:
                        return False
                    t = text.strip()
                    return (
                        (len(t) <= 12 and any(ch.isdigit() for ch in t))
                        or ("-" in t and any(ch.isalpha() for ch in t))
                    )

                candidate = next(
                    (div.get_text(strip=True) for div in stylenumber_divs
                     if looks_like_sku(div.get_text(strip=True))),
                    None
                )
                sku = candidate or stylenumber_divs[0].get_text(strip=True)

            # Combine name + extra text if available
            if extra_text and extra_text not in product_name:
                product_name = f"{product_name} {extra_text}".strip()

            products.append({
                "Product URL": product_url,
                "Image URL": img_url,
                "Product Name": product_name,
                "SKU": sku
            })

        except Exception as e:
            print("Warning: failed to parse a card:", e)
            continue

    return products

def scrape_category(driver, url):
    print(f"\n=== Category start: {url} ===")
    driver.get(url)
    time.sleep(2)

    last_count = 0
    scrolls = 0
    while scrolls < MAX_SCROLLS:
        scrolls += 1
        scroll_to_bottom(driver)
        time.sleep(SCROLL_PAUSE_SECONDS)

        page_html = driver.page_source
        soup = BeautifulSoup(page_html, "html.parser")
        cards = soup.select("div.style_thumbs, div[class*='style_thumbs']")
        if not cards:
            cards = soup.find_all("div", attrs={"stylename": True})
        current_count = len(cards)
        print(f"Scroll {scrolls}: found {current_count} product cards so far.")
        if current_count == last_count:
            print("No new products loaded. Stopping scroll for this category.")
            break
        last_count = current_count

    products = collect_products_from_html(driver.page_source)
    print(f"✅ Parsed {len(products)} products from {url}")
    return products

def main():
    driver = setup_driver(headless=HEADLESS)
    all_products = []

    try:
        for url in START_URLS:
            cat_products = scrape_category(driver, url)
            all_products.extend(cat_products)

        print(f"\n🔢 Total raw products (all categories): {len(all_products)}")

        df = pd.DataFrame(all_products)

        # Drop rows without URL/SKU to avoid junk
        if "Product URL" in df.columns:
            df = df.dropna(subset=["Product URL"])
            df = df.drop_duplicates(subset=["Product URL"], keep="first")
        elif "SKU" in df.columns:
            df = df.dropna(subset=["SKU"])
            df = df.drop_duplicates(subset=["SKU"], keep="first")

        print(f"✅ Final unique products: {len(df)}")

        df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ Saved {len(df)} rows to {OUTPUT_FILE}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
