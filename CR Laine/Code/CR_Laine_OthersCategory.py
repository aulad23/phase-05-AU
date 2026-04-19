import os
import re
import time
import pandas as pd
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ========== CONFIG ==========

BASE_URL = "https://www.crlaine.com"
START_URL = "https://www.crlaine.com/fabrics/CRL/cat/50/category/View%20Fabrics"

# 🔹 Script er folder (jekhane .py file ache)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 🔹 Output Excel same folder-e save hobe
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "Crlaine_Fabrics.xlsx")

HEADLESS = True           # True = invisible Chrome
WAIT_TIME = 3             # seconds after each page load
# ============================


def setup_driver(headless=True):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def extract_fabric_data(page_source):
    """Parse HTML and extract fabric product data."""
    soup = BeautifulSoup(page_source, "html.parser")
    products = []

    for card in soup.select("div.style_thumbs.text-center"):
        try:
            # SKU (from id="fabricSwatch_3721")
            fabric_div = card.find("div", id=re.compile(r"fabricSwatch_\d+"))
            sku = ""
            if fabric_div:
                m = re.search(r"fabricSwatch_(\d+)", fabric_div.get("id"))
                if m:
                    sku = m.group(1)

            # Product URL
            a_tag = card.find("a", class_="pageLoc")
            href = a_tag.get("href") if a_tag else ""
            product_url = urljoin(BASE_URL, href.strip()) if href else ""

            # Image URL
            img = card.find("img", class_="pure-img")
            img_src = img.get("src") or img.get("lazyload") if img else ""
            image_url = urljoin(BASE_URL, img_src.strip()) if img_src else ""

            # Product Name
            name_div = card.find("div", class_="fabName")
            product_name = ""
            if name_div:
                txt = name_div.get_text(separator=" ", strip=True)
                product_name = txt.split("(")[0].strip()

            if product_url and product_name:
                products.append({
                    "Product URL": product_url,
                    "Image URL": image_url,
                    "Product Name": product_name,
                    "SKU": sku
                })
        except Exception as e:
            print(f"⚠️ Error parsing fabric card: {e}")
            continue

    return products


def main():
    print("🚀 Starting CR Laine Fabric Scraper (Pagination Mode)...")
    driver = setup_driver(headless=HEADLESS)
    driver.get(START_URL)
    time.sleep(WAIT_TIME)

    all_products = []
    page_num = 1

    while True:
        print(f"🟦 Scraping page {page_num}...")
        time.sleep(WAIT_TIME)

        html = driver.page_source
        new_data = extract_fabric_data(html)
        print(f"   ↳ Found {len(new_data)} products on this page.")
        all_products.extend(new_data)

        # Try to find "Next" button
        try:
            next_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.nextPage.filterClick"))
            )
            if "display: none" in next_button.get_attribute("style"):
                print("✅ No more pages. Stopping.")
                break

            # Scroll to button and click
            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(1)
            next_button.click()
            page_num += 1
            time.sleep(WAIT_TIME)
        except Exception:
            print("✅ No next page button or end reached.")
            break

    driver.quit()

    # Save to Excel (same folder as script)
    df = pd.DataFrame(all_products).drop_duplicates(subset=["Product URL"])
    df.to_excel(OUTPUT_PATH, index=False)

    print(f"\n✅ Done! {len(df)} fabrics saved to:")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
