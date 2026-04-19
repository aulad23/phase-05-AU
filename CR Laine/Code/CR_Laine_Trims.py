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
START_URL = "https://www.crlaine.com/trims/func/cat/48/category/View%20Trims"  # ✅ Change category link here

# ✅ Script jaigai jaigai run hok, oi folder-ei output save hobe
BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(BASE_FOLDER, "Crlaine_Trims.xlsx")

HEADLESS = True
WAIT_TIME = 3
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


def extract_product_data(page_source):
    """Parse and extract product cards from Fabric/Trim/Leather categories."""
    soup = BeautifulSoup(page_source, "html.parser")
    products = []

    # All product blocks have "style_thumbs text-center"
    for card in soup.select("div.style_thumbs.text-center"):
        try:
            # Product URL
            a_tag = card.find("a", href=True)
            href = a_tag.get("href") if a_tag else ""
            product_url = urljoin(BASE_URL, href.strip()) if href else ""

            # SKU extraction (numeric id)
            sku_match = re.search(r"/id/(\d+)", href) if href else None
            sku = sku_match.group(1) if sku_match else ""

            # Image URL
            img_tag = card.find("img", class_="pure-img")
            img_src = ""
            if img_tag:
                img_src = img_tag.get("src") or img_tag.get("lazyload") or ""
            image_url = urljoin(BASE_URL, img_src.strip()) if img_src else ""

            # Product Name
            name_tag = card.find("div", class_=re.compile(r"(fabName|stylename)"))
            product_name = ""
            if name_tag:
                product_name = name_tag.get_text(separator=" ", strip=True)
                product_name = re.sub(r"\s+", " ", product_name)

            if product_url and product_name:
                products.append({
                    "Product URL": product_url,
                    "Image URL": image_url,
                    "Product Name": product_name,
                    "SKU": sku
                })
        except Exception as e:
            print(f"⚠️ Error parsing card: {e}")
            continue

    return products


def main():
    print("🚀 Starting CR Laine Universal Scraper (Pagination Supported)...")
    driver = setup_driver(headless=HEADLESS)
    driver.get(START_URL)
    time.sleep(WAIT_TIME)

    all_products = []
    page_num = 1

    while True:
        print(f"🟦 Scraping page {page_num}...")
        time.sleep(WAIT_TIME)

        html = driver.page_source
        new_data = extract_product_data(html)
        print(f"   ↳ Found {len(new_data)} products on this page.")
        all_products.extend(new_data)

        # Pagination → click Next button
        try:
            next_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.nextPage.filterClick"))
            )
            style_attr = next_button.get_attribute("style") or ""
            if "display: none" in style_attr.replace(" ", "").lower():
                print("✅ No more pages. Stopping (Next button hidden).")
                break

            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(1)
            next_button.click()
            page_num += 1
            time.sleep(WAIT_TIME)
        except Exception:
            print("✅ No next page button or end reached.")
            break

    driver.quit()

    # -------- SAVE / APPEND LOGIC --------
    df_new = pd.DataFrame(all_products).drop_duplicates(subset=["Product URL"])
    print(f"\n🧮 New scraped unique products this run: {len(df_new)}")

    if os.path.exists(OUTPUT_PATH):
        print(f"📂 Existing file found: {OUTPUT_PATH}")
        try:
            df_old = pd.read_excel(OUTPUT_PATH)
            before_merge = len(df_old)
            df_final = pd.concat([df_old, df_new], ignore_index=True)
            df_final = df_final.drop_duplicates(subset=["Product URL"])
            after_merge = len(df_final)
            added_count = after_merge - before_merge
            print(f"➕ Appended to existing file. New rows actually added: {max(added_count, 0)}")
        except Exception as e:
            print(f"⚠️ Could not read existing file, creating fresh. Reason: {e}")
            df_final = df_new
    else:
        print("📄 No existing file found. Creating new file.")
        df_final = df_new

    df_final.to_excel(OUTPUT_PATH, index=False)

    print(f"\n✅ Done! Total products in file now: {len(df_final)}")
    print(f"📁 Saved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
