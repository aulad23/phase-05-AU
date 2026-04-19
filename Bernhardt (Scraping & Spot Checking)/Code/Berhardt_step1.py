import os
import time
import pandas as pd
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By

# ========== OUTPUT ==========
OUTPUT_FILE = "bernhardt_Desk Chairs.xlsx"

# ========== CATEGORY LINKS ==========
# 👉 Add as many category URLs as you want below
CATEGORY_URLS = [
    "https://www.bernhardt.com/products/luxury-home-office-chairs/#?RoomType=Workspace&$MultiView=Yes&Sub-Category=Chairs&orderBy=WorkspacePosition&context=shop&page={page}"
]

# ========== SCRAPER FUNCTIONS ==========
def extract_category_name(cat_url):
    """Extracts a readable category name from a given URL."""
    path = urllib.parse.urlparse(cat_url).path  # e.g. '/fabrics-leathers/fabrics/'
    category_name = path.strip("/").split("/")[-1] or "Unknown"
    return category_name.replace("-", " ").title()


def scrape_category(driver, base_url, category_name):
    """Scrapes one category page by page until no products remain."""
    all_rows = []
    page = 1

    while True:
        url = base_url.format(page=page)
        driver.get(url)
        time.sleep(8)

        products = driver.find_elements(By.CSS_SELECTOR, "div.grid-item")
        if not products:
            print(f"[!] No more products found for {category_name} (page {page}).")
            break

        for p in products:
            try:
                product_url = p.find_element(By.TAG_NAME, "a").get_attribute("href")
                img_el = p.find_element(By.CSS_SELECTOR, "img.grid-image")
                image_url = img_el.get_attribute("src") or img_el.get_attribute("data-src") or ""
                product_name = p.find_element(By.CSS_SELECTOR, "div.product-header").text.strip()

                # --- SKU logic ---
                sku_main = ""
                sku_components = ""
                try:
                    sku_main = p.find_element(By.CSS_SELECTOR, "span.product-id").text.strip()
                except:
                    pass
                try:
                    sku_components = p.find_element(By.CSS_SELECTOR, "div.meta-component.ng-binding").text.strip()
                except:
                    pass
                sku = f"{sku_main} | {sku_components}" if sku_components else sku_main
                # --- end SKU logic ---

                all_rows.append({
                    "Category": category_name,
                    "Product URL": product_url,
                    "Image URL": image_url,
                    "Product Name": product_name,
                    "SKU": sku
                })

            except Exception:
                continue

        print(f"[+] {category_name} - Page {page} scraped ({len(products)} products).")
        page += 1

    return all_rows


def main():
    driver = webdriver.Chrome()
    all_data = []

    for cat_url in CATEGORY_URLS:
        category_name = extract_category_name(cat_url)
        print(f"\n===== Scraping category: {category_name} =====")
        cat_rows = scrape_category(driver, cat_url, category_name)
        all_data.extend(cat_rows)

    driver.quit()

    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n[✓] Done! Saved {len(all_data)} products across {len(CATEGORY_URLS)} categories to {OUTPUT_FILE}")
    else:
        print("[x] No products scraped.")


if __name__ == "__main__":
    main()
