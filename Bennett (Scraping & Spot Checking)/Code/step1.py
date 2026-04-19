import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import time

# 🏠 Base URL
base_domain = "https://www.bennetttothetrade.com"

# 📂 যে Collection থেকে ডেটা নিতে চাও
base_collection_url = f"{base_domain}/collections/sideboards"
# উদাহরণ: base_collection_url = f"{base_domain}/collections/coffee-tables"

def scrape_collection(base_collection_url):
    page = 1
    all_data = []

    while True:
        url = f"{base_collection_url}?page={page}"
        print(f"🔎 Scraping page {page} — {url}")
        response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"})
        if response.status_code != 200:
            print("⚠️ Page not found or finished scraping.")
            break

        soup = BeautifulSoup(response.text, 'html.parser')
        products = soup.select('div.card-wrapper.product-card-wrapper')

        if not products:
            print("✅ No more products found. Stopping.")
            break

        for product in products:
            # --- Image URL ---
            img_tag = product.select_one('div.media img')
            image_url = None
            if img_tag and img_tag.get('src'):
                image_url = img_tag['src']
                if image_url.startswith("//"):
                    image_url = "https:" + image_url

            # --- Product URL, Name, SKU ---
            a_tag = product.select_one('h3.card__heading.h5 a.full-unstyled-link')
            product_url = name = sku = None
            if a_tag:
                href = a_tag.get('href')
                product_url = base_domain + href if href else None
                name = a_tag.get_text(strip=True)
                sku = name  # Bennett সাইটে SKU আর নাম একই থাকে (যেমন BIZ1002)

            # --- Save each product ---
            if product_url:
                all_data.append({
                    "Product URL": product_url,
                    "Image URL": image_url,
                    "Product Name": name,
                    "SKU": sku
                })

        print(f"✅ Page {page} scraped successfully.")
        page += 1
        time.sleep(1.5)  # সামান্য delay anti-blocking এর জন্য

    # --- Save to Excel file in the SAME folder where this script is located ---
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Script folder
    file_path = os.path.join(script_dir, "bennet_sideboards.xlsx")

    df = pd.DataFrame(all_data)
    df.to_excel(file_path, index=False)
    print(f"\n✅ Done! Total Products: {len(all_data)}")
    print(f"📁 File saved at: {file_path}\n")

# 🚀 Run scraper
scrape_collection(base_collection_url)
