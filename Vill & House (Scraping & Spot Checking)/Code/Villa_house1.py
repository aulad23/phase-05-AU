import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from urllib.parse import urljoin

# ====== AUTO-DETECT SCRIPT FOLDER ======
script_folder = os.path.dirname(os.path.abspath(__file__))
output_file = os.path.join(script_folder, "Vila_desks.xlsx")

# ====== CATEGORY URL LIST ======
category_urls = [
    "https://vandh.com/desks-consoles/"
]

# ====== SCRAPER FUNCTION ======
def scrape_page(url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    print(f"   → GET {url}")
    response = requests.get(url, headers=headers, timeout=20)

    if response.status_code != 200:
        print(f"   ❌ Failed to access {url} (status {response.status_code})")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    products = soup.select("li.product")
    data = []

    if not products:
        print("   ⚠️ No products found on this page selector: li.product")

    for product in products:
        try:
            # Product URL (relative holeo absolute kore nebo)
            link_tag = product.select_one("figure.card-figure a")
            product_url = ""
            if link_tag and link_tag.get("href"):
                product_url = urljoin("https://vandh.com", link_tag["href"].strip())

            # Image URL (src na thakle data-src / data-srcset check kora jete pare)
            img_tag = product.select_one("div.card-img-container img")
            image_url = ""
            if img_tag:
                if img_tag.get("src"):
                    image_url = urljoin("https://vandh.com", img_tag["src"].strip())
                elif img_tag.get("data-src"):
                    image_url = urljoin("https://vandh.com", img_tag["data-src"].strip())

            # Product Name
            name_tag = product.select_one("h3.card-title a")
            product_name = name_tag.get_text(strip=True) if name_tag else ""

            # SKU
            sku_tag = product.select_one("div.card-text strong")
            sku = ""
            if sku_tag:
                sku_text = sku_tag.get_text(strip=True)
                sku = sku_text.replace("SKU :", "").replace("SKU:", "").strip()

            data.append({
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name,
                "SKU": sku
            })
        except Exception as e:
            print(f"   ⚠️ Error parsing product: {e}")
            continue

    return data

# ====== MAIN SCRIPT ======
all_data = []

print("\n🚀 Starting scrape from category URLs:")
for cu in category_urls:
    print("   -", cu)
print("=============================================\n")

for cat_idx, cat_url in enumerate(category_urls, start=1):
    print(f"\n==============================")
    print(f"📂 Category {cat_idx}: {cat_url}")
    print(f"==============================")

    page = 1
    while True:
        # first page = category URL, next pages = ?page=2,3 ...
        if page == 1:
            page_url = cat_url
        else:
            page_url = f"{cat_url}?page={page}"

        print(f"\n🔎 Scraping Page {page} → {page_url}")
        new_data = scrape_page(page_url)

        # jodi ektao product na pai, tahole oi category te ar page nai dhore break
        if not new_data:
            print(f"❌ No products found on page {page}. Stopping this category.\n")
            break

        all_data.extend(new_data)
        print(f"✅ Found {len(new_data)} products on page {page}. Total so far: {len(all_data)}\n")

        page += 1
        time.sleep(2)  # polite delay

# ====== SAVE TO EXCEL ======
if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(output_file, index=False)
    print("=============================================")
    print(f"🎯 Total Products Collected: {len(all_data)}")
    print(f"💾 File Saved To: {output_file}")
    print("=============================================")
else:
    print("\n⚠️ No products found at all.")
