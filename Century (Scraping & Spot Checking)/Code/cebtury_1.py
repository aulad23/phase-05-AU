import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

# ================== CONFIG ==================
BASE_URL = [
    "https://shop.centuryfurniture.com/search?options%5Bprefix%5D=last&page=1&q=Table+Lamps",
    ]
PARAMS = {
    "filter.p.product_type": "",
    "sort_by": "title-ascending"
}
TOTAL_PAGES = 3

DOWNLOAD_PATH = "century_Table_Lamps.xlsx"
# ============================================

all_products = []

for base_url in BASE_URL:  # ✅ প্রতিটি URL এর জন্য loop
    for page in range(1, TOTAL_PAGES + 1):
        print(f"\n🔎 Scraping page {page} from: {base_url[:60]}...")
        PARAMS["page"] = page
        response = requests.get(base_url, params=PARAMS)
        if response.status_code != 200:
            print(f"❌ Failed to fetch page {page}")
            continue

        soup = BeautifulSoup(response.text, "lxml")

        product_wrappers = soup.find_all("div", class_="card-wrapper")
        for wrapper in product_wrappers:
            a_tag = wrapper.find("a", {"id": lambda x: x and "CardLink" in x})
            if not a_tag:
                continue
            product_name = a_tag.get_text(strip=True)
            product_url = "https://shop.centuryfurniture.com" + a_tag.get("href")

            img_tag = wrapper.find("img")
            image_url = None
            if img_tag and img_tag.has_attr("srcset"):
                srcset_url = img_tag["srcset"].split(",")[0].strip().split(" ")[0]
                image_url = "https:" + srcset_url.split("&width")[0]
            elif img_tag and img_tag.has_attr("src"):
                image_url = "https:" + img_tag["src"].split("&width")[0]

            print(f"✅ {product_name} | {image_url}")

            all_products.append({
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name
            })

df = pd.DataFrame(all_products, columns=["Product URL", "Image URL", "Product Name"])
df.to_excel(DOWNLOAD_PATH, index=False)
print(f"\n✅ Scraping complete. File saved to: {DOWNLOAD_PATH}")