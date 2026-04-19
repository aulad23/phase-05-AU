import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from urllib.parse import urlparse

urls = [
    "https://www.mcleanlighting.com/product-category/lighting/lighting-on-hand/",
    "https://www.mcleanlighting.com/product-category/lighting/antique/"
]

headers = {
    "User-Agent": "Mozilla/5.0"
}

VENDOR_NAME = "McLean Lighting"
VENDOR_CODE = VENDOR_NAME.replace(" ", "")[:3].upper()  # MCL

all_products = []
product_index = 1

def get_category_code(category_url):
    slug = category_url.rstrip("/").split("/")[-1]   # lighting-on-hand / antique
    clean_slug = slug.replace("-", "")
    return clean_slug[:2].upper()

for url in urls:
    print(f"Scraping: {url}")
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "lxml")

    category_code = get_category_code(url)

    for li in soup.find_all("li", class_="product"):
        a_tag = li.find("a", href=True)
        img_tag = li.find("img")
        h3_tag = li.find("h3")

        product_url = a_tag["href"] if a_tag else ""
        image_url = img_tag["src"] if img_tag else ""
        product_name = h3_tag.get_text(strip=True) if h3_tag else ""

        sku = f"{VENDOR_CODE}{category_code}{product_index}"

        all_products.append({
            "Product URL": product_url,
            "Image URL": image_url,
            "Product Name": product_name,
            "SKU": sku
        })

        product_index += 1

    time.sleep(1)

# Column order enforce
df = pd.DataFrame(all_products, columns=[
    "Product URL",
    "Image URL",
    "Product Name",
    "SKU"
])

# Save to Excel
df.to_excel("mclean_lighting.xlsx", index=False)

print("✅ Scraping complete with SKU. File saved: mclean_lighting.xlsx")
