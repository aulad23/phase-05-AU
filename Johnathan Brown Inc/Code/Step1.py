import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://jonathanbrowninginc.com"
TARGET_URL = f"{BASE_URL}/products/flush-mounts"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

print("Fetching page...")
response = requests.get(TARGET_URL, headers=headers)
response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

items = soup.select("li.grid-item")
print(f"Found {len(items)} products")

data = []
for item in items:
    product_name = item.get("data-product-name", "").strip()
    deeplink = item.get("data-deeplink", "").strip()
    resting_img = item.get("data-resting", "").strip()
    category = item.get("data-category", "").strip()

    if not product_name:
        continue

    product_url = f"{BASE_URL}/products/{category}?deep={deeplink}" if deeplink and category else TARGET_URL
    image_url = f"{BASE_URL}{resting_img}" if resting_img.startswith("/") else resting_img

    data.append({
        "Product URL": product_url,
        "Image URL": image_url,
        "Product Name": product_name,
    })

df = pd.DataFrame(data, columns=["Product URL", "Image URL", "Product Name"])

output_file = "jonathan_browning_flush-mounts.xlsx"
df.to_excel(output_file, index=False)

print(f"\nDone! {len(df)} products saved to '{output_file}'")
print(df.head())