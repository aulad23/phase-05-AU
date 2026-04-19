import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://palmerhargrave.com/shop/?post_id=138&form_id=cee1aae&queried_type=WP_Post&queried_id=138&categories[]=14"

headers = {
    "User-Agent": "Mozilla/5.0"
}

response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "lxml")

products = []

articles = soup.find_all("article", class_="e-add-post")

for article in articles:
    # Product URL
    product_link = article.find("a", class_="e-add-post-image")
    product_url = product_link["href"] if product_link else ""

    # Image URL
    img = article.find("img")
    image_url = img["src"] if img else ""

    # Product Name
    title = article.select_one("h3.e-add-post-title a")
    product_name = title.get_text(strip=True) if title else ""

    # SKU
    sku = article.select_one("div.e-add-item_custommeta span")
    sku = sku.get_text(strip=True) if sku else ""

    products.append({
        "Product URL": product_url,
        "Image URL": image_url,
        "Product Name": product_name,
        "SKU": sku
    })

# Create DataFrame with required column order
df = pd.DataFrame(products, columns=[
    "Product URL",
    "Image URL",
    "Product Name",
    "SKU"
])

# Save to Excel
df.to_excel("palmerhargrave_products.xlsx", index=False)

print("✅ Excel file created with correct column order")
