import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin

BASE_URL = "https://studiobelvetro.com"
COLLECTION_URL = "https://studiobelvetro.com/collections/"

response = requests.get(COLLECTION_URL)
response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

rows = []

items = soup.find_all("div", class_="col c6")

for item in items:
    a_tag = item.find("a")
    img_tag = item.find("img")
    name_tag = item.find("figcaption")

    if not (a_tag and img_tag and name_tag):
        continue

    product_url = urljoin(BASE_URL, a_tag.get("href"))
    image_url = urljoin(BASE_URL, img_tag.get("src"))
    product_name = name_tag.get_text(strip=True)

    rows.append([
        product_url,
        image_url,
        product_name
    ])

# Create DataFrame with required order
df = pd.DataFrame(
    rows,
    columns=["Product URL", "Image URL", "Product Name"]
)

# Save to Excel
output_file = "studiobelvetro_collections.xlsx"
df.to_excel(output_file, index=False)

print("Excel file created:", output_file)
