import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urlparse

# Files
input_file = "Ottomans.xlsx"
output_file = "Ottomans_final.xlsx"

headers = {
    "User-Agent": "Mozilla/5.0"
}

# Load input Excel
df = pd.read_excel(input_file)

vendor_code = "CAM"
output_data = []

def get_category_code(product_url):
    """
    Extract category from URL and return first two letters as category code
    """
    path = urlparse(product_url).path
    parts = path.split("/")

    # Find category slug (usually before product slug)
    for part in parts:
        if "-" in part and "html" not in part:
            letters = "".join([c for c in part if c.isalpha()])
            return letters[:2].upper()

    return "XX"  # fallback

for index, row in df.iterrows():
    product_url = row["Product URL"]
    image_url = row["Image URL"]
    product_name = row["Product Name"]

    category_code = get_category_code(product_url)
    sku = f"{vendor_code}{category_code}{str(index + 1).zfill(2)}"

    description = ""

    try:
        response = requests.get(product_url, headers=headers, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        desc_div = soup.find("div", class_="product-description")
        if desc_div:
            description = " ".join(desc_div.stripped_strings)

    except Exception:
        description = ""

    output_data.append({
        "Product URL": product_url,
        "Image URL": image_url,
        "Product Name": product_name,
        "SKU": sku,
        "Product Family Id": product_name,
        "Description": description,
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Height": ""
    })

# Save final Excel
output_df = pd.DataFrame(output_data)
output_df.to_excel(output_file, index=False)

print(f"Final Excel saved as: {output_file}")
