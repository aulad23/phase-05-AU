import requests
from bs4 import BeautifulSoup
import pandas as pd

# Target URL
url = "https://cameroncollection.com/upholstery/ottomans.html"

# Request page
headers = {
    "User-Agent": "Mozilla/5.0"
}
response = requests.get(url, headers=headers)
response.raise_for_status()

# Parse HTML
soup = BeautifulSoup(response.text, "html.parser")

data = []

# Find all products
products = soup.find_all("li", class_="grid-block")

for product in products:
    # Product URL
    a_tag = product.find("a", class_="product-image")
    product_url = a_tag["href"] if a_tag else ""

    # Image URL
    img_tag = product.find("img")
    image_url = img_tag["src"] if img_tag else ""

    # Product Name
    name_tag = product.find("h2", class_="product-name")
    product_name = name_tag.get_text(strip=True) if name_tag else ""

    data.append({
        "Product URL": product_url,
        "Image URL": image_url,
        "Product Name": product_name
    })

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_file = "Ottomans.xlsx"
df.to_excel(output_file, index=False)

print(f"Excel file saved as: {output_file}")
