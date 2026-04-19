import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from itertools import count
from urllib.parse import urljoin

BASE_URL = "https://caracole.com"

COLLECTION_URLS = [
    "https://caracole.com/collections/chests-1",
    "https://caracole.com/collections/armoires",
    "https://caracole.com/collections/dressers",

]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
}

session = requests.Session()
session.headers.update(headers)

wb = Workbook()
ws = wb.active
ws.title = "Products"

# Header
ws.append(["Collection URL", "Product URL", "Image URL", "Product Name"])

for collection_url in COLLECTION_URLS:
    print(f"\n=== Scraping Collection: {collection_url} ===")

    for page in count(start=1):
        url = f"{collection_url}?page={page}"
        print(f"Scraping page {page}: {url}")

        r = session.get(url, timeout=30)
        if r.status_code != 200:
            print(f"Stopped: status {r.status_code}")
            break

        soup = BeautifulSoup(r.text, "html.parser")
        items = soup.select("div.grid-item.product-item")

        # Stop when no products found
        if not items:
            print("No items found, stopping pagination.")
            break

        for item in items:
            link_tag = item.select_one("a.product-link")
            if not link_tag:
                continue

            href = link_tag.get("href", "").strip()
            product_url = urljoin(BASE_URL, href)

            title_tag = item.select_one("p.product-item__title")
            product_name = title_tag.get_text(strip=True) if title_tag else ""

            img_tag = item.select_one("img")
            image_url = ""
            if img_tag:
                # Shopify/modern sites sometimes use data-src / data-srcset
                image_url = (
                    img_tag.get("src")
                    or img_tag.get("data-src")
                    or ""
                ).strip()

                if image_url.startswith("//"):
                    image_url = "https:" + image_url
                elif image_url.startswith("/"):
                    image_url = urljoin(BASE_URL, image_url)

            ws.append([collection_url, product_url, image_url, product_name])

file_name = "Caracole_Dressers_Chests.xlsx"
wb.save(file_name)

print(f"\nDone. Total products scraped: {ws.max_row - 1}")
print(f"Excel saved as: {file_name}")
