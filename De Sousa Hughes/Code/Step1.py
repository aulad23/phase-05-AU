"""
De Sousa Hughes - Consoles Scraper
Input  : Multiple category URLs
Output : desousahughes_bedside_tables.xlsx
Columns: Product URL | Image URL | Product Name | SKU
"""

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


BASE_URL = "https://www.desousahughes.com"
TARGET_URLS = [
    "https://www.desousahughes.com/furniture/mirrors/",
    #"https://www.desousahughes.com/furniture/storage/sideboards/"
]
CATEGORY = "Mirrors"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )
}


def build_sku(vendor: str, category: str, index: int) -> str:
    """First 3 letters of vendor + first 2 letters of category + index."""
    v = vendor.replace(" ", "").upper()[:3]
    c = category.replace("-", "").upper()[:2]
    return f"{v}{c}{index}"


def scrape() -> list[dict]:
    products = []
    seen_hrefs = set()  # ✅ Global dedup across ALL URLs
    idx = 1

    for url in TARGET_URLS:  # ✅ Iterate over each URL
        print(f"   🌐 Fetching: {url}")
        resp = requests.get(url, headers=HEADERS, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        links = soup.select("a.el[href]")

        for link in links:
            href = link.get("href", "")
            if not href or href in seen_hrefs:  # ✅ Skip duplicates across pages
                continue
            seen_hrefs.add(href)

            # ── Product URL ──────────────────────────────────────────
            product_url = BASE_URL + href

            # ── Image URL ────────────────────────────────────────────
            img_tag = link.select_one("picture img")
            image_url = img_tag.get("src", "") if img_tag else ""

            # ── Product Name ─────────────────────────────────────────
            p_tags = link.find_all("p")
            if len(p_tags) >= 2:
                vendor       = p_tags[0].get_text(strip=True)
                product_raw  = p_tags[1].get_text(strip=True)
                product_name = f"{vendor} {product_raw}"
            elif len(p_tags) == 1:
                vendor       = p_tags[0].get_text(strip=True)
                product_name = vendor
            else:
                vendor       = "Unknown"
                product_name = "Unknown"

            # ── SKU ──────────────────────────────────────────────────
            sku = build_sku(vendor, CATEGORY, idx)

            products.append({
                "Product URL":  product_url,
                "Image URL":    image_url,
                "Product Name": product_name,
                "SKU":          sku,
            })
            idx += 1

    return products


def save_excel(products: list[dict], filename: str = "desousahughes_Mirrors.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bedside Tables"

    headers = ["Product URL", "Image URL", "Product Name", "SKU"]
    ws.append(headers)

    for product in products:
        ws.append([product.get(key, "") for key in headers])

    wb.save(filename)
    print(f"✅  Saved {len(products)} products → {filename}")


if __name__ == "__main__":
    print("🔍  Scraping:", TARGET_URLS)
    data = scrape()
    print(f"   Found {len(data)} products")
    save_excel(data)