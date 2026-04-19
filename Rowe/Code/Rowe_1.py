import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import os

# ========== SETTINGS ==========
BASE_URL = "https://rowefurniture.com"

# Original category URLs (hash thakleo problem nai, niche handle kore nibo)
CATEGORY_URLS = [
    "https://rowefurniture.com/custom-sectional",
    #"https://rowefurniture.com/sleepers"
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def build_page_url(base_url: str, page: int) -> str:
    """
    base_url theke specific page er URL banabe.
    Jodi base_url er moddhe age theke ? thake tahole &pagenumber= jog hobe,
    na thakle ?pagenumber= jog hobe.
    """
    if "?" in base_url:
        return f"{base_url}&pagenumber={page}"
    else:
        return f"{base_url}?pagenumber={page}"


def scrape_category(category_url):
    """Scrape all product data from a single category URL."""
    products = []
    page = 1

    # 🔹 Hash (#) er porer part remove kore nilam – requests e kaj kore na
    base_url = category_url.split("#", 1)[0].strip()

    while True:
        url = build_page_url(base_url, page)
        print(f"Scraping page {page} → {url}")

        try:
            response = requests.get(url, timeout=15)
        except Exception as e:
            print(f"Request error on page {page}: {e}")
            break

        if response.status_code != 200:
            print(f"Request failed or no more pages. Status: {response.status_code}")
            break

        soup = BeautifulSoup(response.text, "html.parser")
        items = soup.find_all("div", class_="product-item")

        if not items:
            print("No more products found on this page. Stopping.")
            break

        for item in items:
            # Product URL
            picture_div = item.find("div", class_="picture")
            a_tag = picture_div.find("a", href=True) if picture_div else None
            product_url = urljoin(BASE_URL, a_tag["href"]) if a_tag else None

            # Image URL
            img_tag = a_tag.find("img", class_="picture-img") if a_tag else None
            image_url = img_tag.get("src") if img_tag else None

            # Product Name
            title_h2 = item.find("h2", class_="product-title")
            name_tag = title_h2.find("a") if title_h2 else None
            product_name = name_tag.get_text(strip=True) if name_tag else None

            # SKU
            sku_tag = item.find("div", class_="sku")
            sku = sku_tag.get_text(strip=True) if sku_tag else None

            products.append({
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name,
                "SKU": sku
            })

        page += 1

    return products


def save_to_excel(data):
    """Save scraped data to Excel in the same folder as this script."""
    if not data:
        print("No data to save.")
        return

    df = pd.DataFrame(data)
    file_name = "rowefurniture_sectional.xlsx"
    file_path = os.path.join(SCRIPT_DIR, file_name)

    df.to_excel(file_path, index=False)
    print(f"✅ Data saved: {file_path}")
    print(f"Total products: {len(df)}")


if __name__ == "__main__":
    all_products = []

    for url in CATEGORY_URLS:
        print(f"\n=== Scraping category: {url} ===")
        category_products = scrape_category(url)
        print(f"Found {len(category_products)} products in this category.")
        all_products.extend(category_products)

    save_to_excel(all_products)
