import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook
import time

BASE_URL = "https://quatrine.com"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

def scrape_page(url):
    """Scrape a single page and return product data + next page URL."""
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"❌ Request failed: {e}")
        return [], None

    soup = BeautifulSoup(response.text, "html.parser")
    products = []

    articles = soup.find_all("article", class_="product-list-item")
    for article in articles:
        # Product URL & Name
        title_tag = article.find("h2", class_="product-list-item-title").find("a")
        product_url = urljoin(BASE_URL, title_tag["href"])
        product_name = title_tag.get_text(strip=True)

        # Image URL
        img_tag = article.find("figure").find("img")
        image_url = urljoin(BASE_URL, img_tag.get("src", "")) if img_tag else ""

        products.append([product_url, image_url, product_name])

    # Pagination
    next_btn = soup.select_one("li.next a")
    next_page = urljoin(BASE_URL, next_btn["href"]) if next_btn else None

    return products, next_page


def scrape_all_pages(start_url):
    """Scrape all pages starting from start_url."""
    all_products = []
    current_url = start_url

    while current_url:
        print(f"Scraping: {current_url}")
        data, next_page = scrape_page(current_url)
        all_products.extend(data)
        current_url = next_page
        time.sleep(1)  # be polite to the server

    return all_products


if __name__ == "__main__":
    start_urls = [
        "https://quatrine.com/collections/living-room-furniture/ottomans-benches",
        #"https://quatrine.com/collections/bedroom/Headboards"
    ]

    all_products = []
    for url in start_urls:
        all_products.extend(scrape_all_pages(url))

    # -------- Excel File --------
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    # Header
    ws.append(["Product URL", "Image URL", "Product Name"])

    # Data rows
    for product in all_products:
        ws.append(product)

    wb.save("quatrine_ottomans-benches.xlsx")

    print(f"\n✅ Excel file created: quatrine_Sofas_Loveseats.xlsx")
    print(f"✅ Total products scraped: {len(all_products)}")
