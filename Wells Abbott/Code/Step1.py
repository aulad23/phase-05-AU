"""
Wells Abbott - Multi-URL Furniture Product Scraper
===================================================
Multiple URL scrape kore ekta Excel e save korbe.

Requirements:
    pip install playwright beautifulsoup4 openpyxl
    playwright install chromium

Usage:
    python wells_abbott_scraper.py
"""

import sys
import time
from urllib.parse import urljoin

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("❌ Run: pip install playwright && playwright install chromium")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("❌ Run: pip install beautifulsoup4")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("❌ Run: pip install openpyxl")
    sys.exit(1)

# ─── URLs to scrape ───────────────────────────────────────────────────────────
URLS = [
    "https://www.wellsabbott.com/collections/furniture?usf_sort=-date&uff_gah0kc_metafield%3Acustom.style_furniture=Planters",
    #"https://www.wellsabbott.com/collections/furniture?usf_sort=-date&uff_gah0kc_metafield%3Acustom.style_furniture=Rocking%20Chairs",
    #"https://www.wellsabbott.com/collections/furniture?usf_sort=-date&uff_gah0kc_metafield%3Acustom.style_furniture=Occasional%20%26%20Slipper%20Chairs",
    #"https://www.wellsabbott.com/collections/furniture?usf_sort=-date&uff_gah0kc_metafield%3Acustom.style_furniture=Lounge%20Chairs"
    # Add more URLs here as needed:
    # "https://www.wellsabbott.com/collections/furniture?...",
]

OUTPUT_FILE = "wells_abbott_Baskets_Planters.xlsx"
BASE_URL    = "https://www.wellsabbott.com"

# ─────────────────────────────────────────────────────────────────────────────


def scrape_products(page, url: str) -> list[dict]:
    """Scrape one URL and return list of products."""
    print(f"\n🌐 Loading: {url}")
    page.goto(url, wait_until="networkidle", timeout=60_000)

    try:
        page.wait_for_selector("li.usf-sr-product", timeout=20_000)
        print("✅ Products found — parsing...")
    except Exception:
        print("⚠️  Timeout. Parsing whatever is available...")

    # Scroll to load lazy images
    for _ in range(5):
        prev = page.locator("li.usf-sr-product").count()
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1.5)
        if page.locator("li.usf-sr-product").count() == prev:
            break

    soup = BeautifulSoup(page.content(), "html.parser")
    cards = soup.select("li.usf-sr-product")
    print(f"📦 Found {len(cards)} product(s)")

    products = []
    for card in cards:
        name_tag = card.select_one("h3.card__heading a.full-unstyled-link")
        product_name = name_tag.get_text(strip=True) if name_tag else "N/A"

        href = name_tag["href"] if name_tag and name_tag.get("href") else None
        product_url = urljoin(BASE_URL, href) if href else "N/A"

        img_tag = card.select_one(".card__media img") or card.select_one("img")
        img_url = "N/A"
        if img_tag:
            src = img_tag.get("src", "")
            if src.startswith("//"):
                src = "https:" + src
            img_url = src

        products.append({
            "Product URL":  product_url,
            "Image URL":    img_url,
            "Product Name": product_name,
        })

    return products


def save_excel(all_products: list[dict], filepath: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Products"

    ws.append(["Product URL", "Image URL", "Product Name"])

    for p in all_products:
        ws.append([p["Product URL"], p["Image URL"], p["Product Name"]])

    wb.save(filepath)
    print(f"\n💾 Saved {len(all_products)} total rows → {filepath}")


def main():
    all_products = []
    seen_urls = set()  # Duplicate product URL বাদ দেওয়ার জন্য

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            )
        )

        for i, url in enumerate(URLS, 1):
            print(f"\n{'='*60}")
            print(f"🔗 URL {i}/{len(URLS)}")
            products = scrape_products(page, url)

            for p in products:
                if p["Product URL"] not in seen_urls:
                    seen_urls.add(p["Product URL"])
                    all_products.append(p)

        browser.close()

    if not all_products:
        print("⚠️  No products found.")
        return

    print(f"\n✅ Total unique products: {len(all_products)}")
    save_excel(all_products, OUTPUT_FILE)


if __name__ == "__main__":
    main()