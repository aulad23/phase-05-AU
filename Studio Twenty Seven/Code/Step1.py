import requests
from bs4 import BeautifulSoup
import openpyxl
import re
import time

# ============================================================
#  শুধু এই URL টা পরিবর্তন করুন — বাকি সব automatic
# ============================================================
COLLECTION_URL = [
    "https://shop.studiotwentyseven.com/collections/sculptures",
    #"https://shop.studiotwentyseven.com/collections/sofas"
]
# ============================================================

BASE_URL = "https://shop.studiotwentyseven.com"
VENDOR = "studiotwentyseven"

def get_page(url, page=1):
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, params={"page": page}, headers=headers, timeout=15)
    r.raise_for_status()
    return r.text

def clean_image_url(url):
    return re.sub(r"_\d+x(\.\w+)(\?|$)", r"\1\2", url).split("?")[0]

def clean_price(text):
    return re.sub(r"[^\d.]", "", text.replace(" ", ""))

def get_category(url):
    parts = url.rstrip("/").split("/")
    for i, p in enumerate(parts):
        if p == "collections" and i + 1 < len(parts):
            return parts[i + 1]
    return "unknown"

def make_sku(category, index):
    v = VENDOR.upper()[:3]
    c = re.sub(r"[^a-zA-Z]", "", category).upper()[:2]
    return f"{v}{c}{index}"

def scrape_collection(collection_url):
    category = get_category(collection_url)
    output_file = f"{category}.xlsx"
    page = 1
    index = 1
    seen_urls = set()  # ডুপ্লিকেট এড়াতে

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Product URL", "Image URL", "Product Name", "SKU", "List Price (USD)"])

    while True:
        print(f"[{category}] Page {page} scraping...")
        soup = BeautifulSoup(get_page(collection_url, page), "html.parser")
        items = soup.select("div.boost-pfs-filter-product-item, div.block.grid-item")
        if not items:
            break

        for item in items:
            link = item.select_one("div.block-image a")
            product_url = BASE_URL + link["href"] if link else ""

            # একই URL আগে দেখা গেলে skip করো
            if product_url in seen_urls:
                continue
            seen_urls.add(product_url)

            img = item.select_one("img.boost-pfs-filter-product-item-main-image")
            raw = ""
            if img:
                srcset = img.get("data-srcset") or img.get("srcset") or ""
                raw = srcset.strip().split(",")[0].strip().split(" ")[0] if srcset else img.get("src", "")
            image_url = clean_image_url(raw) if raw else ""

            title_tag = item.select_one("h2.block-title")
            vendor_tag = item.select_one("span.block-vendor b")
            vendor = vendor_tag.get_text(strip=True) if vendor_tag else ""
            if title_tag:
                span = title_tag.find("span", class_="block-vendor")
                if span:
                    span.decompose()
                name = title_tag.get_text(strip=True)
            else:
                name = ""

            price_tag = item.select_one("span.boost-pfs-filter-product-item-regular-price")
            price = clean_price(price_tag.get_text()) if price_tag else ""

            ws.append([
                product_url,
                image_url,
                f"{name} {vendor}".strip(),
                make_sku(category, index),
                price,
            ])
            index += 1

        if not soup.select_one("a.boost-pfs-filter-next-btn, li.next a, a[rel='next']"):
            break
        page += 1
        time.sleep(0.5)

    wb.save(output_file)
    print(f"✅ Done! {index-1} products → {output_file}")

# ─── Main: সব URL loop করো ───────────────────────────────
urls = COLLECTION_URL if isinstance(COLLECTION_URL, list) else [COLLECTION_URL]
for url in urls:
    scrape_collection(url)