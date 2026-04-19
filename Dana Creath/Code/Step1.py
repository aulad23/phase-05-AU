"""
Dana Creath Chandeliers Scraper
Install: pip install playwright beautifulsoup4 openpyxl && playwright install chromium
"""

from bs4 import BeautifulSoup
from openpyxl import Workbook
import re, json, time, random

BASE_URLS = [
    "https://danacreath.com/product-category/tables-accessories/mirrors/",
    #"https://danacreath.com/custom-gallery/",
    #"https://danacreath.com/product-category/chandeliers/linear/",
    #"https://danacreath.com/product-category/crystal-chandeliers/",
]

def get_page(url):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0")
        page.goto(url, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(3000)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(2000)
        html = page.content()
        browser.close()
    return BeautifulSoup(html, "html.parser")

def get_image(img):
    if not img: return ""
    srcset = img.get("srcset", "")
    if srcset:
        best, best_w = "", 0
        for part in srcset.split(","):
            t = part.strip().split()
            if len(t) >= 2:
                try:
                    w = int(t[1].replace("w", ""))
                    if w > best_w: best_w, best = w, t[0]
                except: pass
        if best: return best
    for attr in ["data-src", "data-lazy-src", "src"]:
        if img.get(attr, "").startswith("http"): return img[attr]
    return ""

def split_name(title):
    for sep in ["\u2013", " - ", "|"]:
        if sep in title: return title.split(sep)[0].strip(), title
    return title, title

def parse_products(soup):
    products = []
    items = []
    for sel in ["li.product", "li.type-product", "div.product", ".products li", ".product-item"]:
        items = soup.select(sel)
        if items: break

    if items:
        for item in items:
            link = item.select_one("a[href*='/product/']") or item.select_one("a")
            url = link["href"] if link else ""
            img = get_image(item.select_one("img"))
            title = ""
            for ts in ["h2", "h3", ".product-title", ".woocommerce-loop-product__title"]:
                t = item.select_one(ts)
                if t and t.get_text(strip=True): title = t.get_text(strip=True); break
            name, sku = split_name(title)
            if name or url:
                products.append([url, img, name, sku])
    else:
        seen = set()
        for a in soup.find_all("a", href=re.compile(r"/product/[^/]")):
            href = a["href"]
            if href not in seen and "product-category" not in href:
                seen.add(href)
                name, sku = split_name(a.get_text(strip=True))
                products.append([href, get_image(a.find("img")), name, sku])

    if not products:
        for s in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(s.string)
                for item in (data if isinstance(data, list) else data.get("itemListElement", [data])):
                    obj = item.get("item", item)
                    if obj.get("@type") in ("Product", "ListItem"):
                        name, sku = split_name(obj.get("name", ""))
                        img = obj.get("image", "")
                        products.append([obj.get("url",""), img[0] if isinstance(img,list) else img, name, sku])
            except: pass
    return products

def get_total_pages(soup):
    pages = set()
    for a in soup.select("a.page-numbers, .pagination a, nav.woocommerce-pagination a"):
        try: pages.add(int(a.get_text(strip=True)))
        except:
            m = re.search(r"/page/(\d+)", a.get("href",""))
            if m: pages.add(int(m.group(1)))
    return max(pages) if pages else 1

def scrape_category(base_url):
    print(f"\n{'='*60}")
    print(f"Category: {base_url}")
    print(f"{'='*60}")

    print("Fetching page 1...")
    soup = get_page(base_url)
    total = get_total_pages(soup)
    print(f"Total pages: {total}")
    all_prods = parse_products(soup)
    print(f"  -> {len(all_prods)} products")

    for pg in range(2, total + 1):
        w = random.uniform(2, 4)
        print(f"Fetching page {pg}... (wait {w:.1f}s)")
        time.sleep(w)
        all_prods.extend(parse_products(get_page(f"{base_url}page/{pg}/")))
        print(f"  -> page {pg} done")

    return all_prods

def scrape_all():
    all_prods = []
    for url in BASE_URLS:
        all_prods.extend(scrape_category(url))
        time.sleep(random.uniform(2, 4))

    seen, unique = set(), []
    for p in all_prods:
        key = p[0] or p[2]
        if key and key not in seen: seen.add(key); unique.append(p)
    print(f"\nTotal unique products: {len(unique)}")
    return unique

def save_to_excel(products, filename="danacreath_mirrors.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Chandeliers"
    ws.append(["Product URL", "Image URL", "Product Name", "SKU"])
    for p in products:
        ws.append(p)
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 40
    wb.save(filename)
    print(f"\nSaved {len(products)} products -> '{filename}'")

if __name__ == "__main__":
    products = scrape_all()
    if products:
        save_to_excel(products)
    else:
        print("\n0 products found! Site may need CAPTCHA solving.")