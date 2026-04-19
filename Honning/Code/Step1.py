import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://honning.us"
PAGE_URLS = [
    "https://honning.us/mirrors",
    #"https://honning.us/chaises-settees"
    # Add more URLs here as needed
]

VENDOR = "honning"

# ── Navigation / non-product paths to exclude (for fallback only) ──
NAV_PATHS = {
    "/", "/cart", "/new", "/view-all", "/seating", "/tables",
    "/storage-chests", "/desks", "/beds", "/mirrors", "/custom",
    "/finishes-2", "/in-situ", "/lookbook", "/profile-honning",
    "/clients-projects", "/process", "/social", "/press",
    "/contact", "/register-login", "/samplesale", "/collection",
    "/collection-folder", "/gallery-1", "/about",
}


def make_sku(vendor: str, category_slug: str, index: int) -> str:
    vendor_part   = vendor.replace("-", "").upper()[:3]
    category_part = category_slug.replace("-", "").upper()[:2]
    return f"{vendor_part}{category_part}{index}"


def scrape(page_url: str, index_offset: int = 0) -> pd.DataFrame:
    category_slug = page_url.rstrip("/").split("/")[-1]

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
        )
    }

    resp = requests.get(page_url, headers=headers, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    records = []

    # ══════════════════════════════════════════════════════════════
    # METHOD 1: Squarespace Gallery Grid (the REAL structure)
    #
    #   <figure class="gallery-grid-item">
    #     <div>
    #       <a class="gallery-grid-image-link" href="/seating/...">
    #         <img data-src="..." src="...">
    #       </a>
    #     </div>
    #     <figcaption>
    #       <p class="gallery-caption-content">Product Name</p>
    #     </figcaption>
    #   </figure>
    # ══════════════════════════════════════════════════════════════
    items = soup.select("figure.gallery-grid-item")

    if items:
        print(f"  ✅ Found {len(items)} items via figure.gallery-grid-item")
        for idx, fig in enumerate(items, start=1):
            # --- Product URL ---
            a_tag = fig.select_one("a.gallery-grid-image-link")
            href = a_tag["href"] if a_tag and a_tag.get("href") else ""
            product_url = BASE_URL + href if href.startswith("/") else href

            # --- Image URL ---
            img_tag = fig.select_one("img")
            image_url = ""
            if img_tag:
                image_url = (
                    img_tag.get("data-src")
                    or img_tag.get("data-image")
                    or img_tag.get("src", "")
                )

            # --- Product Name ---
            caption = fig.select_one("p.gallery-caption-content")
            product_name = caption.get_text(strip=True) if caption else ""

            # Fallback: try img alt text or URL slug
            if not product_name and img_tag:
                product_name = img_tag.get("alt", "")
            if not product_name and href:
                product_name = href.rstrip("/").split("/")[-1].replace("-", " ").title()

            global_idx = index_offset + idx
            sku = make_sku(VENDOR, category_slug, global_idx)

            records.append({
                "Product URL":  product_url,
                "Image URL":    image_url,
                "Product Name": product_name,
                "SKU":          sku,
            })
            print(f"  [{global_idx:>3}] {sku} | {product_name}")

    # ══════════════════════════════════════════════════════════════
    # METHOD 2: Fallback — link-based detection
    # (for pages where gallery-grid-item is not present)
    # ══════════════════════════════════════════════════════════════
    else:
        print(f"  ⚠ gallery-grid-item not found — using link-based fallback...")
        all_links = soup.find_all("a", href=True)
        seen_hrefs = set()

        for a in all_links:
            href = a["href"]
            parts = [p for p in href.strip("/").split("/") if p]
            if len(parts) < 2:
                continue
            if href.rstrip("/") in NAV_PATHS:
                continue
            if not a.find("img"):
                continue
            if href in seen_hrefs:
                continue
            seen_hrefs.add(href)

            product_url = BASE_URL + href if href.startswith("/") else href

            img_tag = a.select_one("img")
            image_url = ""
            if img_tag:
                image_url = (
                    img_tag.get("data-src")
                    or img_tag.get("data-image")
                    or img_tag.get("src", "")
                )

            # Try name from: h3 inside link → img alt → URL slug
            product_name = ""
            for sel in ["h3", "h2", "p"]:
                el = a.select_one(sel)
                if el and el.get_text(strip=True):
                    product_name = el.get_text(strip=True)
                    break
            if not product_name and img_tag:
                product_name = img_tag.get("alt", "")
            if not product_name:
                product_name = href.rstrip("/").split("/")[-1].replace("-", " ").title()

            idx = len(records) + 1
            global_idx = index_offset + idx
            sku = make_sku(VENDOR, category_slug, global_idx)

            records.append({
                "Product URL":  product_url,
                "Image URL":    image_url,
                "Product Name": product_name,
                "SKU":          sku,
            })
            print(f"  [{global_idx:>3}] {sku} | {product_name}")

    return pd.DataFrame(records)


if __name__ == "__main__":
    all_frames   = []
    index_offset = 0

    for url in PAGE_URLS:
        print(f"\n🔍 Scraping: {url}")
        df = scrape(url, index_offset=index_offset)
        all_frames.append(df)
        index_offset += len(df)

    combined = pd.concat(all_frames, ignore_index=True)

    output_file = "Mirrors.xlsx"
    combined.to_excel(output_file, index=False)

    print(f"\n✅ {len(combined)} টি product পাওয়া গেছে।")
    print(f"📄 Data সেভ হয়েছে: {output_file}")
    print(combined.to_string(index=False))