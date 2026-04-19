from playwright.sync_api import sync_playwright
import pandas as pd
from urllib.parse import urljoin
import time

# ================ CONFIG ===============
BASE_URL = "https://uttermost.com"
URL_TEMPLATE = "https://uttermost.com/furniture?page={page}&brand%5Bfilter%5D=Uttermost%2C72&shop_by_type%5Bfilter%5D=Tables%2C561&shop_by_type_of_product%5Bfilter%5D=Side%2FEnd%2FLamp%2C966"
OUTPUT_FILE = "uttermost_Side_Tables.xlsx"
HEADLESS = False


# ========================================

def progressive_scroll(page, scroll_pause=1.5):
    """
    Scroll করে ধীরে ধীরে নিচে যায় যাতে সব lazy-loaded content লোড হয়
    """
    print("Progressive scrolling to load all products...")

    last_height = page.evaluate("document.body.scrollHeight")
    scroll_step = 500  # প্রতিবার 500px scroll করবে
    current_position = 0

    while True:
        # একটু একটু করে scroll করা
        current_position += scroll_step
        page.evaluate(f"window.scrollTo(0, {current_position})")
        time.sleep(scroll_pause)  # লোড হওয়ার জন্য অপেক্ষা

        # নতুন height চেক করা
        new_height = page.evaluate("document.body.scrollHeight")

        # যদি page এর শেষে পৌঁছে যায়
        if current_position >= new_height:
            # একটু বেশি অপেক্ষা করা শেষ content লোডের জন্য
            time.sleep(2)

            # আবার চেক করা নতুন content আসছে কিনা
            final_height = page.evaluate("document.body.scrollHeight")
            if final_height == new_height:
                print(f"✓ Scrolling complete. Final height: {final_height}px")
                break
            else:
                last_height = final_height
                continue

        last_height = new_height

    # শেষে একবার উপরে scroll করা
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(1)


all_data = []
seen_skus = set()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=HEADLESS)
    context = browser.new_context(
        viewport={"width": 1400, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    )
    page = context.new_page()

    # Block unnecessary resources
    page.route("**/*", lambda route: route.abort()
    if route.request.resource_type in ["stylesheet", "font", "media"]
    else route.continue_())

    # ----------------- Determine total pages -----------------
    first_page_url = URL_TEMPLATE.format(page=1)
    print("Opening first page:", first_page_url)

    page.goto(first_page_url, wait_until="domcontentloaded", timeout=60000)

    # Wait for products to appear
    try:
        page.wait_for_selector("div.item-root-Chs", state="visible", timeout=30000)
        print("✓ Products loaded")
        time.sleep(2)
    except:
        print("⚠ Products not loading, waiting longer...")
        time.sleep(5)

    # Extract total pages
    try:
        pagination_selector = "div.css-1m76rdz-singleValue"
        page.wait_for_selector(pagination_selector, timeout=10000)
        total_pages_text = page.locator(pagination_selector).text_content()
        total_pages = int(total_pages_text.split("of")[-1].strip())
    except Exception as e:
        print(f"Could not read total pages, defaulting to 1: {e}")
        total_pages = 1

    print(f"Total pages to scrape: {total_pages}")

    # ----------------- Loop through pages -----------------
    for page_number in range(1, total_pages + 1):
        page_url = URL_TEMPLATE.format(page=page_number)
        print(f"\n{'=' * 50}")
        print(f"📄 Scraping Page {page_number}/{total_pages}")
        print(f"{'=' * 50}")

        page.goto(page_url, wait_until="domcontentloaded", timeout=60000)

        # Wait for products to load
        try:
            page.wait_for_selector("div.item-root-Chs", timeout=15000)
            print("✓ Initial products loaded")
            time.sleep(2)
        except:
            print(f"⚠ No products found on page {page_number}")
            continue

        # Progressive scroll to load all lazy images and products
        progressive_scroll(page, scroll_pause=1.5)

        # একটু অতিরিক্ত সময় দেওয়া final rendering এর জন্য
        time.sleep(2)

        # ===== Collect products =====
        cards = page.locator("div.item-root-Chs").all()
        print(f"Found {len(cards)} products on page {page_number}")

        scraped_count = 0
        for i, card in enumerate(cards):
            try:
                # Product Name
                name_locator = card.locator("a.item-name-LPg span")
                product_name = name_locator.first.text_content().strip() if name_locator.count() > 0 else "N/A"

                # SKU
                sku_locator = card.locator("p span.font-semibold")
                sku = sku_locator.first.text_content().strip() if sku_locator.count() > 0 else "N/A"

                # Product URL
                url_locator = card.locator("a.item-images--uD")
                product_url = url_locator.get_attribute("href") if url_locator.count() > 0 else ""
                product_url = urljoin(BASE_URL, product_url) if product_url else "N/A"

                # Image URL
                image_url = "N/A"
                try:
                    img_locator = card.locator('img[class*="rounded-"]')
                    if img_locator.count() > 0:
                        img = img_locator.first
                        image_src = img.get_attribute("data-src") or img.get_attribute("src") or ""
                        if image_src:
                            image_url = urljoin(BASE_URL, image_src) if image_src.startswith("/") else image_src
                except Exception as img_error:
                    print(f"  ⚠ Image error for product {i + 1}: {img_error}")

                # Skip incomplete products
                if product_name == "N/A" or sku == "N/A" or product_url == "N/A":
                    print(f"  ⊘ Skipping incomplete product {i + 1}")
                    continue

                # Skip duplicates
                if sku in seen_skus:
                    print(f"  ⊘ Duplicate SKU: {sku}")
                    continue

                seen_skus.add(sku)
                scraped_count += 1

                all_data.append({
                    "Product URL": product_url,
                    "Image URL": image_url,
                    "Product Name": product_name,
                    "SKU": sku
                })

                # Progress indicator
                if scraped_count % 10 == 0:
                    print(f"  ✓ Scraped {scraped_count} products so far...")

            except Exception as e:
                print(f"  ✗ Error scraping product {i + 1}: {e}")
                continue

        print(f"✓ Successfully scraped {scraped_count} unique products from page {page_number}")
        print(f"📊 Total products collected so far: {len(all_data)}")

    browser.close()

# ----------------- Save Excel -----------------
print(f"\n{'=' * 50}")
print("Saving data to Excel...")
print(f"{'=' * 50}")

if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(OUTPUT_FILE, index=False)
    print("\n✅ Scraping complete!")
    print(f"📦 Total unique products scraped: {len(df)}")
    print(f"💾 Excel saved as: {OUTPUT_FILE}")
    print(f"\nColumns: {', '.join(df.columns.tolist())}")
else:
    print("⚠️ No data collected.")