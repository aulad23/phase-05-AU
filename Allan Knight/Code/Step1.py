from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time

# ============================================================
# CATEGORY URL INPUT - Edit this URL as needed
# ============================================================
CATEGORY_URL = "https://allan-knight.com/allan-knight-collections#type=trays"
OUTPUT_FILE  = "AllanKnight_Trays.xlsx"
# ============================================================


def init_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    driver = webdriver.Chrome(options=options)
    return driver


def get_product_cards(driver, category_url):
    """
    Load category page, wait for JS filter, then collect:
      - Product URL  (from visible anchor)
      - Image URL    (from img.card-img-top inside the same card)
    directly from the listing page — no need to visit product pages for images.
    """
    print(f"\n{'='*60}")
    print(f"Opening: {category_url}")
    print(f"{'='*60}\n")

    driver.get(category_url)

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
    except TimeoutException:
        print("Page load timeout.")
        return []

    print("Waiting for JS filter to apply...")
    time.sleep(5)

    # Scroll to load all lazy-loaded images
    print("Scrolling to load all products...")
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
    print("✓ Scroll complete\n")

    # ----------------------------------------------------------------
    # Collect product cards: each card has one <a> link + one <img>
    # We look for the PARENT container of img.card-img-top
    # and check if it (or its ancestor) is visible.
    # ----------------------------------------------------------------
    print("Collecting visible product cards (URL + Image)...")

    base_domain = "https://www.allanknight.com"

    EXCLUDE_KEYWORDS = [
        "allan-knight-collections",
        "mailto:", "tel:",
        "/about", "/contact", "/blog",
        "/lookbook", "/tag-sales",
        ".jpg", ".png", ".pdf", ".css", ".js",
    ]

    # Find all img.card-img-top elements
    card_imgs = driver.find_elements(By.CSS_SELECTOR, "img.card-img-top")

    product_cards = []
    seen_urls = set()

    for img in card_imgs:
        try:
            # Skip hidden images (their parent card is display:none)
            if not img.is_displayed():
                continue

            # Get image src
            image_url = img.get_attribute("src") or ""

            # Also check data-src for lazy loaded images
            if not image_url or "placeholder" in image_url.lower():
                image_url = img.get_attribute("data-src") or image_url

            # Find the closest ancestor <a> tag with a product href
            # Walk up the DOM using JavaScript to find the link
            product_url = driver.execute_script("""
                var el = arguments[0];
                while (el && el.tagName !== 'A') {
                    el = el.parentElement;
                }
                return el ? el.href : null;
            """, img)

            if not product_url:
                # Try sibling/descendant <a> within the card container
                try:
                    card_container = driver.execute_script(
                        "return arguments[0].closest('.col, .card, .product, article, li, div[class*=\"col\"]');",
                        img
                    )
                    if card_container:
                        link_el = card_container.find_element(By.TAG_NAME, "a")
                        product_url = link_el.get_attribute("href") or ""
                except Exception:
                    pass

            if not product_url:
                continue

            product_url = product_url.strip().split("?")[0].split("#")[0]

            if base_domain not in product_url:
                continue
            if any(kw in product_url for kw in EXCLUDE_KEYWORDS):
                continue

            path = product_url.replace(base_domain, "").strip("/")
            parts = [p for p in path.split("/") if p]
            if len(parts) < 2:
                continue

            if product_url not in seen_urls:
                seen_urls.add(product_url)
                product_cards.append({
                    "Product URL": product_url,
                    "Image URL":   image_url,
                })

        except Exception:
            continue

    print(f"✓ Found {len(product_cards)} visible product cards\n")

    if not product_cards:
        print("No cards found — saving debug HTML...")
        with open("debug_page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("✓ Saved debug_page_source.html")
    else:
        print("Sample cards:")
        for card in product_cards[:3]:
            print(f"  URL  : {card['Product URL']}")
            print(f"  Image: {card['Image URL']}")
            print()

    return product_cards


def get_product_name(driver, product_url):
    """
    Visit individual product page to get the Product Name from <h1>.
    """
    try:
        print(f"  Fetching: {product_url}")
        driver.get(product_url)

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "h1"))
            )
        except TimeoutException:
            pass

        time.sleep(1)

        product_name = ""
        try:
            product_name = driver.find_element(By.TAG_NAME, "h1").text.strip()
        except NoSuchElementException:
            pass

        # Fallback: page title
        if not product_name:
            title = driver.title.strip()
            product_name = title.split("|")[0].strip() if "|" in title else title

        print(f"  OK: {product_name or '(not found)'}")
        return product_name

    except Exception as e:
        print(f"  Error: {e}")
        return ""


def save_to_excel(data, output_file):
    if not data:
        print("\nNo data to save!")
        return
    df = pd.DataFrame(data, columns=["Product Name", "Product URL", "Image URL"])
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n{'='*60}")
    print(f"Saved to '{output_file}'")
    print(f"Total products : {len(data)}")
    print(f"With names     : {df['Product Name'].astype(bool).sum()}")
    print(f"With images    : {df['Image URL'].astype(bool).sum()}")
    print(f"{'='*60}\n")


def main():
    print("\n" + "=" * 60)
    print(" " * 15 + "ALLAN KNIGHT SCRAPER - STEP 1")
    print("=" * 60 + "\n")
    print(f"Target URL : {CATEGORY_URL}")
    print(f"Output File: {OUTPUT_FILE}\n")

    driver = init_driver()

    try:
        # Step 1: Get product URL + Image from category listing page
        product_cards = get_product_cards(driver, CATEGORY_URL)

        if not product_cards:
            print("No products found. Exiting.")
            return

        # Step 2: Visit each product page to get Product Name only
        print(f"{'='*60}")
        print(f"Getting product names from {len(product_cards)} pages...")
        print(f"{'='*60}\n")

        results = []
        for idx, card in enumerate(product_cards, 1):
            print(f"[{idx}/{len(product_cards)}]")
            name = get_product_name(driver, card["Product URL"])
            results.append({
                "Product Name": name,
                "Product URL":  card["Product URL"],
                "Image URL":    card["Image URL"],
            })
            time.sleep(1.5)

        # Step 3: Save
        save_to_excel(results, OUTPUT_FILE)
        print("Scraping completed successfully!")

    finally:
        driver.quit()
        print("Browser closed.")


if __name__ == "__main__":
    main()