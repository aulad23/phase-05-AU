"""
Verellen Scraper - Button Pagination Version
Requirements:
    pip install selenium openpyxl webdriver-manager
Run:
    python verellen_scraper.py
"""

import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook


# ─── CONFIG ───────────────────────────────────────────────────────────
BASE_URL = "https://verellen.biz"
TARGET_URLS = [
    "https://verellen.biz/collections/product-type/outdoor",
   #"https://verellen.biz/collections/product-type/banquettes",floor-lamps
    #"https://verellen.biz/collections/product-type/chairs/wing-chairs",
    #"https://verellen.biz/collections/product-type/chairs/swivel-chairs",
    #"https://verellen.biz/collections/product-type/chairs/armless-chairs",
    #"https://verellen.biz/collections/product-type/chairs/lounges",
    #"https://verellen.biz/collections/product-type/chaises"
]
MANUFACTURER = "Verellen"
OUTPUT_FILE = "verellen_Outdoor_Seating.xlsx"


def format_product_name(raw_name: str) -> str:
    return raw_name.strip().title()


def setup_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(30)
    return driver


def wait_for_products(driver, timeout=15):
    """Wait until at least one product-like element appears."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                ".product-widget, .product-item, .product-card, [class*='product']"))
        )
    except TimeoutException:
        pass


def smart_scroll(driver):
    """Scroll gradually to trigger lazy-loaded products."""
    total_height = driver.execute_script("return document.body.scrollHeight")
    viewport = driver.execute_script("return window.innerHeight")
    current = 0
    step = viewport

    while current < total_height:
        current += step
        driver.execute_script(f"window.scrollTo(0, {current});")
        time.sleep(0.8)
        total_height = driver.execute_script("return document.body.scrollHeight")

    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.5)


def extract_products_from_page(driver):
    products = []
    seen_in_page = set()

    product_elements = driver.find_elements(
        By.CSS_SELECTOR,
        ".product-widget, .product-item, .product-card"
    )

    if not product_elements:
        print("  [!] No product elements found, using fallback link scan")
        all_links = driver.find_elements(By.TAG_NAME, "a")
        for link in all_links:
            href = link.get_attribute("href") or ""
            text = link.text.strip()

            if not text or not href:
                continue
            if BASE_URL not in href:
                continue
            if any(skip in href.lower() for skip in [
                "materials", "finishes", "resources", "story", "inspiration",
                "attic", "account", "login", "search", "cart", "collections",
                "pages", "blogs"
            ]):
                continue

            key = (format_product_name(text), href)
            if key in seen_in_page:
                continue
            seen_in_page.add(key)

            if text.isupper() and len(text.split()) >= 2:
                products.append({
                    "manufacturer": MANUFACTURER,
                    "source": href,
                    "product_name": format_product_name(text),
                })
    else:
        print(f"  Found {len(product_elements)} product elements in DOM")
        for elem in product_elements:
            try:
                raw_name = ""
                for selector in [
                    ".product-name",
                    "[itemprop='name']",
                    ".name-wishlist-container a",
                    "h2", "h3", "a"
                ]:
                    try:
                        el = elem.find_element(By.CSS_SELECTOR, selector)
                        text = el.text.strip()
                        if text:
                            raw_name = text
                            break
                    except:
                        continue

                if not raw_name:
                    continue

                href = ""
                for selector in ["a.p-0", "a[href*='/products/']", "a[href]"]:
                    try:
                        el = elem.find_element(By.CSS_SELECTOR, selector)
                        h = el.get_attribute("href") or ""
                        if h:
                            href = h if h.startswith("http") else BASE_URL + h
                            break
                    except:
                        continue

                key = (format_product_name(raw_name), href)
                if key in seen_in_page:
                    continue
                seen_in_page.add(key)

                products.append({
                    "manufacturer": MANUFACTURER,
                    "source": href,
                    "product_name": format_product_name(raw_name),
                })
            except Exception as e:
                print(f"  [!] Product parse error: {e}")

    return products


def get_all_page_numbers(driver):
    """
    Dynamically read ALL pagination buttons present in the DOM right now.
    Scans every .button-wrapper button inside .pagination regardless of nesting depth.
    Returns a sorted list of integer page numbers, e.g. [1, 2, 3].
    Returns [1] if no pagination found (single page).
    """
    try:
        buttons = driver.find_elements(
            By.CSS_SELECTOR,
            ".pagination .button-wrapper button"
        )
        page_numbers = []
        for btn in buttons:
            text = btn.text.strip().replace(" ", "")
            if text.isdigit():
                page_numbers.append(int(text))
        if page_numbers:
            pages = sorted(set(page_numbers))
            print(f"  Pagination buttons detected: {pages}")
            return pages
    except Exception as e:
        print(f"  [!] Could not read pagination buttons: {e}")
    print("  No pagination found — single page")
    return [1]


def click_page_button(driver, page_number: int, timeout=15):
    """
    Click whichever pagination button matches page_number.
    Button text is zero-padded like '01', '02', '03' etc.
    Re-reads buttons fresh every call so stale DOM is never an issue.
    Returns True on success, False if button not found.
    """
    target_label = f"{page_number:02d}"

    # Scroll pagination bar into view
    try:
        pagination = driver.find_element(By.CSS_SELECTOR, ".pagination")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", pagination)
        time.sleep(0.3)
    except:
        pass

    def find_and_click():
        # Always re-query to avoid stale references
        btns = driver.find_elements(
            By.CSS_SELECTOR,
            ".pagination .button-wrapper button"
        )
        for btn in btns:
            normalized = btn.text.strip().replace(" ", "").zfill(2)
            if normalized == target_label:
                driver.execute_script("arguments[0].click();", btn)
                return True
        return False

    try:
        clicked = find_and_click()
    except StaleElementReferenceException:
        time.sleep(1)
        clicked = find_and_click()  # One retry after stale

    if not clicked:
        print(f"  [!] Button for page '{target_label}' not found — skipping")
        return False

    print(f"  ✓ Clicked page button '{target_label}'")
    time.sleep(1.5)
    wait_for_products(driver, timeout=timeout)
    time.sleep(0.5)
    return True


def scrape_page_content(driver):
    """Scroll current view and extract all products."""
    smart_scroll(driver)
    return extract_products_from_page(driver)


def scrape_with_button_pagination(driver, start_url):
    """
    Load start_url, detect total pages from JS-rendered pagination buttons,
    then click through each page to harvest all products.
    """
    all_products = []

    print(f"  Loading: {start_url}")
    driver.get(start_url)
    wait_for_products(driver)
    time.sleep(1.5)  # Let React pagination render

    # Dynamically read whatever page buttons exist right now
    page_numbers = get_all_page_numbers(driver)
    total_pages = len(page_numbers)

    # --- Page 1 (already loaded, no click needed) ---
    first_page = page_numbers[0]
    print(f"  Scraping page {first_page} of {total_pages} (already loaded) ...")
    products = scrape_page_content(driver)
    print(f"  → {len(products)} products collected")
    all_products.extend(products)

    # --- Remaining pages: click each button dynamically ---
    for page_num in page_numbers[1:]:
        print(f"  Scraping page {page_num} of {total_pages} ...")
        success = click_page_button(driver, page_num)
        if not success:
            print(f"  [!] Skipping page {page_num}")
            continue

        products = scrape_page_content(driver)
        print(f"  → {len(products)} products collected")
        all_products.extend(products)

    return all_products


def save_to_excel(products, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(["Manufacturer", "Source", "Product Name"])
    for p in products:
        ws.append([p["manufacturer"], p["source"], p["product_name"]])
    wb.save(filename)
    print(f"\nExcel saved: {filename} ({len(products)} products)")


def main():
    print("Verellen Scraper - Button Pagination")
    print("=" * 50)

    driver = setup_driver()
    all_products = []

    try:
        for url in TARGET_URLS:
            print(f"\nScraping: {url}")
            products = scrape_with_button_pagination(driver, url)
            for p in products:
                print(f"   -> {p['product_name']}")
            all_products.extend(products)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

    if all_products:
        seen = set()
        unique = []
        for p in all_products:
            key = (p["product_name"], p["source"]) if p["source"] else p["product_name"]
            if key not in seen:
                seen.add(key)
                unique.append(p)

        print(f"\nTotal collected: {len(all_products)} | After dedup: {len(unique)}")
        save_to_excel(unique, OUTPUT_FILE)
    else:
        print("No products found.")


if __name__ == "__main__":
    main()