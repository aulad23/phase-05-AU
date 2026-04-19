# filename: tfp_normal_scrape.py
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time, re, os
from openpyxl import Workbook

# ================== CONFIG (change only these) ==================

# 1️⃣ Product listing page URLs (one or multiple)
PRODUCT_LIST_URLS = [
    "https://www.thefutureperfect.com/browse/furniture/tables/dining/",   # Put your category URL here
     #"https://www.thefutureperfect.com/browse/wall-coverings/claudy-jongstra/",  # Add more pages if needed

    ]

# 2️⃣ Output file name (without extension)
OUTPUT_NAME = "thefutureperfec_dining_Table"   # Final file will be: TFP_rugs.xlsx

# 3️⃣ Run browser in headless mode (True/False)
HEADLESS = True

# 4️⃣ Scroll pause time (seconds) – increase if site is slow
SCROLL_PAUSE = 5

# ===============================================================


def make_driver():
    """Create and return a Chrome WebDriver instance."""
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1400,1000")
    return webdriver.Chrome(options=opts)


def scroll_to_load_all(driver):
    """
    Scroll to the bottom of the page repeatedly
    until no new content is loaded.
    """
    previous_height = 0
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == previous_height:
            break
        previous_height = new_height


def extract_price_text(price_raw: str) -> str:
    """
    Extract a clean price string from raw text/HTML.
    Examples:
        '<span class="woocommerce-Price-currencySymbol">$</span>7,000'
        '$</span>7,000'
        '$7,000'
    Output should look like: '$7,000'
    """
    if not price_raw:
        return ""

    text = price_raw.strip()
    # Match patterns like $7,000 / €1,200 / £3,500
    m = re.search(r'[€£$]\s*[\d,]+(?:\.\d+)?', text)
    if m:
        # Remove any spaces inside the price
        return m.group(0).replace(" ", "")
    return text


def scrape_single_list_page(driver, url):
    """Scrape a single listing page and return product rows."""
    print(f"\n🔎 Scraping list page: {url}")
    driver.get(url)
    scroll_to_load_all(driver)

    products = driver.find_elements(By.CSS_SELECTOR, "li.product")
    print(f"   ➜ Found {len(products)} product elements on page")

    rows = []

    for p in products:
        try:
            # Product URL
            a_tag = p.find_element(By.CSS_SELECTOR, "a.woocommerce-LoopProduct-link")
            product_url = a_tag.get_attribute("href")

            # Product Name
            title = p.find_element(By.CSS_SELECTOR, "h2.woocommerce-loop-product__title").text.strip()

            # Image URL (first URL from data-srcset)
            try:
                img_div = p.find_element(By.CSS_SELECTOR, "div.image_wrapper.product_thumbnail div.lazy.image")
                srcset = img_div.get_attribute("data-srcset") or ""
                match = re.match(r"(https?://[^\s,]+)", srcset)
                image_url = match.group(1) if match else ""
            except:
                image_url = ""

            # List Price
            try:
                # Most WooCommerce themes use these selectors for price
                price_el = p.find_element(By.CSS_SELECTOR, "span.price, p.price, div.price")
                raw_price = price_el.get_attribute("innerText") or price_el.text
                list_price = extract_price_text(raw_price)
            except:
                list_price = ""

            rows.append([product_url, image_url, title, list_price])
        except Exception as e:
            print("   ⚠️ Skipped one product:", e)
            continue

    print(f"   ✅ Collected {len(rows)} rows from this page")
    return rows


def scrape_all_pages(url_list):
    """Loop through all listing URLs and combine product rows."""
    driver = make_driver()
    all_rows = []

    for url in url_list:
        try:
            rows = scrape_single_list_page(driver, url)
            all_rows.extend(rows)
        except Exception as e:
            print(f"❌ Error scraping {url}: {e}")

    driver.quit()

    # Remove duplicate products based on Product URL
    unique_rows = []
    seen = set()
    for row in all_rows:
        if row[0] not in seen:
            seen.add(row[0])
            unique_rows.append(row)

    print(f"\n🧮 Total unique products: {len(unique_rows)}")
    return unique_rows


def save_to_excel(output_name, rows):
    """Save scraped data into an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    # Header row (with new List Price column)
    ws.append(["Product URL", "Image URL", "Product Name", "List Price"])

    for row in rows:
        ws.append(row)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_name = f"{output_name}.xlsx"
    excel_path = os.path.join(script_dir, file_name)

    wb.save(excel_path)
    print(f"📁 File saved at: {excel_path}")


def main():
    """Main entry point."""
    if not PRODUCT_LIST_URLS:
        print("❌ PRODUCT_LIST_URLS is empty. Please add at least one URL.")
        return

    rows = scrape_all_pages(PRODUCT_LIST_URLS)
    if rows:
        save_to_excel(OUTPUT_NAME, rows)
        print("\n✅ Scraping completed!")
    else:
        print("\n⚠️ No products found.")


if __name__ == "__main__":
    main()
