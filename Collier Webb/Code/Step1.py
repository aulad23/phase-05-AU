import undetected_chromedriver as uc
from bs4 import BeautifulSoup
import time
from openpyxl import Workbook

urls = [
    "https://collierwebb.com/us/hardware/escutcheons",
    #"https://collierwebb.com/us/hardware/cupboard-knobs"
]

# Undetected Chrome setup
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")

driver = uc.Chrome(options=options, version_main=None)

# -------- EXCEL SETUP --------
wb = Workbook()
ws = wb.active
ws.title = "Products"
ws.append(["Product URL", "Image URL", "Product Name"])

seen = set()

# -------- LOOP THROUGH URLS --------
for url in urls:
    print(f"\n{'=' * 60}")
    print(f"Scraping: {url}")
    print(f"{'=' * 60}")

    driver.get(url)

    # Cloudflare bypass wait - extra time debo
    print("⏳ Waiting for page to fully load (Cloudflare check)...")
    time.sleep(15)

    # Check if still blocked
    if "blocked" in driver.page_source.lower() or "you have been blocked" in driver.page_source.lower():
        print("❌ Still blocked by Cloudflare!")
        print("💡 Please solve CAPTCHA manually if it appears...")
        time.sleep(30)  # Manual CAPTCHA solve er jonno time

    # -------- AGGRESSIVE SCROLL (LAZY LOAD) --------
    print("🔄 Starting scroll to load all products...")

    scroll_pause_time = 6  # Aro beshi wait
    last_height = driver.execute_script("return document.body.scrollHeight")
    scroll_count = 0
    max_scrolls = 25
    no_change_count = 0

    while scroll_count < max_scrolls:
        # Slowly scroll down (more human-like)
        current_position = 0
        scroll_step = last_height // 5

        for step in range(5):
            current_position += scroll_step
            driver.execute_script(f"window.scrollTo(0, {current_position});")
            time.sleep(1)

        # Final scroll to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        print(f"📜 Scroll #{scroll_count + 1} - Waiting {scroll_pause_time}s...")
        time.sleep(scroll_pause_time)

        # Check products loaded
        temp_soup = BeautifulSoup(driver.page_source, "html.parser")
        current_products = len(temp_soup.select("li.item.product.product-item"))
        print(f"   → Currently visible: {current_products} products")

        # Calculate new height
        new_height = driver.execute_script("return document.body.scrollHeight")

        if new_height == last_height:
            no_change_count += 1
            print(f"   → Height unchanged ({no_change_count}/3)")
            if no_change_count >= 3:
                print("   ✓ All products loaded!")
                break
        else:
            no_change_count = 0

        last_height = new_height
        scroll_count += 1

    # Final wait
    print("⏸ Final wait before extracting data...")
    time.sleep(5)

    # -------- PARSE --------
    soup = BeautifulSoup(driver.page_source, "html.parser")
    products = soup.select("li.item.product.product-item")

    print(f"\n📦 Total products found: {len(products)}")

    if len(products) == 0:
        # Save HTML for debugging
        debug_file = f"debug_page_{urls.index(url)}.html"
        with open(debug_file, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"⚠️ No products found! HTML saved to {debug_file}")

        # Screenshot niye rakhun
        screenshot_file = f"debug_screenshot_{urls.index(url)}.png"
        driver.save_screenshot(screenshot_file)
        print(f"📸 Screenshot saved to {screenshot_file}")
        continue

    # Extract data
    added_count = 0
    for idx, product in enumerate(products, 1):
        name_tag = product.select_one("a.product-item-link")
        img_tag = product.select_one("img")

        if not name_tag or not img_tag:
            continue

        product_url = name_tag.get("href", "")
        product_name = name_tag.get_text(strip=True)
        image_url = img_tag.get("src", "")

        if product_url in seen:
            continue

        seen.add(product_url)
        ws.append([product_url, image_url, product_name])
        added_count += 1

    print(f"✅ Added {added_count} new products from this page")

print("\n🔒 Closing browser...")
driver.quit()

# -------- SAVE --------
file_name = "collierwebb_Backplates.xlsx"
wb.save(file_name)

print(f"\n{'=' * 60}")
print(f"✅ Excel saved: {file_name}")
print(f"📊 Total unique products: {len(seen)}")
print(f"{'=' * 60}")