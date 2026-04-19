import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
)
from bs4 import BeautifulSoup
import re


def setup_driver():
    """Chrome driver setup with headless option"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    )
    # Image load disable kore speed baranor dorkar nai — image URL lagbe
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(60)
    return driver


def click_show_all_or_load_more(driver):
    """
    Page e 'Show All', 'Load More', 'View All' button thakle click kora.
    Magento site e prায়ই toolbar e 'show all' option thake.
    """
    selectors_to_try = [
        # "Show All" link/button (Magento toolbar)
        "//a[contains(text(), 'Show All')]",
        "//a[contains(text(), 'show all')]",
        "//a[contains(text(), 'View All')]",
        "//a[contains(text(), 'view all')]",
        "//option[contains(text(), 'All')]",
        # "Load More" button
        "//button[contains(text(), 'Load More')]",
        "//a[contains(text(), 'Load More')]",
        "//button[contains(@class, 'load-more')]",
        "//a[contains(@class, 'load-more')]",
        # Magento toolbar "show all" — limiter dropdown
        "//select[contains(@class, 'limiter')]//option[contains(text(), 'All')]",
    ]

    for xpath in selectors_to_try:
        try:
            elements = driver.find_elements(By.XPATH, xpath)
            for el in elements:
                if el.is_displayed():
                    # If it's an <option> inside a <select>, handle differently
                    tag = el.tag_name.lower()
                    if tag == 'option':
                        from selenium.webdriver.support.ui import Select
                        parent_select = el.find_element(By.XPATH, '..')
                        select_obj = Select(parent_select)
                        select_obj.select_by_visible_text(el.text.strip())
                        print(f"  -> Selected '{el.text.strip()}' from dropdown")
                    else:
                        driver.execute_script("arguments[0].click();", el)
                        print(f"  -> Clicked: '{el.text.strip()}'")
                    time.sleep(5)  # Wait for products to load
                    return True
        except Exception:
            continue

    return False


def scroll_and_load_all_products(driver, url):
    """
    Improved scroll function:
    1. First try 'Show All' / 'Load More' button
    2. Then scroll slowly to trigger lazy loading
    3. Multiple retry for stubborn lazy loaders
    """
    print(f"  Loading URL: {url}")
    driver.get(url)

    # Wait for page to initially load
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'li.product-item, .product-items'))
        )
        print("  -> Page loaded, product items found.")
    except TimeoutException:
        print("  -> Timeout waiting for products. Trying anyway...")

    time.sleep(3)

    # Step 1: Try to click "Show All" or "Load More"
    print("  Checking for 'Show All' / 'Load More' button...")
    found_button = click_show_all_or_load_more(driver)
    if found_button:
        time.sleep(5)

    # Step 2: Handle pagination — check if there are multiple pages
    handle_pagination = False
    all_page_sources = []

    try:
        pagination = driver.find_elements(By.CSS_SELECTOR, '.pages .items .item a')
        page_urls = set()
        for p in pagination:
            href = p.get_attribute('href')
            if href:
                page_urls.add(href)
        if page_urls:
            handle_pagination = True
            print(f"  -> Found {len(page_urls)} additional page(s)")
    except Exception:
        pass

    # Step 3: Scroll current page slowly
    print("  Scrolling to load all lazy-loaded products...")
    scroll_page_slowly(driver)

    all_page_sources.append(driver.page_source)

    # Step 4: If pagination, visit each page
    if handle_pagination:
        for page_url in sorted(page_urls):
            print(f"  Loading next page: {page_url}")
            driver.get(page_url)
            time.sleep(3)
            scroll_page_slowly(driver)
            all_page_sources.append(driver.page_source)

    # Also check "Next" button pagination
    while True:
        try:
            next_btn = driver.find_element(
                By.CSS_SELECTOR, '.pages .action.next'
            )
            if next_btn and next_btn.is_displayed():
                next_url = next_btn.get_attribute('href')
                if next_url and next_url not in [url] + list(page_urls if handle_pagination else []):
                    print(f"  Loading next page (via Next button): {next_url}")
                    driver.get(next_url)
                    time.sleep(3)
                    scroll_page_slowly(driver)
                    all_page_sources.append(driver.page_source)
                else:
                    break
            else:
                break
        except NoSuchElementException:
            break
        except Exception:
            break

    return all_page_sources


def scroll_page_slowly(driver):
    """Slowly scroll down the page to trigger lazy loading of images and products"""
    total_height = driver.execute_script("return document.body.scrollHeight")
    viewport_height = driver.execute_script("return window.innerHeight")
    current_position = 0
    scroll_step = viewport_height // 2  # Half viewport at a time

    no_change_count = 0

    while True:
        current_position += scroll_step
        driver.execute_script(f"window.scrollTo(0, {current_position});")
        time.sleep(1.5)  # Wait for lazy load

        new_total_height = driver.execute_script("return document.body.scrollHeight")

        if current_position >= new_total_height:
            # We've reached the bottom, but check if more content loaded
            if new_total_height == total_height:
                no_change_count += 1
                if no_change_count >= 3:
                    break
            else:
                total_height = new_total_height
                no_change_count = 0
        else:
            total_height = new_total_height
            no_change_count = 0

    # Final scroll to top and back to bottom (sometimes triggers remaining lazy loads)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)


def extract_product_data(html_content):
    """HTML theke product data extract kora — improved selectors"""
    soup = BeautifulSoup(html_content, 'html.parser')
    products = []

    # Sob product item khuje ber kora — multiple selector try
    product_items = soup.find_all('li', class_=lambda c: c and 'product-item' in c)

    if not product_items:
        # Alternate selector
        product_items = soup.select('li.item.product.product-item')

    if not product_items:
        # Another alternate
        product_items = soup.select('.products-grid .product-item')

    for item in product_items:
        try:
            # ---------- Product URL ----------
            product_url = ''
            # Method 1: product photo link
            url_tag = item.find('a', class_=lambda c: c and 'product' in c and 'photo' in c)
            if url_tag and url_tag.get('href'):
                product_url = url_tag['href']

            # Method 2: product-item-link
            if not product_url:
                url_tag = item.find('a', class_='product-item-link')
                if url_tag and url_tag.get('href'):
                    product_url = url_tag['href']

            # Method 3: any <a> with href containing the domain
            if not product_url:
                all_links = item.find_all('a', href=True)
                for link in all_links:
                    href = link['href']
                    if 'interludehome.com' in href and '/catalog/' not in href:
                        product_url = href
                        break

            # ---------- Image URL ----------
            image_url = ''
            img_tag = item.find('img', class_=lambda c: c and 'product-image' in c)
            if img_tag:
                # Check data-src first (lazy loaded images)
                image_url = (
                    img_tag.get('data-original')
                    or img_tag.get('data-src')
                    or img_tag.get('data-lazy')
                    or img_tag.get('src', '')
                )

            if not image_url:
                # Fallback: any img inside the product item
                img_tag = item.find('img')
                if img_tag:
                    image_url = (
                        img_tag.get('data-original')
                        or img_tag.get('data-src')
                        or img_tag.get('data-lazy')
                        or img_tag.get('src', '')
                    )

            # Skip placeholder images
            if image_url and 'placeholder' in image_url.lower():
                image_url = ''

            # ---------- Product Name ----------
            product_name = ''
            name_tag = item.find('a', class_='product-item-link')
            if name_tag:
                product_name = name_tag.get_text(strip=True)

            if not product_name:
                name_tag = item.find('strong', class_='product-item-name')
                if name_tag:
                    product_name = name_tag.get_text(strip=True)

            # ---------- SKU ----------
            sku = ''
            sku_div = item.find('div', class_=lambda c: c and 'sku' in c.lower()) if item else None
            if sku_div:
                sku_text = sku_div.find('div') or sku_div
                sku = sku_text.get_text(strip=True)

            if not sku:
                # Try data attribute
                sku_el = item.find(attrs={'data-sku': True})
                if sku_el:
                    sku = sku_el['data-sku']

            if not sku:
                # Try text search for SKU pattern
                all_text = item.get_text()
                sku_match = re.search(r'(?:SKU|sku)[:\s]*(\S+)', all_text)
                if sku_match:
                    sku = sku_match.group(1)

            # ---------- List Price ----------
            list_price = ''
            # Try special price first, then regular price
            price_box = item.find('div', class_='price-box')
            if price_box:
                # Check for old-price (original/MSRP price)
                old_price = price_box.find('span', class_='old-price')
                if old_price:
                    price_span = old_price.find('span', class_='price')
                else:
                    price_span = price_box.find('span', class_='price')

                if price_span:
                    price_text = price_span.get_text(strip=True)
                    list_price = price_text.replace('$', '').replace(',', '').strip()
            else:
                price_span = item.find('span', class_='price')
                if price_span:
                    price_text = price_span.get_text(strip=True)
                    list_price = price_text.replace('$', '').replace(',', '').strip()

            # ---------- Compile Data ----------
            product_data = {
                'Product URL': product_url,
                'Image URL': image_url,
                'Product Name': product_name,
                'SKU': sku,
                'List Price': list_price,
            }

            products.append(product_data)

        except Exception as e:
            print(f"  Error extracting product: {e}")
            continue

    return products


def deduplicate_products(products):
    """Duplicate product remove kora (SKU ba URL diye)"""
    seen = set()
    unique = []
    for p in products:
        # Use SKU + Product Name as unique key
        key = (p.get('SKU', ''), p.get('Product Name', ''), p.get('Product URL', ''))
        if key not in seen:
            seen.add(key)
            unique.append(p)
    return unique


def save_to_excel(products, filename='interlude_Dining_Tables.xlsx'):
    """Product data Excel file e save kora"""
    df = pd.DataFrame(products)

    # Reorder columns
    cols = ['Product URL', 'Image URL', 'Product Name', 'SKU', 'List Price']
    for c in cols:
        if c not in df.columns:
            df[c] = ''
    df = df[cols]

    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"\nData successfully saved to {filename}")
    print(f"Total products saved: {len(products)}")


def main():
    """Main function"""
    urls = [
        "https://interludehome.com/ih-collections/dining.html",
    ]

    print("=" * 60)
    print("  Interlude Home - Dining Tables Scraper (Fixed)")
    print("=" * 60)

    driver = setup_driver()

    try:
        all_products = []

        for url in urls:
            print(f"\nProcessing: {url}")
            print("-" * 50)

            # Load page(s) and get HTML
            page_sources = scroll_and_load_all_products(driver, url)

            print(f"\n  Extracting product data from {len(page_sources)} page(s)...")

            for i, html in enumerate(page_sources, 1):
                products = extract_product_data(html)
                print(f"    Page {i}: {len(products)} products found")
                all_products.extend(products)

        # Remove duplicates
        all_products = deduplicate_products(all_products)
        print(f"\nTotal unique products: {len(all_products)}")

        # Save to Excel
        if all_products:
            save_to_excel(all_products)

            # Sample data print
            print("\nSample data (first 5 products):")
            for i, product in enumerate(all_products[:5], 1):
                print(f"\n  Product {i}:")
                for key, value in product.items():
                    display_val = value[:80] + '...' if len(str(value)) > 80 else value
                    print(f"    {key}: {display_val}")

            # Check for missing data
            missing_url = sum(1 for p in all_products if not p['Product URL'])
            missing_img = sum(1 for p in all_products if not p['Image URL'])
            missing_name = sum(1 for p in all_products if not p['Product Name'])
            missing_sku = sum(1 for p in all_products if not p['SKU'])
            missing_price = sum(1 for p in all_products if not p['List Price'])

            print(f"\n--- Data Quality Report ---")
            print(f"  Missing Product URL: {missing_url}/{len(all_products)}")
            print(f"  Missing Image URL:   {missing_img}/{len(all_products)}")
            print(f"  Missing Product Name:{missing_name}/{len(all_products)}")
            print(f"  Missing SKU:         {missing_sku}/{len(all_products)}")
            print(f"  Missing List Price:  {missing_price}/{len(all_products)}")
        else:
            print("\nNo products found!")

    except Exception as e:
        print(f"\nError occurred: {e}")
        import traceback
        traceback.print_exc()

    finally:
        driver.quit()
        print("\nScraping completed!")


if __name__ == "__main__":
    main()