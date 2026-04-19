from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time

# Configure Chrome options
chrome_options = Options()
# Uncomment below to run without GUI (headless mode)
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")


def create_driver():
    return webdriver.Chrome(options=chrome_options)


def scroll_and_load(driver, max_scrolls=50, pause_time=2):
    """
    Scroll page to load all lazy-loaded products
    """
    print("🔄 Starting lazy load scroll...")

    previous_count = 0
    no_change_count = 0

    for scroll_num in range(max_scrolls):
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            print(f"↓ Scroll {scroll_num + 1}/{max_scrolls}")

            time.sleep(pause_time)

            products = driver.find_elements(By.CLASS_NAME, "product-list-item")
            current_count = len(products)

            print(f"  Products loaded: {current_count}")

            if current_count == previous_count:
                no_change_count += 1
                if no_change_count >= 3:
                    print(f"✓ All products loaded! Total: {current_count}")
                    break
            else:
                no_change_count = 0

            previous_count = current_count

        except Exception as e:
            print(f"⚠ Error during scroll: {e}")
            continue

    return driver.find_elements(By.CLASS_NAME, "product-list-item")


def extract_product_data(driver, products):
    """
    Extract product information from product elements
    """
    print(f"\n📊 Extracting data from {len(products)} products...")

    product_data = []

    for idx, product in enumerate(products, 1):
        try:
            # Extract Product URL
            product_link = product.find_element(By.CLASS_NAME, "slide__item")
            product_url = product_link.get_attribute('href')
            if not product_url.startswith('http'):
                product_url = "https://www.hectorfinch.com" + product_url

            # Extract Image URL
            try:
                img_tag = product.find_element(By.CLASS_NAME, "asp__bd")
                image_url = img_tag.get_attribute('src')
                image_url = image_url.split('?')[0] if '?' in image_url else image_url
            except:
                image_url = "N/A"

            # Extract Product Name
            try:
                product_name_tag = product.find_element(By.CLASS_NAME, "product-item__name")
                product_name = product_name_tag.text
            except:
                product_name = "N/A"

            # Extract SKU
            try:
                sku_tag = product.find_element(By.CLASS_NAME, "product-item-sku")
                sku = sku_tag.text
            except:
                sku = "N/A"

            product_data.append({
                'Product URL': product_url,
                'Image URL': image_url,
                'Product Name': product_name,
                'SKU': sku
            })

            print(f"✓ [{idx}] {product_name} ({sku})")

        except Exception as e:
            print(f"✗ Error extracting product {idx}: {e}")
            continue

    return product_data


def main():
    urls = [
        "https://www.hectorfinch.com/lighting/picture-lights",
        #"https://www.hectorfinch.com/lighting/wall-lights",
        #"https://www.hectorfinch.com/lighting/lanterns-on-brackets"
    ]

    all_product_data = []

    for url_idx, url in enumerate(urls, 1):
        print(f"\n{'=' * 60}")
        print(f"📄 URL {url_idx}/{len(urls)}: {url}")
        print(f"{'=' * 60}")

        driver = None
        try:
            print("🚀 Starting Chrome driver...")
            driver = create_driver()
            driver.get(url)

            print("⏳ Waiting 10-15 seconds for initial page load...")
            time.sleep(10)

            products = scroll_and_load(driver, max_scrolls=50, pause_time=2)
            product_data = extract_product_data(driver, products)
            all_product_data.extend(product_data)

            print(f"\n✓ Extracted {len(product_data)} products from this URL")

        except Exception as e:
            print(f"⚠ Error processing URL {url}: {e}")

        finally:
            if driver:
                print("🔌 Closing browser...")
                driver.quit()

    # Create DataFrame
    df = pd.DataFrame(all_product_data)

    # Save to Excel
    output_file = 'hector_finch_Lighting.xlsx'
    df.to_excel(output_file, index=False)

    print(f"\n{'=' * 60}")
    print(f"✅ SUCCESS!")
    print(f"{'=' * 60}")
    print(f"📊 Total products extracted: {len(all_product_data)}")
    print(f"💾 File saved: {output_file}")
    print(f"\nData preview:")
    print(df.head(10))


if __name__ == "__main__":
    main()