from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
from time import sleep


def scrape_with_selenium():
    """
    Scrape using Selenium (browser automation) - Fixed duplicate issue
    """
    all_products = []
    seen_urls = set()

    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=chrome_options)

    try:
        base_urls = [
            "https://www.newclassicfurniture.com/product-category/living-room/sectionals/",
            #"https://www.newclassicfurniture.com/product-category/living-room/stationary/"
        ]

        # ✅ Multi-link support (double link thakle duita tei scrape hobe)
        for base_url in base_urls:
            page = 1

            while True:
                if page == 1:
                    url = base_url
                else:
                    url = f"{base_url}page/{page}/"

                print(f"Scraping page {page}: {url}")

                driver.get(url)
                sleep(3)

                products = driver.find_elements(
                    By.CSS_SELECTOR,
                    '.woolentor-grid-view-content .woolentor-product-image'
                )

                if not products:
                    print("No products found. Ending scrape for this base_url.")
                    break

                page_product_count = 0

                for product in products:
                    try:
                        link = product.find_element(By.TAG_NAME, 'a')
                        product_url = link.get_attribute('href')

                        if not product_url or product_url in seen_urls:
                            continue

                        product_name = link.get_attribute('title')

                        img = link.find_element(By.TAG_NAME, 'img')
                        image_url = img.get_attribute('src')

                        seen_urls.add(product_url)

                        all_products.append({
                            'Product URL': product_url,
                            'Image URL': image_url,
                            'Product Name': product_name
                        })

                        page_product_count += 1
                        print(f"  Found: {product_name}")

                    except Exception as e:
                        print(f"  Error extracting product: {e}")
                        continue

                print(f"  Total unique products on page {page}: {page_product_count}")

                try:
                    driver.find_element(By.CSS_SELECTOR, 'a.next.page-numbers')
                    page += 1
                    sleep(2)
                except:
                    print("No more pages for this base_url.")
                    break

    finally:
        driver.quit()

    return all_products


if __name__ == "__main__":
    print("Starting web scraping with Selenium...")
    print("=" * 50)

    products = scrape_with_selenium()

    if products:
        df = pd.DataFrame(products)
        df = df.drop_duplicates(subset=['Product URL'], keep='first')

        df.to_excel('bedroom_Sectionals.xlsx', index=False, engine='openpyxl')
        print(f"\n✓ Data saved to bedroom_Sectionals.xlsx")
        print(f"✓ Total unique products scraped: {len(df)}")

        print("\nFirst 5 products:")
        print("=" * 50)
        for i, product in enumerate(products[:5], 1):
            print(f"\n{i}. {product['Product Name']}")
            print(f"   URL: {product['Product URL']}")
            print(f"   Image: {product['Image URL'][:80]}...")

    else:
        print("No products were scraped.")
