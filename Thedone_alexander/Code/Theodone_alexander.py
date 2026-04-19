from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from bs4 import BeautifulSoup
import time
import pandas as pd

# --- Setup Selenium ---
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

# 👈 Put all category URLs here
category_urls = [
    "https://theodorealexander.com/item/category/type/value/floor-lighting",

]

data = []

for url in category_urls:
    page = 1
    total_pages = 1 # manual page limit per URL

    driver.get(url)
    time.sleep(3)

    while page <= total_pages:
        print(f"Scraping {url}, page {page}...")

        # Wait for product grid to load
        try:
            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.info")))
        except TimeoutException:
            print("❌ Timeout waiting for products to load.")
            break

        # --- Parse products ---
        soup = BeautifulSoup(driver.page_source, "html.parser")
        products = soup.find_all("div", class_="info")

        for product in products:
            a_tag = product.find("a", class_="productImage")
            product_url = "https://theodorealexander.com" + a_tag["href"] if a_tag else None

            img_tag = a_tag.find("img") if a_tag else None
            image_url = img_tag["src"] if img_tag else None

            name_tag = product.find("div", class_="name").find("a")
            product_name = name_tag["title"] if name_tag else None

            sku_tag = product.find("div", class_="sku")
            sku = sku_tag.text.strip() if sku_tag else None

            data.append({
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name,
                "SKU": sku
            })

        # Stop if reached manual page limit
        if page == total_pages:
            break

        # --- Click Next Button ---
        try:
            next_button = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.page-link.first-last[data-page='next']")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
            time.sleep(1)
            next_button.click()
            page += 1
            time.sleep(2)
        except (TimeoutException, StaleElementReferenceException):
            print("⚠️ Next button not found — stopping early.")
            break

# --- Cleanup ---
driver.quit()

# --- Save to Excel ---
df = pd.DataFrame(data)
df.to_excel("theodore_alexander_Floor_lamps.xlsx", index=False)
print(f"✅ Done! {len(df)} products scraped and saved.")
