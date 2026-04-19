from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os

# =========================
# Config
# =========================
last_page = 1  # <-- set total number of pages

# Multiple categories supported
base_urls = [
    "https://sunpan.com/collections/pendants",

]

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 15)
all_data = []

# =========================
# Scraping Loop
# =========================
for base_url in base_urls:
    for page in range(1, last_page + 1):

        # Build correct URL
        url = base_url if page == 1 else f"{base_url}?page={page}"

        print(f"🔍 Loading: {url}")
        driver.get(url)
        time.sleep(3)

        # Lazy load scroll
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Wait until product cards load
        try:
            cards = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "div.card-wrapper.product-card-wrapper")
            ))
        except:
            cards = []

        print(f"Page {page}: {len(cards)} products found")

        # Extract product data
        for card in cards:

            # Product URL
            try:
                a_tag = card.find_element(By.XPATH, "./ancestor::a[1]")
                product_url = a_tag.get_attribute("href").strip()
            except:
                product_url = ""

            # SKU
            try:
                sku = card.find_element(By.CSS_SELECTOR, "div.product__sku span.sku").text.strip()
            except:
                sku = ""

            # Product Name
            try:
                product_name = card.find_element(By.CSS_SELECTOR, "h3.card__heading.h3").text.strip()
            except:
                product_name = ""

            # Image URL
            try:
                image_url = card.find_element(By.CSS_SELECTOR, "img.card-product-image").get_attribute("src").strip()
            except:
                image_url = ""

            all_data.append({
                "Product URL": product_url,
                "Product Name": product_name,
                "Image URL": image_url,
                "SKU": sku
            })

# =========================
# Save to Excel
# =========================
script_folder = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_folder, "sunpan-pendants.xlsx")

df = pd.DataFrame(all_data)
df = df[["Product URL", "Image URL", "Product Name", "SKU"]]

df.to_excel(file_path, index=False)

driver.quit()
print(f"✅ Step 1 complete. Data saved to {file_path}")
