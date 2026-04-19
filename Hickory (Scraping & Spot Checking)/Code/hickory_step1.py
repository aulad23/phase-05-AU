import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# ---------- SETUP ----------
chrome_driver_path = "C:/chromedriver.exe"

options = Options()
options.add_argument("--start-maximized")

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

# ---------- URL LIST ----------
urls = [
    "https://www.hickorychair.com/Products/ShowResults?TypeID=76&SearchName=Ottomans+%26+Benches",


]

data = []
base_url = "https://www.hickorychair.com"

# ---------- LOOP THROUGH EACH PAGE ----------
for page_url in urls:
    driver.get(page_url)
    time.sleep(3)

    # ---------- SCROLL TO LOAD ALL PRODUCTS ----------
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)  # wait for lazy load
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    # ---------- SCRAPE DATA ----------
    products = driver.find_elements(By.CSS_SELECTOR, "div.search-item")

    for idx, product in enumerate(products, start=1):
        try:
            link_elem = product.find_element(By.TAG_NAME, "a")
            href = link_elem.get_attribute("href")

            # ✅ Fix: Add base URL only if relative
            if href.startswith("/"):
                product_url = base_url + href
            else:
                product_url = href

            # Image
            try:
                img_elem = product.find_element(By.TAG_NAME, "img")
                image_url = img_elem.get_attribute("src")
                if "_small" in image_url:
                    image_url = image_url.replace("_small", "_hires")
                if image_url.startswith("/"):
                    image_url = base_url + image_url
            except:
                image_url = ""

            # SKU
            try:
                sku = product.find_element(By.CLASS_NAME, "search-item-sku").text.strip()
            except:
                sku = ""

            # Product Name
            try:
                name = product.find_element(By.CLASS_NAME, "search-item-name").text.strip()
            except:
                name = ""

            row = {
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": name,
                "SKU": sku
            }
            data.append(row)

            # ---- PRINT IN TERMINAL ----
            print(f"\nProduct {idx}")
            print(f"URL: {product_url}")
            print(f"Image: {image_url}")
            print(f"Name: {name}")
            print(f"SKU: {sku}")
            print("-" * 50)

        except Exception as e:
            print("❌ Error on product:", e)

# ---------- CLOSE DRIVER ----------
driver.quit()

# ---------- SAVE TO EXCEL ----------
df = pd.DataFrame(data)
df.drop_duplicates(subset=["Product URL"], inplace=True)
df.to_excel("Hocker_Ottomans.xlsx", index=False)

print(f"\n✅ Scraping complete! {len(df)} unique products saved to Hocker_Dressers_Chests.xlsx")
