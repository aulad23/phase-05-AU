import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import requests

# ---------- CONFIG ----------
category_url = "https://highlandhousefurniture.com/Consumer/ShowItems.aspx?TypeID=74"
output_file = r"C:\Users\ATS\Downloads\HighlandHouse_All_Product_Details.xlsx"
batch_size = 5
# ----------------------------

# Ensure output folder exists
os.makedirs(os.path.dirname(output_file), exist_ok=True)

# ---------- Selenium Setup ----------
options = webdriver.ChromeOptions()
# options.add_argument("--headless")  # Optional: run without opening browser
driver = webdriver.Chrome(options=options)  # Selenium Manager handles driver version

# ---------- Step 1: Scrape All Product URLs ----------
driver.get(category_url)
time.sleep(2)

all_product_urls = set()
page_num = 1

while True:
    print(f"Scraping page {page_num} for product URLs ...")
    time.sleep(2)

    # Collect product URLs
    product_elements = driver.find_elements(By.CSS_SELECTOR, "li.prodListingDiv div.prodSearchDiv a[href]")
    for elem in product_elements:
        href = elem.get_attribute("href")
        if href:
            all_product_urls.add(href)

    # Try clicking "View All"
    try:
        view_all = driver.find_element(By.CSS_SELECTOR, "span.viewAll.prodPageNavItem")
        if view_all.is_displayed():
            driver.execute_script("arguments[0].scrollIntoView(true);", view_all)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", view_all)
            time.sleep(4)
            print("✅ 'View All' clicked - all products loaded.")
            # Re-collect all product URLs from full list
            product_elements = driver.find_elements(By.CSS_SELECTOR, "li.prodListingDiv div.prodSearchDiv a[href]")
            for elem in product_elements:
                href = elem.get_attribute("href")
                if href:
                    all_product_urls.add(href)
            break
    except NoSuchElementException:
        pass

    # Try clicking Next page
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, "span.nextPage.prodPageNavItem")
        style = next_btn.get_attribute("style")
        if "hidden" in style or not next_btn.is_displayed():
            print("No more pages found. Stopping pagination.")
            break
        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", next_btn)
        page_num += 1
        time.sleep(3)
    except NoSuchElementException:
        print("No Next button found. Ending pagination.")
        break

print(f"\nTotal unique product URLs collected: {len(all_product_urls)}")
driver.quit()

# ---------- Step 2: Scrape Product Details ----------
all_data = []

def extract_dimension_value(soup, span_id):
    tag = soup.find('span', id=span_id)
    if tag and tag.text.strip():
        return tag.text.strip().replace("in", "").strip()
    return ""

for idx, url in enumerate(all_product_urls, start=1):
    try:
        response = requests.get(url, timeout=15)
        soup = BeautifulSoup(response.text, "html.parser")

        # Product Name
        name_tag = soup.find("span", id="ItemName")
        product_name = name_tag.text.strip() if name_tag else ""

        # SKU
        sku_tag = soup.find("span", id="ItemNumber")
        sku = sku_tag.text.strip() if sku_tag else ""

        # Image URL
        img_tag = soup.find("img", class_="prodSearchImage")
        image_url = img_tag['src'] if img_tag else ""

        # Dimensions
        width = extract_dimension_value(soup, "width")
        depth = extract_dimension_value(soup, "depth")
        height = extract_dimension_value(soup, "height")
        diameter = extract_dimension_value(soup, "diameter")

        # Weight
        weight = ""
        weight_tag = soup.find("tr", id="weightRow")
        if weight_tag:
            td_tag = weight_tag.find("td", id="weight")
            if td_tag:
                weight = td_tag.text.strip()

        # Description (optional)
        description = ""

        all_data.append({
            "Product URL": url,
            "Product Name": product_name,
            "SKU": sku,
            "Image URL": image_url,
            "Description": description,
            "Weight": weight,
            "Width": width,
            "Depth": depth,
            "Diameter": diameter,
            "Height": height
        })

        print(f"[{idx}/{len(all_product_urls)}] ✅ {product_name} | W:{width} D:{depth} H:{height} WT:{weight}")

        # Save in batches
        if idx % batch_size == 0 or idx == len(all_product_urls):
            df_out = pd.DataFrame(all_data)
            df_out.to_excel(output_file, index=False)
            print(f"💾 Saved batch up to product {idx} → {output_file}")

        time.sleep(1)

    except Exception as e:
        print(f"❌ Error scraping {url}: {e}")
        continue

print(f"\n🎉 All product details scraped successfully! Saved to: {output_file}")
