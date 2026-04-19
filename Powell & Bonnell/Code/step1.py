from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
from bs4 import BeautifulSoup
import re

URLS = [
    "https://www.powellandbonnell.com/product-category/mirrors/",
    #"https://www.powellandbonnell.com/product-category/tables/?_categories=coffee-tables",
    #"https://www.powellandbonnell.com/product-category/tables/?_categories=bar-tables"
]

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

rows = []
seen_urls = set()

for url in URLS:
    driver.get(url)
    time.sleep(2)

    # Infinite scroll
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    soup = BeautifulSoup(driver.page_source, "html.parser")
    products = soup.find_all("li", class_="product")

    for product in products:
        a_tag = product.find("a", class_="woocommerce-LoopProduct-link")
        product_url = a_tag["href"] if a_tag else ""

        if not product_url or product_url in seen_urls:
            continue
        seen_urls.add(product_url)

        title_tag = product.find("h2", class_="woocommerce-loop-product__title")
        product_name = title_tag.get_text(strip=True) if title_tag else ""

        img_div = product.find("div", class_="product__img")
        image_url = ""
        if img_div and img_div.has_attr("style"):
            match = re.search(r"url\(['\"]?(.*?)['\"]?\)", img_div["style"])
            if match:
                image_url = match.group(1)

        rows.append([
            product_url,
            image_url,
            product_name
        ])

driver.quit()

df = pd.DataFrame(rows, columns=[
    "Product URL",
    "Image URL",
    "Product Name"
])

df.to_excel("powell_bonnell_Mirrors.xlsx", index=False)

print("✅ Excel saved with correct column order")
