import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def setup_driver():
    options = Options()
    # Uncomment for headless mode (browser not visible)
    # options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    # ChromeDriver path
    service = Service(r"C:\chromedriver-win64\chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def load_all_products(driver, url):
    driver.get(url)
    time.sleep(3)  # initial load

    while True:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, "div.load-more button")
            driver.execute_script("arguments[0].scrollIntoView();", btn)
            time.sleep(1)
            btn.click()
            time.sleep(2)  # wait new products
        except:
            # no more "See more" button
            break

    return driver.page_source

def parse_products(html):
    soup = BeautifulSoup(html, "html.parser")
    products = []
    product_list = soup.find("ul", class_="product-catalogue__list")

    if not product_list:
        return products

    for li in product_list.find_all("li"):
        a = li.find("a", href=True)
        img = li.find("img", class_="img-responsive")
        caption = li.find("figcaption")

        if not a or not img or not caption:
            continue

        # ✅ Desired column order
        product_url = a["href"].strip()
        image_url = img.get("src", "").strip()
        product_name = caption.find(text=True, recursive=False)
        if product_name:
            product_name = product_name.strip()

        products.append({
            "Product URL": product_url,
            "Image URL": image_url,
            "Product Name": product_name
        })

    return products

if __name__ == "__main__":
    category_urls = [
        "https://www.studioliaigre.com/en/furniture-and-lighting/decorative-items/",
        #"https://www.studioliaigre.com/en/furniture-and-lighting/seats/banquettes/"
    ]

    driver = setup_driver()
    all_products = []

    try:
        for url in category_urls:
            html = load_all_products(driver, url)
            products = parse_products(html)
            all_products.extend(products)
    finally:
        driver.quit()

    # Save to Excel
    df = pd.DataFrame(all_products, columns=["Product URL", "Image URL", "Product Name"])
    df.to_excel("studioliaigre_Objects.xlsx", index=False)

    print(f"Total products scraped: {len(all_products)}")
