from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
import time
import os
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

START_URL = "https://www.sutherlandfurniture.com/suth/products/chairs/?subclass%5B%5D=ottoman#productListing__cards"
OUTPUT_NAME = "sutherland_ottoman.xlsx"

# ---------- helper: build next page url without breaking subclass params ----------
def build_page_url(start_url: str, page: int) -> str:
    parts = urlparse(start_url)
    qs = parse_qs(parts.query, keep_blank_values=True)

    # set / replace pg param
    qs["pg"] = [str(page)]

    new_query = urlencode(qs, doseq=True)

    return urlunparse((
        parts.scheme,
        parts.netloc,
        parts.path,
        parts.params,
        new_query,
        parts.fragment
    ))

# ---------- chrome ----------
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wb = Workbook()
ws = wb.active
ws.title = "Products"
ws.append(["Product URL", "Image URL", "Product Name"])

page = 1

while True:
    url = START_URL if page == 1 else build_page_url(START_URL, page)

    driver.get(url)
    time.sleep(6)

    product_cards = driver.find_elements(
        By.CSS_SELECTOR,
        "div.grid__lg-quarter div.suthQuickViewCard.productCard__quickview"
    )

    if not product_cards:
        print("No more cards found. Stopping.")
        break

    for card in product_cards:
        try:
            product_url = card.find_element(By.CSS_SELECTOR, "a.links__overlay").get_attribute("href")

            image_url = ""
            try:
                image_el = card.find_element(By.CSS_SELECTOR, "img.js-dynamic-image.productCard__img")
                image_url = image_el.get_attribute("src") or image_el.get_attribute("data-src") or ""
            except:
                image_url = ""

            product_name = card.find_element(
                By.CSS_SELECTOR, "div.productCard__name.notranslate"
            ).text.strip()

            ws.append([product_url, image_url, product_name])

        except:
            continue

    print(f"Page {page} scraped -> {url}")
    page += 1

driver.quit()

file_path = os.path.join(os.getcwd(), OUTPUT_NAME)
wb.save(file_path)
print(f"\nExcel file saved successfully:\n{file_path}")
