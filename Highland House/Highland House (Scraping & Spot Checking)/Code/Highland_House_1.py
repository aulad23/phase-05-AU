import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# ---------- CONFIG ----------
category_url = "https://highlandhousefurniture.com/Consumer/ShowItems.aspx?TypeID=74"
save_path = r"C:\Users\ATS\Downloads\HighlandHouse_Coffee_Cocktail_Tables_Products.xlsx"
base_url = "https://highlandhousefurniture.com"
# ----------------------------

# Chrome driver setup (NO PATH NEEDED)
options = webdriver.ChromeOptions()
# options.add_argument("--headless")  # Optional
driver = webdriver.Chrome(options=options)

wait = WebDriverWait(driver, 10)
driver.get(category_url)
time.sleep(2)

all_html = ""
page_num = 1

while True:
    print(f"Scraping page {page_num} ...")
    time.sleep(2)
    all_html += driver.page_source

    # Try click "View All"
    try:
        view_all = driver.find_element(By.CSS_SELECTOR, "span.viewAll.prodPageNavItem")
        driver.execute_script("arguments[0].scrollIntoView(true);", view_all)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", view_all)
        time.sleep(3)
        print("✅ 'View All' clicked — all products loaded on one page.")
        all_html = driver.page_source
        break
    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"⚠️ Could not click 'View All': {e}")

    # Try click Next page
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, "span.nextPage.prodPageNavItem")
        if "visibility: hidden" in next_btn.get_attribute("style"):
            print("No more pages found. Stopping.")
            break
        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", next_btn)
        page_num += 1
        time.sleep(3)
    except NoSuchElementException:
        print("No Next button found — ending pagination.")
        break

driver.quit()

# ---------- SCRAPE PRODUCTS ----------
print("Extracting product data...")

soup = BeautifulSoup(all_html, "html.parser")
li_items = soup.find_all("li", class_="prodListingDiv")

products = []
for li in li_items:
    div = li.find("div", class_="prodSearchDiv")

    # Product URL
    a_tag = div.find("a", href=True)
    product_url = f"{base_url}{a_tag['href']}" if a_tag else "N/A"

    # Image URL
    img_tag = div.find("img", class_="prodSearchImage")
    image_url = f"{base_url}{img_tag['src']}" if img_tag else "N/A"

    # SKU
    sku_tag = div.find("div", style=lambda x: x and "margin:5px 0;" in x)
    sku = sku_tag.find("strong").text.strip() if sku_tag else "N/A"

    products.append({
        "Product Url": product_url,
        "Image Url": image_url,
        "Sku": sku
    })

# ---------- SAVE TO EXCEL ----------
df = pd.DataFrame(products)
df.to_excel(save_path, index=False)

print(f"\n✅ Scraping complete! {len(products)} products saved to:\n{save_path}")
