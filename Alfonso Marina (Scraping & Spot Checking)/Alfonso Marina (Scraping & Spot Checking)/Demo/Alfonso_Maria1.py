from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
import pandas as pd
import time

# -----------------------------
# Step 1: Setup Chrome
# -----------------------------
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Comment out to see browser
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# -----------------------------
# URLs (must be separate strings + comma)
# -----------------------------
urls = [
    "https://alfonsomarina.com/product-category/furniture-eng/accessories-eng/boxes-eng/",
    #"https://alfonsomarina.com/product-category/furniture-eng/storage-eng/buffets-sideboards-eng/",
]

# -----------------------------
# Helper: pick real image url
# -----------------------------
def pick_real_image(img_tag):
    if not img_tag:
        return "N/A"

    src = (img_tag.get("src") or "").strip()

    for key in ["data-src", "data-lazy-src", "data-original", "data-ks-lazyload", "data-srcset"]:
        val = (img_tag.get(key) or "").strip()
        if val and "1px.png" not in val:
            if "," in val and (" " in val):
                parts = [p.strip().split(" ")[0] for p in val.split(",") if p.strip()]
                if parts:
                    return parts[-1]
            return val

    srcset = (img_tag.get("srcset") or "").strip()
    if srcset:
        parts = [p.strip().split(" ")[0] for p in srcset.split(",") if p.strip()]
        if parts:
            return parts[-1]

    if src and "1px.png" not in src:
        return src

    return "N/A"

# -----------------------------
# Main scrape loop
# -----------------------------
all_data = []

for url in urls:
    driver.get(url)

    # Wait products
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.registroProducto")))

    # Scroll to load all products
    last_count = 0
    same_count_rounds = 0

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2.5)

        cards = driver.find_elements(By.CSS_SELECTOR, "div.registroProducto")
        current_count = len(cards)

        if current_count == last_count:
            same_count_rounds += 1
        else:
            same_count_rounds = 0

        if same_count_rounds >= 3:
            break

        last_count = current_count

    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    products = soup.find_all("div", class_="registroProducto")

    for product in products:
        try:
            product_url = product.find("a", class_="registroImagen")["href"]

            img_tag = product.find("img")
            image_url = pick_real_image(img_tag)

            name_tag = product.find("a", class_="registroTitulo")
            product_name = name_tag.get_text(strip=True) if name_tag else "N/A"

            all_data.append({
                "Category URL": url,
                "Product Name": product_name,
                "Product URL": product_url,
                "Image URL": image_url
            })
        except Exception as e:
            print("Skipping product due to error:", e)
            continue

driver.quit()

# Save
df = pd.DataFrame(all_data)
df.to_excel("alfonsomarina_Boxes.xlsx", index=False)
print("✅ Scraping complete! File saved as alfonsomarina_Boxes.xlsx")
