import time
import pandas as pd
from urllib.parse import urljoin

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


BASE_URL = "https://www.brightchair.com/"
START_URLS = [
    "https://www.brightchair.com/_seating/swivels",
    #"https://www.brightchair.com/_seating/traditional"
]
OUTPUT_XLSX = "brightchair_Desk_Chairs.xlsx"


def make_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    # chrome_options.add_argument("--headless=new")  # চাইলে চালু করবেন
    return webdriver.Chrome(options=chrome_options)


def scroll_to_load_all(driver, pause=1.2, max_idle_rounds=6):
    idle = 0
    last_count = 0

    while True:
        items = driver.find_elements(By.CSS_SELECTOR, "ul#rig.rig > a")
        count = len(items)

        if count > last_count:
            last_count = count
            idle = 0
        else:
            idle += 1

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)

        if idle >= max_idle_rounds:
            break


def extract_slug(raw_href: str) -> str:
    if not raw_href:
        return ""
    h = raw_href.strip()

    if "#" in h:
        h = h.split("#")[-1].strip()

    h = h.rstrip("/")
    return h.split("/")[-1].strip()


def scrape_one_list(driver, list_url: str):
    driver.get(list_url)

    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "ul#rig.rig"))
    )

    scroll_to_load_all(driver)

    rows = []
    anchors = driver.find_elements(By.CSS_SELECTOR, "ul#rig.rig > a")

    for a in anchors:
        raw_href = (a.get_attribute("href") or "").strip()
        slug = extract_slug(raw_href)
        if not slug:
            continue

        product_url = f"{list_url}#{slug}"
        product_name = slug

        img_url = ""
        try:
            img = a.find_element(By.CSS_SELECTOR, "li img")
            src = (img.get_attribute("src") or "").strip()
            data_original = (img.get_attribute("data-original") or "").strip()
            chosen = data_original if data_original else src
            img_url = chosen if chosen.startswith("http") else urljoin(BASE_URL, chosen)
        except Exception:
            img_url = ""

        rows.append({
            "Product URL": product_url,
            "Image URL": img_url,
            "Product Name": product_name
        })

    return rows


def main():
    driver = make_driver()
    try:
        all_rows = []
        for url in START_URLS:
            all_rows.extend(scrape_one_list(driver, url))

        df = pd.DataFrame(all_rows).drop_duplicates(subset=["Product URL"]).reset_index(drop=True)
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"✅ Saved: {OUTPUT_XLSX} | Rows: {len(df)}")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
