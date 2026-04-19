import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin

BASE_URL = "https://www.pierrefrey.com"

START_URLS = [
    "https://www.pierrefrey.com/en/carpets-rugs",
    #"https://www.pierrefrey.com/en/furniture?types[]=ottomans&q=&productPerLine=3",
    #"https://www.pierrefrey.com/en/furniture?types[]=poufs&q=&productPerLine=3",
    #"https://www.pierrefrey.com/en/furniture?types[]=armless_sofas&q=&productPerLine=3",
]

headers = {
    "User-Agent": "Mozilla/5.0"
}

visited_pages = set()
all_products = []


def scrape_page(page_url):
    response = requests.get(page_url, headers=headers, timeout=30)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    items = soup.find_all("div", class_="resultListItem")

    for item in items:
        link_tag = item.find("a", class_="resultListItem__link")
        img_tag = item.find("img", class_="resultListItem__img")
        sup_title = item.find("div", class_="resultListItem__supTitle")
        title = item.find("div", class_="resultListItem__title")
        sub_title = item.find("div", class_="resultListItem__subTitle")

        if not link_tag or not img_tag:
            continue

        product_url = urljoin(BASE_URL, link_tag.get("href"))
        image_url = img_tag.get("src")

        product_name = ""
        if sup_title and title:
            product_name = f"{sup_title.get_text(strip=True)} {title.get_text(strip=True)}"

        sku = sub_title.get_text(strip=True) if sub_title else ""

        all_products.append({
            "Product URL": product_url,
            "Image URL": image_url,
            "Product Name": product_name,
            "SKU": sku
        })

    return soup


# -------- Start Scraping --------
for start_url in START_URLS:
    if start_url in visited_pages:
        continue

    print(f"Scraping page 1: {start_url}")
    soup = scrape_page(start_url)
    visited_pages.add(start_url)

    while True:
        pagination_links = soup.select(
            "ul.pagination__list a.pagination__button--num"
        )

        new_page_found = False

        for link in pagination_links:
            href = link.get("href")
            full_url = urljoin(BASE_URL, href)

            if full_url not in visited_pages:
                print(f"Scraping {full_url}")
                visited_pages.add(full_url)
                soup = scrape_page(full_url)
                new_page_found = True
                break

        if not new_page_found:
            break


# -------- Excel Output --------
df = pd.DataFrame(all_products)
df.to_excel("pierrefrey_carpets.xlsx", index=False)

print(f"Completed. Total products scraped: {len(df)}")
