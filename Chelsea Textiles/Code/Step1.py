import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://www.chelseatextiles.com"
START_URLS = [
    "https://www.chelseatextiles.com/us/furniture/categories/tables-consoles-desks",
    #"https://www.chelseatextiles.com/us/wallpaper/patrick-kinmonth",
    #"https://www.chelseatextiles.com/us/wallpaper/chelsea-textiles-collection",
    #"https://www.chelseatextiles.com/us/fabrics/embroidery/delicate-vines",
    #"https://www.chelseatextiles.com/us/fabrics/embroidery/small-sprigs",
    #"https://www.chelseatextiles.com/us/fabrics/prints-wovens/prints",
    #"https://www.chelseatextiles.com/us/fabrics/prints-wovens/checks-stripes",
    #"https://www.chelseatextiles.com/us/fabrics/prints-wovens/silks",
    #"https://www.chelseatextiles.com/us/fabrics/prints-wovens/textures",
    #"https://www.chelseatextiles.com/us/fabrics/prints-wovens/velvets",
    #"https://www.chelseatextiles.com/us/fabrics/prints-wovens/wools-plains",
    #"https://www.chelseatextiles.com/us/fabrics/designers/alidad",
    #"https://www.chelseatextiles.com/us/fabrics/designers/domenica-more-gordon",
    #"https://www.chelseatextiles.com/us/fabrics/designers/kit-kemp",
    #"https://www.chelseatextiles.com/us/fabrics/designers/neisha-crosland",
    #"https://www.chelseatextiles.com/us/fabrics/designers/patrick-kinmonth",
    #"https://www.chelseatextiles.com/us/fabrics/designers/robert-kime"
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

data = []


def get_soup(url):
    """Fetch and parse HTML"""
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return BeautifulSoup(r.text, "lxml")


def get_total_pages(url):
    """Get total number of pages for a category"""
    try:
        soup = get_soup(url)
        pagination = soup.select(".c-pagination__page")

        pages = [1]
        for p in pagination:
            text = p.get_text(strip=True)
            if text.isdigit():
                pages.append(int(text))

        return max(pages) if pages else 1
    except Exception as e:
        print(f"  → Error getting pages: {e}")
        return 1


def scrape_page(url):
    """Scrape a single page and return product data"""
    soup = get_soup(url)

    # Select all items - updated to select any data-type (fabrics, cushions, etc.)
    items = soup.select('li.o-gridlist__item[data-type]')

    page_data = []
    for item in items:
        try:
            # Get category type (fabrics, cushions, etc.)
            category_type = item.get('data-type', 'unknown')

            # Get product URL from <a> tag
            a_tag = item.select_one("a.c-gridcard")
            if not a_tag or not a_tag.get("href"):
                continue
            product_url = a_tag["href"]
            if not product_url.startswith("http"):
                product_url = BASE_URL + product_url

            # Get image URL from <img> tag
            img_tag = item.select_one("img.c-gridcard__img")
            image_url = img_tag["src"] if img_tag and img_tag.get("src") else ""

            # Get SKU and Product Name from <h4> inside <div class="c-gridcard__text">
            title_tag = item.select_one("h4.c-gridcard__title")

            sku = ""
            product_name = ""

            if title_tag:
                # Get all text content and split by <br>
                parts = title_tag.get_text(separator="|", strip=True).split("|")

                if len(parts) >= 2:
                    sku = parts[0].strip()
                    product_name = parts[1].strip()
                elif len(parts) == 1:
                    # Sometimes only product name exists
                    product_name = parts[0].strip()

            page_data.append({
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name,
                "SKU": sku,
                "Category Type": category_type
            })

        except Exception as e:
            print(f"  → Error parsing item: {e}")
            continue

    return page_data


# --------- Main Scraping Logic ----------
print("=" * 60)
print("Starting Chelsea Textiles Scraper")
print("=" * 60)

for category_url in START_URLS:
    print(f"\n{'=' * 60}")
    print(f"Processing Category: {category_url}")
    print(f"{'=' * 60}")

    # Get total pages for this category
    total_pages = get_total_pages(category_url)
    print(f"Total Pages Found: {total_pages}")

    # Loop through all pages in this category
    for page in range(1, total_pages + 1):
        if page == 1:
            url = category_url
        else:
            url = f"{category_url}/p{page}"

        print(f"Scraping Page {page}: {url}")

        try:
            page_data = scrape_page(url)
            data.extend(page_data)
            print(f"  → Found {len(page_data)} products")
        except Exception as e:
            print(f"  → Error: {e}")
            continue

# --------- Export to Excel ----------
print(f"\n{'=' * 60}")
print(f"Total Products Scraped: {len(data)}")

if data:
    df = pd.DataFrame(data)
    # Remove duplicates based on Product URL
    df = df.drop_duplicates(subset=['Product URL'], keep='first')
    df.to_excel("chelsea_Side_Tables.xlsx", index=False)
    print(f"Unique Products: {len(df)}")
    print("Excel file created: chelsea_Side_Tables.xlsx")
else:
    print("No data scraped!")

print(f"{'=' * 60}")