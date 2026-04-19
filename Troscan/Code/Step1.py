import requests
from bs4 import BeautifulSoup
import pandas as pd

URL        = "https://www.troscandesign.com/products/seating/sofas"
BASE_URL   = "https://www.troscandesign.com"
VENDOR     = "Troscan"
CATEGORY   = "Sofas & Loveseats"
SKU_PREFIX = VENDOR[:3].upper() + "-" + CATEGORY[:2].upper()


def scrape_products():
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    response = requests.get(URL, headers=headers, timeout=15)
    response.raise_for_status()

    soup   = BeautifulSoup(response.text, "html.parser")
    thumbs = soup.find_all("div", class_="thumb")

    products = []
    for idx, thumb in enumerate(thumbs, start=1):
        name_tag     = thumb.find("div", class_="name")
        product_name = name_tag.get_text(strip=True) if name_tag else "N/A"

        deeplink    = thumb.get("deeplink", "").strip()
        product_url = f"{URL}#{deeplink}" if deeplink else URL

        img_tag   = thumb.find("img")
        img_src   = img_tag["src"] if img_tag and img_tag.get("src") else ""
        image_url = BASE_URL + img_src if img_src.startswith("/") else img_src

        sku = f"{SKU_PREFIX}-{str(idx).zfill(3)}"

        products.append({
            "Product URL":  product_url,
            "Image URL":    image_url,
            "Product Name": product_name,
            "SKU":          sku,
        })

    return products


if __name__ == "__main__":
    products = scrape_products()

    df = pd.DataFrame(products)
    df.to_excel("troscan_Sofas_Loveseats.xlsx", index=False)

    print(f"Done! {len(products)} products saved to troscan_Sofas_Loveseats.xlsx")