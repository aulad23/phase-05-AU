import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


def scrape_stahlandband_products(url):
    """Scrape product data from stahlandband.com"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        products = []

        grid_items = soup.find_all('div', class_=lambda x: x and 'countergrid' in x)

        for item in grid_items:
            try:
                link_tag = item.find('a', class_='box')
                product_url = link_tag['href'] if link_tag else ''

                img_tag = item.find('img')
                image_url = img_tag['src'] if img_tag else ''

                name_tag = item.find('h3')
                product_name = name_tag.get_text(strip=True) if name_tag else ''

                products.append({
                    'Product URL': product_url,
                    'Image URL': image_url,
                    'Product Name': product_name,
                })

            except Exception as e:
                print(f"Error processing item: {e}")
                continue

        return products

    except Exception as e:
        print(f"Error fetching URL {url}: {e}")
        return []


def save_to_excel(products, output_file):
    """Save scraped products to Excel file"""
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Products"

    # Add headers
    headers = ['Product URL', 'Image URL', 'Product Name', 'SKU']
    sheet.append(headers)

    # Add data rows
    for product in products:
        sheet.append([
            product['Product URL'],
            product['Image URL'],
            product['Product Name'],
            product['SKU']
        ])

    wb.save(output_file)
    print(f"Data saved to {output_file}")
    print(f"Total products scraped: {len(products)}")


def main():
    # Multiple URLs to scrape
    urls = [
        "https://stahlandband.com/collections/other/hardware/",
        #"https://stahlandband.com/collections/other/painting//"
    ]

    all_products = []
    product_index = 1

    # Scrape each URL
    for url in urls:
        print(f"\nScraping from: {url}")
        products = scrape_stahlandband_products(url)

        # Add SKU to each product
        for product in products:
            product['SKU'] = f"STA-PU-{str(product_index).zfill(3)}"
            product_index += 1

        all_products.extend(products)
        print(f"Found {len(products)} products")

    if all_products:
        # Save to Excel
        output_file = "stahlandband_Pulls.xlsx"
        save_to_excel(all_products, output_file)

        # Display preview
        print("\n" + "=" * 60)
        print("PREVIEW - First 3 products:")
        print("=" * 60)
        for i, product in enumerate(all_products[:3], 1):
            print(f"\n{i}. {product['Product Name']}")
            print(f"   SKU: {product['SKU']}")
            print(f"   URL: {product['Product URL']}")
    else:
        print("No products found. Check the URLs or website structure.")


if __name__ == "__main__":
    main()