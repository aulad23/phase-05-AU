from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time
import re


def setup_driver():
    """Chrome driver setup"""
    chrome_options = Options()
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return driver


def extract_product_family_id(product_name):
    """Product Name থেকে variant part বাদ দিয়ে Product Family ID তৈরি করবে"""
    if '–' in product_name:
        return product_name.split('–')[0].strip()
    elif ' - ' in product_name:
        return product_name.split(' - ')[0].strip()
    return product_name.strip()


def extract_wattage(details_text):
    """Max Wattage থেকে শুধু wattage number extract করবে (100W)"""
    match = re.search(r'Max Wattage:\s*(\d+W)', details_text, re.IGNORECASE)
    if match:
        return match.group(1)
    return ""


def clean_dimension_value(value):
    """
    Dimension value থেকে শুধু number বের করবে
    17″ → 17
    29 1/4″ → 29 1/4
    4 1/2″ Dia. → 4 1/2
    """
    if not value:
        return ""
    # ″, ′, Dia., dia. remove করবে, শুধু number এবং fraction রাখবে
    cleaned = re.sub(r'[″′]', '', value).strip()
    cleaned = re.sub(r'\s*Dia\.?', '', cleaned, flags=re.IGNORECASE).strip()
    return cleaned


def extract_shade_details(notes_text):
    """
    Shade Details থেকে শুধু shade code এবং dimensions নিবে
    Input: "Shade: S-23 21 1/4″ x 23 1/4″ x 19 1/4″H Ivory Silk / White Silk/ Linen Shade is not included."
    Output: "S-23 21 1/4″ x 23 1/4″ x 19 1/4″H"
    """
    if 'Shade:' not in notes_text:
        return ""

    shade_start = notes_text.find('Shade:')
    shade_section = notes_text[shade_start + 6:].strip()

    lines = shade_section.split('\n')
    if not lines:
        return ""

    first_line = lines[0].strip()

    patterns_to_remove = [
        r'\s*Ivory Silk.*',
        r'\s*White Silk.*',
        r'\s*Linen.*',
        r'\s*Shade is not included.*',
        r'\s*Estimated lead-time.*'
    ]

    result = first_line
    for pattern in patterns_to_remove:
        result = re.sub(pattern, '', result, flags=re.IGNORECASE)

    return result.strip()


def scrape_product_details_selenium(url, driver):
    """Selenium দিয়ে product details scrape করবে"""
    try:
        print(f"  🌐 Loading page...")
        driver.get(url)
        time.sleep(3)

        soup = BeautifulSoup(driver.page_source, "lxml")

        details = {
            'Description': '',
            'Dimension': '',
            'Wattage': '',
            'Finish': '',
            'Shade Details': '',
            'Weight': '',
            'Width': '',
            'Depth': '',
            'Diameter': '',
            'Height': '',
            'Canopy': ''  # নতুন field
        }

        # Details section থেকে Wattage extract
        details_section = soup.find('h2', string='Details:')
        if details_section:
            details_text_elem = details_section.find_next('div', class_='elementor-widget-container')
            if details_text_elem:
                details_text = details_text_elem.get_text(strip=True)
                details['Wattage'] = extract_wattage(details_text)
                if details['Wattage']:
                    print(f"  ✓ Wattage: {details['Wattage']}")

        # Finish extract
        finish_section = soup.find('h2', string='Finish Shown:')
        if finish_section:
            finish_elem = finish_section.find_next('div', class_='elementor-widget-container')
            if finish_elem:
                details['Finish'] = finish_elem.get_text(strip=True)
                if details['Finish']:
                    print(f"  ✓ Finish: {details['Finish']}")

        # Standard Dimensions - পুরো text Dimension column এ রাখবে
        dimensions_section = soup.find('h2', string='Standard Dimensions:')
        if dimensions_section:
            dim_elem = dimensions_section.find_next('div', class_='elementor-widget-container')
            if dim_elem:
                dim_text = dim_elem.get_text().strip()

                # পুরো dimension text Dimension column এ রাখবে
                details['Dimension'] = dim_text
                print(f"  ✓ Dimension: {dim_text}")

                # এখন আলাদা আলাদা করে extract করবে
                width_match = re.search(r'Width:\s*([\d\s/″′\-]+)', dim_text, re.IGNORECASE)
                if width_match:
                    details['Width'] = clean_dimension_value(width_match.group(1))

                height_match = re.search(r'Height:\s*([\d\s/″′\-]+)', dim_text, re.IGNORECASE)
                if height_match:
                    details['Height'] = clean_dimension_value(height_match.group(1))

                depth_match = re.search(r'Depth:\s*([\d\s/″′\-]+)', dim_text, re.IGNORECASE)
                if depth_match:
                    details['Depth'] = clean_dimension_value(depth_match.group(1))

                diameter_match = re.search(r'Diameter:\s*([\d\s/″′\-]+)', dim_text, re.IGNORECASE)
                if diameter_match:
                    details['Diameter'] = clean_dimension_value(diameter_match.group(1))

                weight_match = re.search(r'Weight:\s*([\d\s\.,lbskg]+)', dim_text, re.IGNORECASE)
                if weight_match:
                    details['Weight'] = clean_dimension_value(weight_match.group(1))

                # Canopy extract - নতুন
                canopy_match = re.search(r'Canopy:\s*([\d\s/″′\-]+(?:\s*Dia\.?)?)', dim_text, re.IGNORECASE)
                if canopy_match:
                    details['Canopy'] = clean_dimension_value(canopy_match.group(1))
                    print(f"  ✓ Canopy: {details['Canopy']}")

                print(f"  ✓ Individual dimensions extracted")

        # Shade Details - শুধু shade code এবং dimensions
        notes_section = soup.find('h2', string='Notes:')
        if notes_section:
            notes_elem = notes_section.find_next('div', class_='elementor-widget-container')
            if notes_elem:
                notes_text = notes_elem.get_text()
                shade_details = extract_shade_details(notes_text)
                if shade_details:
                    details['Shade Details'] = shade_details
                    print(f"  ✓ Shade Details: {shade_details}")

        return details

    except Exception as e:
        print(f"  ❌ Error: {str(e)}")
        return {
            'Description': '',
            'Dimension': '',
            'Wattage': '',
            'Finish': '',
            'Shade Details': '',
            'Weight': '',
            'Width': '',
            'Depth': '',
            'Diameter': '',
            'Height': '',
            'Canopy': ''
        }


# Main execution
print("📂 Reading Excel file...")
df = pd.read_excel("palmerhargrave_products.xlsx")

# নতুন columns যোগ করা
new_columns = ['Product Family Id', 'Description', 'Dimension', 'Wattage',
               'Finish', 'Shade Details', 'Weight', 'Width', 'Depth',
               'Diameter', 'Height', 'Canopy']

for col in new_columns:
    df[col] = ''

# Product Family Id তৈরি করা
df['Product Family Id'] = df['Product Name'].apply(extract_product_family_id)

print(f"🔍 Starting Selenium browser...")
driver = setup_driver()

try:
    print(f"🚀 Scraping details for {len(df)} products...\n")

    # প্রতিটি product এর জন্য details scrape করা
    for index, row in df.iterrows():
        product_url = row['Product URL']
        product_name = row['Product Name']

        print(f"{'=' * 60}")
        print(f"[{index + 1}/{len(df)}] {product_name}")
        print(f"URL: {product_url}")

        details = scrape_product_details_selenium(product_url, driver)

        # DataFrame update করা
        for key, value in details.items():
            df.at[index, key] = value

        print(f"✅ Completed")

        # Rate limiting
        wait_time = 2
        print(f"⏳ Waiting {wait_time} seconds...\n")
        time.sleep(wait_time)

finally:
    # Browser বন্ধ করা
    print("\n🔒 Closing browser...")
    driver.quit()

# Final column order - আপনার specified order অনুযায়ী
final_columns = [
    'Product URL',
    'Image URL',
    'Product Name',
    'SKU',
    'Product Family Id',
    'Description',
    'Weight',
    'Width',
    'Depth',
    'Diameter',
    'Height',
    'Canopy',
    'Wattage',
    'Finish',
    'Shade Details',
    'Dimension'
]

df = df[final_columns]

# Save করা
output_file = "palmerhargrave_products_detailed.xlsx"
df.to_excel(output_file, index=False)

print(f"\n{'=' * 60}")
print(f"✅ সব product এর details scrape করা হয়েছে!")
print(f"📊 Output file: {output_file}")
print(f"📈 Total products: {len(df)}")