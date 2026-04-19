import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup


def setup_driver():
    """Chrome driver setup with headless option"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def get_product_family_id(product_name):
    """
    Product Name theke Product Family Id derive kora.
    Rules:
      'Cascade Rug (Taupe)'               → 'Cascade Rug'
      'Cascade Rug, Taupe'                 → 'Cascade Rug'
      'Shadow 30" Wall Art - Midnight'     → 'Shadow 30" Wall Art'
      'Duval Grand Hugging Table - Deep Bronze' → 'Duval Grand Hugging Table'
      'Ornette Sofa'                       → 'Ornette Sofa'
    """
    if not product_name:
        return ''
    name = product_name.strip()

    if '(' in name:
        name = re.split(r'\s*\(', name)[0].strip()
        return name
    if ',' in name:
        name = name.split(',')[0].strip()
        return name
    if ' - ' in name:
        name = name.split(' - ')[0].strip()
        return name
    return name


def convert_fraction_to_decimal(value_str):
    """
    Fraction ke decimal e convert kora.
    Examples:
      '96 1/2'  → '96.5'
      '29 1/2'  → '29.5'
      '23-25'   → '23-25'  (range, keep as-is)
      '30'      → '30'
      '17 1/4'  → '17.25'
      '96 3/4'  → '96.75'
    """
    if not value_str:
        return value_str

    value_str = value_str.strip().replace('"', '').replace("'", '').strip()

    # Range like "23-25" → keep as-is
    if re.match(r'^\d+\s*-\s*\d+$', value_str):
        return value_str

    # Mixed fraction: "96 1/2" → 96.5
    mixed_match = re.match(r'^(\d+)\s+(\d+)/(\d+)$', value_str)
    if mixed_match:
        whole = int(mixed_match.group(1))
        numerator = int(mixed_match.group(2))
        denominator = int(mixed_match.group(3))
        if denominator != 0:
            result = whole + (numerator / denominator)
            # Clean: 96.5 not 96.50, but 96.25 stays
            if result == int(result):
                return str(int(result))
            return str(round(result, 2))

    # Simple fraction: "1/2" → 0.5
    frac_match = re.match(r'^(\d+)/(\d+)$', value_str)
    if frac_match:
        numerator = int(frac_match.group(1))
        denominator = int(frac_match.group(2))
        if denominator != 0:
            result = numerator / denominator
            return str(round(result, 2))

    # Plain number
    return value_str


def parse_dimensions(dim_string):
    """
    Dimension string parse kore alada columns e bhag kora.

    Input examples:
      '33"H x 96 1/2"W x 38"D; SD 23-25"; SH 19"; AH 30"'
      '29 1/2"H x 91"W x 35"D; SH 17 1/2"; AH 29 1/2"; SD 25"'
      '30" DIA'
      'MULTIPLE SIZES AVAILABLE'
      '20"H x 20"W x 1 1/2"D; Weight: 8 lbs'

    Output dict:
      Height, Width, Depth, Diameter, Seat Depth, Seat Height, Arm Height, Weight
    """
    result = {
        'Height': '',
        'Width': '',
        'Depth': '',
        'Diameter': '',
        'Seat Depth': '',
        'Seat Height': '',
        'Arm Height': '',
        'Weight': '',
    }

    if not dim_string or dim_string.strip() == '' or 'MULTIPLE' in dim_string.upper():
        return result

    # Normalize: replace '' (double single quote) with "
    dim = dim_string.strip().replace("''", '"')

    # ========== WEIGHT ==========
    # Match: "Weight: 8 lbs", "Weight: 10.5 kg", "8 IBS", "15 lbs", "8 IB"
    weight_match = re.search(
        r'(?:Weight\s*[:=]?\s*)?([\d.]+)\s*(?:lbs?|LBS?|IBS?|Ibs?|IB|kg|KG)',
        dim, re.IGNORECASE
    )
    if weight_match:
        result['Weight'] = weight_match.group(1)

    # ========== DIA (Diameter) - CHECK BEFORE D/Depth ==========
    # Match: '30" DIA', '30"DIA', 'DIA 30"', '30 DIA'
    dia_match = re.search(r'([\d]+(?:\s+\d+/\d+)?)\s*"?\s*DIA', dim, re.IGNORECASE)
    if not dia_match:
        dia_match = re.search(r'DIA\s*([\d]+(?:\s+\d+/\d+)?)\s*"?', dim, re.IGNORECASE)
    if dia_match:
        result['Diameter'] = convert_fraction_to_decimal(dia_match.group(1))

    # ========== HEIGHT (H) ==========
    # Match: '33"H', '29 1/2"H', '33" H'
    h_match = re.search(r'([\d]+(?:\s+\d+/\d+)?)\s*"?\s*H(?:\s|x|;|$)', dim)
    if h_match:
        result['Height'] = convert_fraction_to_decimal(h_match.group(1))

    # ========== WIDTH (W) ==========
    # Match: '96 1/2"W', '91"W'
    w_match = re.search(r'([\d]+(?:\s+\d+/\d+)?)\s*"?\s*W(?:\s|x|;|$)', dim)
    if w_match:
        result['Width'] = convert_fraction_to_decimal(w_match.group(1))

    # ========== DEPTH (D) - only if no DIA found for this value ==========
    # Match: '38"D', '35"D', '1 1/2"D'
    # Make sure we don't match DIA
    d_match = re.search(r'([\d]+(?:\s+\d+/\d+)?)\s*"?\s*D(?![Ii][Aa])(?:\s|x|;|$)', dim)
    if d_match:
        result['Depth'] = convert_fraction_to_decimal(d_match.group(1))

    # ========== SEAT DEPTH (SD) ==========
    # Match: 'SD 23-25"', 'SD 25"', 'SD: 23"'
    sd_match = re.search(r'SD\s*[:=]?\s*([\d]+(?:\s+\d+/\d+)?(?:\s*-\s*\d+(?:\s+\d+/\d+)?)?)\s*"?', dim)
    if sd_match:
        result['Seat Depth'] = convert_fraction_to_decimal(sd_match.group(1))

    # ========== SEAT HEIGHT (SH) ==========
    # Match: 'SH 19"', 'SH 17 1/2"'
    sh_match = re.search(r'SH\s*[:=]?\s*([\d]+(?:\s+\d+/\d+)?(?:\s*-\s*\d+(?:\s+\d+/\d+)?)?)\s*"?', dim)
    if sh_match:
        result['Seat Height'] = convert_fraction_to_decimal(sh_match.group(1))

    # ========== ARM HEIGHT (AH) ==========
    # Match: 'AH 30"', 'AH 29 1/2"'
    ah_match = re.search(r'AH\s*[:=]?\s*([\d]+(?:\s+\d+/\d+)?(?:\s*-\s*\d+(?:\s+\d+/\d+)?)?)\s*"?', dim)
    if ah_match:
        result['Arm Height'] = convert_fraction_to_decimal(ah_match.group(1))

    return result


def extract_list_price_from_page(soup):
    """Page theke MSRP / List Price extract"""
    price_info = soup.find('div', class_='product-info-price')
    if price_info:
        price_text = price_info.get_text()
        if 'Custom Price' in price_text:
            return 'Custom Price'
        msrp_match = re.search(r'Msrp\s*\$?([\d,]+(?:\.\d+)?)', price_text)
        if msrp_match:
            return msrp_match.group(1).replace(',', '')

    price_span = soup.find('span', class_='price')
    if price_span:
        price_text = price_span.get_text(strip=True)
        cleaned = price_text.replace('$', '').replace(',', '').strip()
        if cleaned and cleaned.replace('.', '').isdigit():
            return cleaned
    return ''


def extract_detail_page_data(driver, product_url):
    """Single product detail page theke data extract"""
    try:
        driver.get(product_url)
        time.sleep(3)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.product-info-main'))
            )
        except:
            pass

        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')

        # PRODUCT NAME
        page_product_name = ''
        h1_tag = soup.find('h1', class_='page-title')
        if h1_tag:
            span = h1_tag.find('span')
            if span:
                page_product_name = span.get_text(strip=True)

        # SKU
        page_sku = ''
        sku_value = soup.find('span', class_='ih-pdp__sku-value')
        if sku_value:
            page_sku = sku_value.get_text(strip=True)

        # LIST PRICE
        page_list_price = extract_list_price_from_page(soup)

        # DESCRIPTION
        description = ''
        desc_div = soup.find('div', class_='ih-pdp-description__content')
        if desc_div:
            description = desc_div.get_text(strip=True)

        # DIMENSIONS (raw string)
        dimensions = ''
        dim_div = soup.find('div', class_='ih-pdp-description__dimensions-value')
        if dim_div:
            dimensions = dim_div.get_text(strip=True)

        # ADDITIONAL ATTRIBUTES
        finish = ''
        content = ''
        attr_items = soup.find_all('div', class_='ih-pdp-additional-attributes__item')
        for attr_item in attr_items:
            label_div = attr_item.find('div', class_='ih-pdp-additional-attributes__item-label')
            value_div = attr_item.find('div', class_='ih-pdp-additional-attributes__item-value')
            if label_div and value_div:
                label_text = label_div.get_text(strip=True).lower()
                value_text = value_div.get_text(strip=True)
                if label_text == 'finish':
                    finish = value_text
                elif 'fabric content' in label_text or label_text == 'content':
                    content = value_text

        # SWATCH VARIATIONS
        swatch_options = soup.find_all('div', class_='swatch-option')
        variations = []

        if swatch_options and len(swatch_options) > 0:
            for swatch in swatch_options:
                var_data = {}
                var_data['SKU'] = swatch.get('option-sku', '')

                msrp = swatch.get('option-prophetmsrp', '')
                if msrp:
                    msrp = msrp.replace('$', '').replace(',', '').strip()
                var_data['List Price'] = msrp if msrp else page_list_price

                var_data['Color'] = swatch.get('option-label', '')

                var_finish = swatch.get('option-confinish', '')
                var_data['Finish'] = var_finish if var_finish else finish

                var_content = swatch.get('option-confabriccontent', '')
                var_data['Content'] = var_content if var_content else content

                var_dims = swatch.get('option-condimensions', '')
                if var_dims:
                    var_dims = var_dims.replace('@@', '"')
                var_data['Dimension'] = var_dims if var_dims else dimensions

                var_desc = swatch.get('option-desc', '')
                var_data['Description'] = var_desc if var_desc else description

                variations.append(var_data)

            # Click swatches for Image URL
            print(f"    Found {len(swatch_options)} color variations. Clicking swatches...")
            for i, swatch in enumerate(swatch_options):
                try:
                    option_id = swatch.get('option-id', '')
                    if not option_id:
                        variations[i]['Image URL'] = ''
                        continue

                    swatch_selector = f'div.swatch-option[option-id="{option_id}"]'
                    try:
                        swatch_elem = driver.find_element(By.CSS_SELECTOR, swatch_selector)
                        driver.execute_script("arguments[0].click();", swatch_elem)
                        time.sleep(2)

                        updated_soup = BeautifulSoup(driver.page_source, 'html.parser')
                        img_url = ''
                        for selector in [
                            'img.fotorama__img',
                            '.gallery-placeholder img',
                            '.product.media img',
                            '[data-gallery-role="gallery-placeholder"] img'
                        ]:
                            img_tag = updated_soup.select_one(selector)
                            if img_tag and img_tag.get('src'):
                                img_url = img_tag['src']
                                break

                        variations[i]['Image URL'] = img_url
                        print(f"      [{i+1}/{len(swatch_options)}] {variations[i].get('Color', '')} ✓")
                    except Exception as click_err:
                        print(f"      [{i+1}] Click error: {click_err}")
                        variations[i]['Image URL'] = swatch.get('option-tooltip-thumb', '')
                except Exception as e:
                    print(f"      [{i+1}] Error: {e}")
                    variations[i]['Image URL'] = ''

        base_data = {
            'Product Name (from page)': page_product_name,
            'SKU': page_sku,
            'List Price': page_list_price,
            'Description': description,
            'Dimension': dimensions,
            'Finish': finish,
            'Content': content,
        }

        return base_data, variations

    except Exception as e:
        print(f"    Error loading detail page: {e}")
        return {}, []


def main():
    """Main function - Step 2"""
    input_file = 'interlude_Dining_Tables.xlsx'
    output_file = 'interlude_Dining_Tables_final.xlsx'

    print("=" * 60)
    print("STEP 2: Detail Page Scraper (v3 - Dimension Parsing)")
    print("=" * 60)

    try:
        df = pd.read_excel(input_file, engine='openpyxl')
        print(f"\nLoaded {len(df)} products from {input_file}")
    except FileNotFoundError:
        print(f"\nError: {input_file} not found! Run Step 1 first.")
        return

    print(f"Columns: {list(df.columns)}\n")

    driver = setup_driver()
    all_products = []

    try:
        total = len(df)
        for idx, row in df.iterrows():
            product_url = str(row.get('Product URL', ''))
            original_image = str(row.get('Image URL', ''))
            product_name = str(row.get('Product Name', ''))
            original_sku = str(row.get('SKU', ''))
            original_price = str(row.get('List Price', ''))

            print(f"\n[{idx+1}/{total}] {product_name}")
            print(f"    URL: {product_url}")

            if not product_url or product_url == 'nan':
                print("    ⚠ Skipping - No URL")
                continue

            product_family_id = get_product_family_id(product_name)
            base_data, variations = extract_detail_page_data(driver, product_url)

            if variations:
                print(f"    → {len(variations)} variations")
                for var in variations:
                    dim_raw = var.get('Dimension', '') or base_data.get('Dimension', '')
                    dim_parsed = parse_dimensions(dim_raw)

                    product_row = {
                        'Product URL': product_url,
                        'Image URL': var.get('Image URL', '') or original_image,
                        'Product Name': product_name,
                        'SKU': var.get('SKU', '') or original_sku,
                        'Product Family Id': product_family_id,
                        'Description': var.get('Description', '') or base_data.get('Description', ''),
                        'Weight': dim_parsed['Weight'],
                        'Width': dim_parsed['Width'],
                        'Depth': dim_parsed['Depth'],
                        'Diameter': dim_parsed['Diameter'],
                        'Height': dim_parsed['Height'],
                        'Seat Depth': dim_parsed['Seat Depth'],
                        'Seat Height': dim_parsed['Seat Height'],
                        'Arm Height': dim_parsed['Arm Height'],
                        'List Price': var.get('List Price', '') or original_price,
                        'Finish': var.get('Finish', '') or base_data.get('Finish', ''),
                        'Content': var.get('Content', '') or base_data.get('Content', ''),
                        'Dimension': dim_raw,
                    }
                    all_products.append(product_row)
            else:
                print(f"    → Single product")
                dim_raw = base_data.get('Dimension', '')
                dim_parsed = parse_dimensions(dim_raw)

                product_row = {
                    'Product URL': product_url,
                    'Image URL': original_image,
                    'Product Name': product_name,
                    'SKU': base_data.get('SKU', '') or original_sku,
                    'Product Family Id': product_family_id,
                    'Description': base_data.get('Description', ''),
                    'Weight': dim_parsed['Weight'],
                    'Width': dim_parsed['Width'],
                    'Depth': dim_parsed['Depth'],
                    'Diameter': dim_parsed['Diameter'],
                    'Height': dim_parsed['Height'],
                    'Seat Depth': dim_parsed['Seat Depth'],
                    'Seat Height': dim_parsed['Seat Height'],
                    'Arm Height': dim_parsed['Arm Height'],
                    'List Price': base_data.get('List Price', '') or original_price,
                    'Finish': base_data.get('Finish', ''),
                    'Content': base_data.get('Content', ''),
                    'Dimension': dim_raw,
                }
                all_products.append(product_row)

            time.sleep(1)

    except Exception as e:
        print(f"\nFatal error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

    # ========== SAVE ==========
    if all_products:
        column_order = [
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
            'Seat Depth',
            'Seat Height',
            'Arm Height',
            'List Price',
            'Finish',
            'Content',
            'Dimension',
        ]
        result_df = pd.DataFrame(all_products, columns=column_order)
        result_df.to_excel(output_file, index=False, engine='openpyxl')

        print(f"\n{'=' * 60}")
        print(f"✅ DONE! Saved to: {output_file}")
        print(f"   Total rows: {len(result_df)}")
        print(f"   Unique Product Family Ids: {result_df['Product Family Id'].nunique()}")
        print(f"{'=' * 60}")

        # ---- Dimension parsing test print ----
        print("\n--- Dimension Parsing Samples ---")
        for i, (_, r) in enumerate(result_df.head(5).iterrows(), 1):
            print(f"\nRow {i}: {r['Product Name']}")
            print(f"  Raw Dimension: {r['Dimension']}")
            print(f"  Height: {r['Height']} | Width: {r['Width']} | Depth: {r['Depth']} | Diameter: {r['Diameter']}")
            print(f"  Seat Depth: {r['Seat Depth']} | Seat Height: {r['Seat Height']} | Arm Height: {r['Arm Height']}")
            print(f"  Weight: {r['Weight']}")
    else:
        print("\n❌ No products extracted!")


# ========== TEST DIMENSION PARSER ==========
if __name__ == "__main__":
    # Quick test before running
    print("--- Dimension Parser Test ---")
    test_cases = [
        '33"H x 96 1/2"W x 38"D; SD 23-25"; SH 19"; AH 30"',
        '29 1/2"H x 91"W x 35"D; SH 17 1/2"; AH 29 1/2"; SD 25"',
        '30" DIA',
        "24''H x 22''dia",
        '20"H x 15"dia',
        'MULTIPLE SIZES AVAILABLE',
        '20"H x 20"W x 1 1/2"D',
        '48"H x 24"W x 18"D; Weight: 8 lbs',
    ]

    for tc in test_cases:
        parsed = parse_dimensions(tc)
        print(f"\n  Input:    {tc}")
        print(f"  Height:   {parsed['Height']}")
        print(f"  Width:    {parsed['Width']}")
        print(f"  Depth:    {parsed['Depth']}")
        print(f"  Diameter: {parsed['Diameter']}")
        print(f"  SD:       {parsed['Seat Depth']}")
        print(f"  SH:       {parsed['Seat Height']}")
        print(f"  AH:       {parsed['Arm Height']}")
        print(f"  Weight:   {parsed['Weight']}")

    print("\n\n--- Starting Main Scraper ---\n")
    main()