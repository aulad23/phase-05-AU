from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, NoSuchElementException
import pandas as pd
from time import sleep
import re


def safe_click(driver, element, max_attempts=3):
    """
    Safely click an element with retry logic
    """
    for attempt in range(max_attempts):
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except:
            try:
                element.click()
                return True
            except ElementClickInterceptedException:
                if attempt < max_attempts - 1:
                    sleep(0.5)
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                          element)
                    sleep(0.5)
                else:
                    return False
    return False


def parse_dimension(dimension_str):
    """
    Parse dimension string and extract Width, Depth, Diameter, Height
    Priority: First check for Diameter, then parse W/D/H
    """
    result = {
        'width': '',
        'depth': '',
        'diameter': '',
        'height': ''
    }

    if not dimension_str or dimension_str in ['Not Available', 'Error', '']:
        return result

    # Clean the dimension string
    dim = dimension_str.strip()

    # First, check for Diameter (priority)
    diameter_patterns = [
        r'(\d+\.?\d*)\s*(?:DIA|Dia|diam|DIAM)\b',
        r'Diameter[:\s]*(\d+\.?\d*)',
        r'(?:^|[^\w])(\d+\.?\d*)(?:DIA|Dia|diam)(?:[^\w]|$)'
    ]

    for pattern in diameter_patterns:
        match = re.search(pattern, dim, re.IGNORECASE)
        if match:
            result['diameter'] = match.group(1).strip()
            # Remove diameter part from string
            dim = re.sub(pattern, '', dim, flags=re.IGNORECASE)
            break

    # Parse Width, Depth, Height
    # Pattern 1: Standard format with W, D, H
    pattern1 = r'(\d+\.?\d*)\s*["\']?\s*W\s*[xX×]\s*(\d+\.?\d*)\s*["\']?\s*D\s*[xX×]\s*(\d+\.?\d*)\s*["\']?\s*H'
    match = re.search(pattern1, dim, re.IGNORECASE)
    if match:
        result['width'] = match.group(1).strip()
        result['depth'] = match.group(2).strip()
        result['height'] = match.group(3).strip()
        return result

    # Pattern 2: Without the last H
    pattern2 = r'(\d+\.?\d*)\s*["\']?\s*W\s*[xX×]\s*(\d+\.?\d*)\s*["\']?\s*D\s*[xX×]\s*(\d+\.?\d*)\s*["\']?'
    match = re.search(pattern2, dim, re.IGNORECASE)
    if match:
        result['width'] = match.group(1).strip()
        result['depth'] = match.group(2).strip()
        result['height'] = match.group(3).strip()
        return result

    # Pattern 3: Individual extraction
    width_match = re.search(r'(\d+\.?\d*)\s*["\']?\s*W\b', dim, re.IGNORECASE)
    if width_match:
        result['width'] = width_match.group(1).strip()

    # Only extract depth if diameter wasn't found
    if not result['diameter']:
        depth_match = re.search(r'(\d+\.?\d*)\s*["\']?\s*D\b', dim, re.IGNORECASE)
        if depth_match:
            result['depth'] = depth_match.group(1).strip()

    height_match = re.search(r'(\d+\.?\d*)\s*["\']?\s*H\b', dim, re.IGNORECASE)
    if height_match:
        result['height'] = height_match.group(1).strip()

    return result


def extract_product_details(driver, product_url):
    """
    Extract detailed information from a product page
    """
    try:
        print(f"\n  Opening: {product_url}")
        driver.get(product_url)
        sleep(3)

        driver.execute_script("window.scrollTo(0, 0);")
        sleep(0.5)

        product_data = {
            'sku': '',
            'product_family_id': '',
            'description': '',
            'details': '',
            'color': '',
            'weight': '',
            'dimension': '',
            'variations': []
        }

        # Extract SKU
        try:
            sku_element = driver.find_element(By.CSS_SELECTOR, '.woolentor_product_sku_info .sku')
            product_data['sku'] = sku_element.text.strip()
            print(f"    ✓ SKU: {product_data['sku']}")
        except:
            print("    ⚠ SKU not found")

        # Extract Description
        try:
            desc_element = driver.find_element(By.CSS_SELECTOR, '.woocommerce_product_description p')
            product_data['description'] = desc_element.text.strip()
            print(f"    ✓ Description extracted ({len(product_data['description'])} chars)")
        except:
            print("    ⚠ Description not found")

        # Extract Color
        try:
            color_element = driver.find_element(By.XPATH, "//div[contains(text(), 'Color:')]")
            color_text = color_element.text.strip()
            product_data['color'] = color_text.replace('Color:', '').strip()
            print(f"    ✓ Color (Method 1): {product_data['color']}")
        except:
            try:
                color_element = driver.find_element(By.XPATH, "//*[contains(text(), 'Color:')]")
                color_text = color_element.text.strip()
                product_data['color'] = re.sub(r'.*Color:\s*', '', color_text).strip()
                print(f"    ✓ Color (Method 2): {product_data['color']}")
            except:
                print("    ⚠ Color not found (will try in Details)")

        # Extract Weight
        try:
            weight_element = driver.find_element(By.XPATH,
                                                 "//p[contains(text(), 'lbs') or contains(text(), 'lb') or contains(text(), 'Weight')]")
            weight_text = weight_element.text.strip()

            weight_patterns = [
                r'Weight[:\s]+(\d+\.?\d*\s*lbs?)',
                r':\s*(\d+\.?\d*\s*lbs?)',
                r'(\d+\.?\d*\s*lbs?)',
            ]

            for pattern in weight_patterns:
                match = re.search(pattern, weight_text, re.IGNORECASE)
                if match:
                    product_data['weight'] = match.group(1).strip()
                    print(f"    ✓ Weight: {product_data['weight']}")
                    break
        except:
            print("    ⚠ Weight not found (will try in Details)")

        # Click on Details accordion
        details_opened = False
        try:
            details_accordion = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'summary[data-accordion-index="1"]'))
            )

            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                  details_accordion)
            sleep(1)

            if safe_click(driver, details_accordion):
                details_opened = True
                sleep(1.5)

                # Extract Details text
                try:
                    details_list = driver.find_elements(By.CSS_SELECTOR, '#e-n-accordion-item-7760 li')
                    details_text = []
                    for li in details_list:
                        text = li.text.strip()
                        if text:
                            details_text.append(text)
                    product_data['details'] = ' | '.join(details_text)
                    print(f"    ✓ Details extracted ({len(details_text)} points)")
                except:
                    print("    ⚠ Details content not found")

                # Extract Color from Details if not found
                if not product_data['color']:
                    try:
                        color_in_details = driver.find_element(By.CSS_SELECTOR,
                                                               '#e-n-accordion-item-7760 [data-id="d67df19"]')
                        color_text = color_in_details.text.strip()
                        if 'Color:' in color_text:
                            product_data['color'] = color_text.replace('Color:', '').strip()
                            print(f"    ✓ Color found in Details: {product_data['color']}")
                    except:
                        pass

                # Extract Weight from Details if not found
                if not product_data['weight']:
                    try:
                        weight_in_details = driver.find_element(By.CSS_SELECTOR,
                                                                '#e-n-accordion-item-7760 [data-id="9018af3"] p')
                        weight_text = weight_in_details.text.strip()
                        weight_match = re.search(r'(\d+\.?\d*\s*lbs?)', weight_text, re.IGNORECASE)
                        if weight_match:
                            product_data['weight'] = weight_match.group(1)
                            print(f"    ✓ Weight found in Details: {product_data['weight']}")
                    except:
                        pass
            else:
                print("    ⚠ Could not click Details accordion")

        except Exception as e:
            print(f"    ⚠ Details accordion error: {e}")

        # Click on Dimensions accordion
        dimensions_found = False
        try:
            if details_opened:
                try:
                    safe_click(driver, details_accordion)
                    sleep(0.5)
                except:
                    pass

            dimensions_accordion = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'summary[data-accordion-index="2"]'))
            )

            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                  dimensions_accordion)
            sleep(1)

            if safe_click(driver, dimensions_accordion):
                sleep(2)  # Give more time for content to load

                # Try multiple methods to find the table
                try:
                    # Method 1: Direct table search in accordion
                    table_selectors = [
                        '#e-n-accordion-item-7761 table tbody tr',
                        '#e-n-accordion-item-7761 tbody tr',
                        '.e-n-accordion-item[open] table tbody tr',
                        'details[open] table tbody tr'
                    ]

                    table_rows = []
                    for selector in table_selectors:
                        try:
                            rows = driver.find_elements(By.CSS_SELECTOR, selector)
                            if rows and len(rows) > 0:
                                table_rows = rows
                                print(f"    ℹ Found table using selector: {selector}")
                                break
                        except:
                            continue

                    if table_rows and len(table_rows) > 0:
                        print(f"    ✓ Found {len(table_rows)} variations in dimensions table")
                        dimensions_found = True

                        for row_idx, row in enumerate(table_rows):
                            try:
                                cells = row.find_elements(By.TAG_NAME, 'td')

                                # Debug: Print number of cells
                                if row_idx == 0:
                                    print(f"    ℹ Table has {len(cells)} columns per row")

                                if len(cells) >= 3:
                                    variation_sku = cells[0].text.strip()
                                    variation_dimension = cells[2].text.strip()
                                    variation_weight = ''

                                    # Skip if SKU is empty
                                    if not variation_sku:
                                        continue

                                    # Try to find weight for this SKU
                                    if variation_sku and details_opened:
                                        try:
                                            safe_click(driver, dimensions_accordion)
                                            sleep(0.5)
                                            safe_click(driver, details_accordion)
                                            sleep(0.5)

                                            weight_xpath = f"//*[contains(text(), '{variation_sku}')]"
                                            sku_elements = driver.find_elements(By.XPATH, weight_xpath)

                                            for elem in sku_elements:
                                                elem_text = elem.text.strip()
                                                if 'lbs' in elem_text.lower() or 'lb' in elem_text.lower():
                                                    weight_match = re.search(r'(\d+\.?\d*\s*lbs?)', elem_text,
                                                                             re.IGNORECASE)
                                                    if weight_match:
                                                        variation_weight = weight_match.group(1)
                                                        break

                                            safe_click(driver, details_accordion)
                                            sleep(0.5)
                                            safe_click(driver, dimensions_accordion)
                                            sleep(0.5)
                                        except:
                                            pass

                                    variation = {
                                        'sku': variation_sku,
                                        'dimension': variation_dimension,
                                        'weight': variation_weight
                                    }

                                    product_data['variations'].append(variation)
                                    print(f"      - {variation['sku']}: {variation['dimension']}")
                                elif len(cells) >= 2:
                                    # Some tables might have only 2 columns
                                    variation_sku = cells[0].text.strip()
                                    variation_dimension = cells[1].text.strip()

                                    if variation_sku:
                                        variation = {
                                            'sku': variation_sku,
                                            'dimension': variation_dimension,
                                            'weight': ''
                                        }
                                        product_data['variations'].append(variation)
                                        print(f"      - {variation['sku']}: {variation['dimension']}")
                            except Exception as e:
                                print(f"    ⚠ Error parsing row {row_idx}: {e}")
                                continue
                    else:
                        print("    ⚠ No table rows found")
                except Exception as e:
                    print(f"    ⚠ Could not extract table: {e}")

                # If no table found, try text format
                if not dimensions_found:
                    try:
                        dimension_container = driver.find_element(By.CSS_SELECTOR, '#e-n-accordion-item-7761')
                        dimension_text = dimension_container.text

                        if dimension_text and len(dimension_text.strip()) > 10:
                            print(f"    ℹ Dimension text length: {len(dimension_text)} chars")

                            dimension_patterns = [
                                r'(\d+\.?\d*\s*["\']?\s*W\s*[xX×]\s*\d+\.?\d*\s*["\']?\s*D\s*[xX×]\s*\d+\.?\d*\s*["\']?\s*H)',
                                r'(\d+\.?\d*\s*["\']?\s*W\s*[xX×]\s*\d+\.?\d*\s*["\']?\s*D)',
                                r'Dimensions?:\s*([^\n]+)',
                            ]

                            for pattern in dimension_patterns:
                                match = re.search(pattern, dimension_text, re.IGNORECASE)
                                if match:
                                    product_data['dimension'] = match.group(1).strip()
                                    dimensions_found = True
                                    print(f"    ✓ Dimension (text): {product_data['dimension']}")
                                    break
                    except Exception as e:
                        print(f"    ⚠ Error extracting text dimensions: {e}")
            else:
                print("    ⚠ Could not click Dimensions accordion")

        except Exception as e:
            print(f"    ⚠ Dimensions accordion error: {str(e)[:150]}")

        if not dimensions_found and not product_data['variations']:
            print("    ℹ No dimensions found for this product")

        return product_data

    except Exception as e:
        print(f"    ✗ Major error: {e}")
        import traceback
        traceback.print_exc()
        return None


def scrape_product_details_from_excel(excel_file='bedroom_products.xlsx'):
    """
    Read products from Excel and scrape detailed information
    """
    df_products = pd.read_excel(excel_file)
    print(f"Found {len(df_products)} products in {excel_file}")

    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()

    all_detailed_data = []

    try:
        for index, row in df_products.iterrows():
            product_url = row['Product URL']
            product_name = row['Product Name']
            image_url = row['Image URL']

            print(f"\n{'=' * 60}")
            print(f"[{index + 1}/{len(df_products)}] Processing: {product_name}")
            print(f"{'=' * 60}")

            product_details = extract_product_details(driver, product_url)

            if product_details:
                if product_details['variations']:
                    for variation in product_details['variations']:
                        # Parse dimension into W/D/Dia/H
                        dim_parsed = parse_dimension(variation['dimension'])

                        detailed_row = {
                            'Product URL': product_url,
                            'Image URL': image_url,
                            'Product Name': product_name,
                            'SKU': variation['sku'],
                            'Product Family ID': product_name,
                            'Description': product_details['description'],
                            'Weight': variation['weight'] if variation['weight'] else product_details['weight'],
                            'Width': dim_parsed['width'],
                            'Depth': dim_parsed['depth'],
                            'Diameter': dim_parsed['diameter'],
                            'Height': dim_parsed['height'],
                            'Color': product_details['color'],
                            'Dimension': variation['dimension']
                        }
                        all_detailed_data.append(detailed_row)
                    print(f"\n  ✓ Added {len(product_details['variations'])} variations")
                else:
                    # Parse base dimension
                    dim_parsed = parse_dimension(product_details['dimension'])

                    detailed_row = {
                        'Product URL': product_url,
                        'Image URL': image_url,
                        'Product Name': product_name,
                        'SKU': product_details['sku'],
                        'Product Family ID': product_name,
                        'Description': product_details['description'],
                        'Weight': product_details['weight'],
                        'Width': dim_parsed['width'],
                        'Depth': dim_parsed['depth'],
                        'Diameter': dim_parsed['diameter'],
                        'Height': dim_parsed['height'],
                        'Color': product_details['color'],
                        'Dimension': product_details['dimension'] if product_details['dimension'] else 'Not Available'
                    }
                    all_detailed_data.append(detailed_row)
                    print(f"\n  ✓ Added base product")
            else:
                detailed_row = {
                    'Product URL': product_url,
                    'Image URL': image_url,
                    'Product Name': product_name,
                    'SKU': '',
                    'Product Family ID': product_name,
                    'Description': '',
                    'Weight': '',
                    'Width': '',
                    'Depth': '',
                    'Diameter': '',
                    'Height': '',
                    'Color': '',
                    'Dimension': 'Error'
                }
                all_detailed_data.append(detailed_row)
                print(f"\n  ⚠ Error occurred")

            sleep(2)

    finally:
        driver.quit()

    return all_detailed_data


def save_detailed_data(data, filename='bedroom_Sectionals_detailed.xlsx'):
    """
    Save detailed product data to Excel
    """
    if not data:
        print("No data to save!")
        return

    df = pd.DataFrame(data)

    # Final column order
    column_order = [
        'Product URL',
        'Image URL',
        'Product Name',
        'SKU',
        'Product Family ID',
        'Description',
        'Weight',
        'Width',
        'Depth',
        'Diameter',
        'Height',
        'Color',
        'Dimension'
    ]

    df = df[column_order]
    df.to_excel(filename, index=False, engine='openpyxl')

    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"✓ File saved: {filename}")
    print(f"✓ Total rows: {len(df)}")
    print(f"✓ Unique products: {df['Product Family ID'].nunique()}")
    print(f"✓ Rows with Color: {len(df[df['Color'] != ''])}")
    print(f"✓ Rows with Weight: {len(df[df['Weight'] != ''])}")
    print(f"✓ Rows with Width: {len(df[df['Width'] != ''])}")
    print(f"✓ Rows with Depth: {len(df[df['Depth'] != ''])}")
    print(f"✓ Rows with Diameter: {len(df[df['Diameter'] != ''])}")
    print(f"✓ Rows with Height: {len(df[df['Height'] != ''])}")


if __name__ == "__main__":
    print("=" * 60)
    print("Step 2: Extracting Product Details")
    print("=" * 60)

    detailed_data = scrape_product_details_from_excel('bedroom_Sectionals.xlsx')

    if detailed_data:
        save_detailed_data(detailed_data)
    else:
        print("No data was scraped.")