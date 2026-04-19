"""
Collier Webb Scraper - FINAL VERSION
Correct HTML selectors based on actual website structure
"""

import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random
import re

def convert_kg_to_lbs(kg_value):
    """Convert KG to LBS (1 kg = 2.20462 lbs)"""
    try:
        return round(float(kg_value) * 2.20462, 0)
    except:
        return kg_value

def convert_mm_to_inches(mm_value):
    """Convert MM to Inches (1 inch = 25.4 mm)"""
    try:
        return round(float(mm_value) / 25.4, 0)
    except:
        return mm_value

def extract_dimension_value(text, keywords):
    """Extract numeric value from dimension text"""
    try:
        # Find the dimension line
        for keyword in keywords:
            if keyword.upper() in text.upper():
                # Extract the part after the keyword
                match = re.search(rf'{keyword}[:\s]*([0-9.]+)', text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return ''
    except:
        return ''

def parse_dimensions(dimension_text):
    """
    Parse dimension string and extract individual measurements
    Returns dict with Weight, Width, Height, Depth, Diameter, Length, Shade Details
    """
    result = {
        'Weight': '',
        'Width': '',
        'Height': '',
        'Depth': '',
        'Diameter': '',
        'Length': '',
        'Shade Details': ''
    }

    if not dimension_text or dimension_text == 'ERROR':
        return result

    # Split by | or newline
    lines = dimension_text.replace('|', '\n').split('\n')

    shade_details = []  # Collect all shade-related measurements

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Shade diameter/details - collect separately
        if 'SHADE' in line.upper():
            # Extract inch value from shade measurements
            inch_match = re.search(r'(\d+\.?\d*)"', line)
            if inch_match:
                shade_details.append(inch_match.group(1))
            else:
                # Try MM and convert
                mm_match = re.search(r'(\d+\.?\d*)\s*MM', line, re.IGNORECASE)
                if mm_match:
                    mm_value = float(mm_match.group(1))
                    shade_details.append(str(int(convert_mm_to_inches(mm_value))))
            continue

        # Weight: Extract LBS value, convert KG if needed
        if 'WEIGHT' in line.upper():
            # Check for LBS first
            lbs_match = re.search(r'(\d+\.?\d*)\s*LBS', line, re.IGNORECASE)
            if lbs_match:
                result['Weight'] = lbs_match.group(1)
            else:
                # Check for KG and convert
                kg_match = re.search(r'(\d+\.?\d*)\s*KG', line, re.IGNORECASE)
                if kg_match:
                    kg_value = float(kg_match.group(1))
                    result['Weight'] = str(int(convert_kg_to_lbs(kg_value)))

        # Width: Extract inches value, convert MM if needed
        elif 'WIDTH' in line.upper():
            # Check for inches first (")
            inch_match = re.search(r'(\d+\.?\d*)"', line)
            if inch_match:
                result['Width'] = inch_match.group(1)
            else:
                # Check for MM and convert
                mm_match = re.search(r'(\d+\.?\d*)\s*MM', line, re.IGNORECASE)
                if mm_match:
                    mm_value = float(mm_match.group(1))
                    result['Width'] = str(int(convert_mm_to_inches(mm_value)))

        # Height: Extract inches value, convert MM if needed
        elif 'HEIGHT' in line.upper():
            inch_match = re.search(r'(\d+\.?\d*)"', line)
            if inch_match:
                result['Height'] = inch_match.group(1)
            else:
                mm_match = re.search(r'(\d+\.?\d*)\s*MM', line, re.IGNORECASE)
                if mm_match:
                    mm_value = float(mm_match.group(1))
                    result['Height'] = str(int(convert_mm_to_inches(mm_value)))

        # Depth: Extract inches value, convert MM if needed
        elif 'DEPTH' in line.upper():
            inch_match = re.search(r'(\d+\.?\d*)"', line)
            if inch_match:
                result['Depth'] = inch_match.group(1)
            else:
                mm_match = re.search(r'(\d+\.?\d*)\s*MM', line, re.IGNORECASE)
                if mm_match:
                    mm_value = float(mm_match.group(1))
                    result['Depth'] = str(int(convert_mm_to_inches(mm_value)))

        # Diameter: Extract inches value, convert MM if needed
        elif 'DIAMETER' in line.upper():
            inch_match = re.search(r'(\d+\.?\d*)"', line)
            if inch_match:
                result['Diameter'] = inch_match.group(1)
            else:
                mm_match = re.search(r'(\d+\.?\d*)\s*MM', line, re.IGNORECASE)
                if mm_match:
                    mm_value = float(mm_match.group(1))
                    result['Diameter'] = str(int(convert_mm_to_inches(mm_value)))

        # Length: Extract inches value, convert MM if needed
        elif 'LENGTH' in line.upper():
            inch_match = re.search(r'(\d+\.?\d*)"', line)
            if inch_match:
                result['Length'] = inch_match.group(1)
            else:
                mm_match = re.search(r'(\d+\.?\d*)\s*MM', line, re.IGNORECASE)
                if mm_match:
                    mm_value = float(mm_match.group(1))
                    result['Length'] = str(int(convert_mm_to_inches(mm_value)))

    # Join shade details with comma
    if shade_details:
        result['Shade Details'] = ', '.join(shade_details)

    return result

def setup_driver():
    """Setup Chrome driver with Cloudflare bypass"""
    options = uc.ChromeOptions()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--window-size=1920,1080')

    print("🚀 Starting Chrome (Cloudflare bypass mode)...")
    driver = uc.Chrome(options=options, version_main=None)
    return driver

def wait_for_cloudflare(driver, max_wait=30):
    """Wait for Cloudflare check to complete"""
    print("  ⏳ Waiting for Cloudflare...")
    start_time = time.time()

    while time.time() - start_time < max_wait:
        try:
            page_source = driver.page_source.lower()
            if "cloudflare" in page_source and "checking" in page_source:
                time.sleep(1)
                continue
            else:
                print("  ✓ Page loaded!")
                return True
        except:
            pass
        time.sleep(1)

    return False

def get_product_details(driver, url):
    """Scrape product details with correct selectors"""
    try:
        print(f"  🌐 Loading: {url}")
        driver.get(url)

        # Wait for Cloudflare
        if not wait_for_cloudflare(driver):
            print("  ⚠ Cloudflare timeout")
            return create_error_result("Cloudflare timeout")

        # Extra wait for page load
        time.sleep(random.uniform(3, 5))

        result = {
            'SKU': '',
            'Product Family Id': '',
            'Description': '',
            'Dimension': '',  # Original full dimension text
            'Weight': '',
            'Width': '',
            'Height': '',
            'Depth': '',
            'Diameter': '',
            'Length': '',
            'Shade Details': '',
            'Finish': ''
        }

        # Extract SKU from form data-product-sku attribute
        try:
            # Find form with data-product-sku
            forms = driver.find_elements(By.CSS_SELECTOR, 'form[data-product-sku]')
            if forms:
                sku = forms[0].get_attribute('data-product-sku')
                if sku:
                    result['SKU'] = sku.strip()
                    print(f"  ✓ SKU: {result['SKU']}")
        except Exception as e:
            print(f"  ⚠ SKU error: {e}")

        # Product Family Id = Product Name (from input data)
        # This will be filled from Excel data

        # Extract Description - "Product Details" section with multiple fallbacks
        try:
            description_found = False

            # Method 1: Most specific - look for xx7avey class (the exact description container)
            try:
                # Target: .xx7avey .mgz-element-inner p span
                desc_elements = driver.find_elements(By.CSS_SELECTOR, '.xx7avey .mgz-element-inner p span')
                if desc_elements:
                    desc_text = desc_elements[0].text.strip()
                    if desc_text and len(desc_text) > 10:
                        result['Description'] = desc_text
                        description_found = True
                        print(f"  ✓ Description (Method 1): {desc_text[:50]}...")
            except:
                pass

            # Method 1b: Product Details heading + following paragraph
            if not description_found:
                try:
                    desc_xpath = "//h2[contains(text(), 'Product Details')]/following-sibling::div//p"
                    desc_elements = driver.find_elements(By.XPATH, desc_xpath)

                    if desc_elements:
                        desc_text = desc_elements[0].text.strip()
                        if desc_text and len(desc_text) > 10:
                            result['Description'] = desc_text
                            description_found = True
                            print(f"  ✓ Description (Method 1b): {desc_text[:50]}...")
                except:
                    pass

            # Method 2: Try class-based selector (.xx7avey is the description wrapper)
            if not description_found:
                try:
                    # First try to get span text inside p tag
                    desc_elements = driver.find_elements(By.CSS_SELECTOR, '.xx7avey p span')
                    if desc_elements:
                        desc_text = desc_elements[0].text.strip()
                        if desc_text and len(desc_text) > 10:
                            result['Description'] = desc_text
                            description_found = True
                            print(f"  ✓ Description (Method 2a): {desc_text[:50]}...")

                    # If span didn't work, try p tag directly
                    if not description_found:
                        desc_elements = driver.find_elements(By.CSS_SELECTOR, '.xx7avey p')
                        if desc_elements:
                            desc_text = desc_elements[0].text.strip()
                            if desc_text and len(desc_text) > 10:
                                result['Description'] = desc_text
                                description_found = True
                                print(f"  ✓ Description (Method 2b): {desc_text[:50]}...")
                except:
                    pass

            # Method 3: Look for any paragraph in product details area
            if not description_found:
                try:
                    desc_elements = driver.find_elements(By.CSS_SELECTOR, '.product-details-wrapper-row p')
                    for elem in desc_elements:
                        desc_text = elem.text.strip()
                        # Skip if it's dimensions or other info
                        if desc_text and len(desc_text) > 20 and not any(kw in desc_text for kw in ['Width:', 'Height:', 'DIMENSIONS', 'Care and Maintenance']):
                            result['Description'] = desc_text
                            description_found = True
                            print(f"  ✓ Description (Method 3): {desc_text[:50]}...")
                            break
                except:
                    pass

            # Method 4: Look in product-attribute-description div
            if not description_found:
                try:
                    desc_elements = driver.find_elements(By.CSS_SELECTOR, '.product.attribute.description p')
                    if desc_elements:
                        desc_text = desc_elements[0].text.strip()
                        if desc_text and len(desc_text) > 10:
                            result['Description'] = desc_text
                            description_found = True
                            print(f"  ✓ Description (Method 4): {desc_text[:50]}...")
                except:
                    pass

            # Method 5: Look for mgz-element-text class
            if not description_found:
                try:
                    desc_elements = driver.find_elements(By.CSS_SELECTOR, '.mgz-element-text p')
                    for elem in desc_elements:
                        desc_text = elem.text.strip()
                        if desc_text and len(desc_text) > 20 and 'The ' in desc_text:
                            result['Description'] = desc_text
                            description_found = True
                            print(f"  ✓ Description (Method 5): {desc_text[:50]}...")
                            break
                except:
                    pass

            # Method 6: Get from page-main-description if exists
            if not description_found:
                try:
                    desc_div = driver.find_elements(By.CSS_SELECTOR, '#description .value p')
                    if desc_div:
                        desc_text = desc_div[0].text.strip()
                        if desc_text and len(desc_text) > 10:
                            result['Description'] = desc_text
                            description_found = True
                            print(f"  ✓ Description (Method 6): {desc_text[:50]}...")
                except:
                    pass

            # Method 7: Get from meta description tag as last resort
            if not description_found:
                try:
                    meta_desc = driver.find_elements(By.CSS_SELECTOR, 'meta[name="description"]')
                    if meta_desc:
                        desc_text = meta_desc[0].get_attribute('content')
                        if desc_text and len(desc_text) > 10:
                            result['Description'] = desc_text.strip()
                            description_found = True
                            print(f"  ✓ Description (Method 7 - Meta): {desc_text[:50]}...")
                except:
                    pass

            if not description_found:
                print(f"  ⚠ No description found")
                result['Description'] = 'N/A'

        except Exception as e:
            print(f"  ⚠ Description error: {e}")
            result['Description'] = 'N/A'

        # Extract Dimensions from DIMENSIONS section
        try:
            dimensions = []
            dimension_found = False

            # Method 1: Find DIMENSIONS heading and get following text from p tags
            if not dimension_found:
                try:
                    dim_xpath = "//h4[contains(text(), 'DIMENSIONS')]/following-sibling::div//p"
                    dim_elements = driver.find_elements(By.XPATH, dim_xpath)

                    if dim_elements:
                        for elem in dim_elements:
                            text = elem.text.strip()
                            if text and any(kw in text.lower() for kw in ['width:', 'height:', 'depth:', 'weight:', 'diameter:', 'length:', 'shade']):
                                dimensions.append(text)
                                dimension_found = True
                except:
                    pass

            # Method 2: Try class-based selector for p tags in dimension wrapper
            if not dimension_found:
                try:
                    dim_elements = driver.find_elements(By.CSS_SELECTOR, '.dimension-text-wrapper p')
                    if dim_elements:
                        for elem in dim_elements:
                            text = elem.text.strip()
                            if text:
                                dimensions.append(text)
                                dimension_found = True
                except:
                    pass

            # Method 3: Look for ul/li tags (some products use list format)
            if not dimension_found:
                try:
                    # Try to find DIMENSIONS heading first, then get ul after it
                    dim_heading = driver.find_elements(By.XPATH, "//h4[contains(text(), 'DIMENSIONS')]")
                    if dim_heading:
                        # Get the ul that follows
                        ul_elements = driver.find_elements(By.XPATH, "//h4[contains(text(), 'DIMENSIONS')]/following-sibling::div//ul")
                        if ul_elements:
                            li_elements = ul_elements[0].find_elements(By.TAG_NAME, 'li')
                            for elem in li_elements:
                                text = elem.text.strip()
                                # Skip section headings (short text without colon and without mm/")
                                if text and not (len(text) < 15 and ':' not in text and 'mm' not in text.lower() and '"' not in text):
                                    dimensions.append(text)
                                    dimension_found = True
                except:
                    pass

            # Method 4: Look for li tags directly in dimension-text-wrapper
            if not dimension_found:
                try:
                    dim_elements = driver.find_elements(By.CSS_SELECTOR, '.dimension-text-wrapper li')
                    if dim_elements:
                        for elem in dim_elements:
                            text = elem.text.strip()
                            # Skip section headings
                            if text and not (len(text) < 15 and text[0].isupper() and ':' not in text and 'mm' not in text.lower() and '"' not in text):
                                dimensions.append(text)
                                dimension_found = True
                except:
                    pass

            # Method 5: Get all text from dimension wrapper and parse line by line
            if not dimension_found:
                try:
                    dim_wrapper = driver.find_elements(By.CSS_SELECTOR, '.dimension-text-wrapper')
                    if dim_wrapper:
                        full_text = dim_wrapper[0].text
                        # Split by newlines
                        for line in full_text.split('\n'):
                            line = line.strip()
                            # Include lines with measurements
                            if line and any(kw in line.lower() for kw in ['width:', 'height:', 'depth:', 'weight:', 'diameter:', 'length:', 'shade', 'mm/', '"']):
                                dimensions.append(line)
                                dimension_found = True
                except:
                    pass

            # Method 6: Look in any element with "ypju7i7" or similar dynamic classes (alternative dimension wrapper)
            if not dimension_found:
                try:
                    # Find elements by partial class match that might contain dimensions
                    all_divs = driver.find_elements(By.XPATH, "//div[contains(@class, 'dimension')]")
                    for div in all_divs:
                        text = div.text.strip()
                        if text and any(kw in text.lower() for kw in ['width:', 'height:', 'depth:', 'weight:', 'diameter:', 'shade']):
                            # Parse line by line
                            for line in text.split('\n'):
                                line = line.strip()
                                if line and any(kw in line.lower() for kw in ['width:', 'height:', 'depth:', 'weight:', 'diameter:', 'shade', 'mm/', '"']):
                                    if line not in dimensions:  # Avoid duplicates
                                        dimensions.append(line)
                                        dimension_found = True
                except:
                    pass

            result['Dimension'] = ' | '.join(dimensions)
            if result['Dimension']:
                print(f"  ✓ Dimensions found ({len(dimensions)} items)")
            else:
                print(f"  ⚠ No dimensions found")

            # Parse dimensions into individual fields (always run, even if empty)
            parsed_dims = parse_dimensions(result['Dimension'])
            result['Weight'] = parsed_dims['Weight']
            result['Width'] = parsed_dims['Width']
            result['Height'] = parsed_dims['Height']
            result['Depth'] = parsed_dims['Depth']
            result['Diameter'] = parsed_dims['Diameter']
            result['Length'] = parsed_dims['Length']
            result['Shade Details'] = parsed_dims['Shade Details']

            # Show parsed values if any
            if result['Dimension']:
                parsed_values = []
                if result['Weight']:
                    parsed_values.append(f"Weight:{result['Weight']}lbs")
                if result['Width']:
                    parsed_values.append(f"Width:{result['Width']}\"")
                if result['Height']:
                    parsed_values.append(f"Height:{result['Height']}\"")
                if result['Depth']:
                    parsed_values.append(f"Depth:{result['Depth']}\"")
                if result['Diameter']:
                    parsed_values.append(f"Diameter:{result['Diameter']}\"")
                if result['Length']:
                    parsed_values.append(f"Length:{result['Length']}\"")
                if result['Shade Details']:
                    parsed_values.append(f"Shade:{result['Shade Details']}\"")

                if parsed_values:
                    print(f"      → Parsed: {', '.join(parsed_values)}")
        except Exception as e:
            print(f"  ⚠ Dimension error: {e}")

        # Extract Finish options from select dropdown
        try:
            finishes = []

            # Find select with id amprot-swatch_588 or name options[588]
            select_selectors = [
                'select[id^="amprot-swatch"]',
                'select[name^="options"]',
                'select.amprot-swatch-input'
            ]

            for selector in select_selectors:
                try:
                    select_elem = driver.find_elements(By.CSS_SELECTOR, selector)
                    if select_elem:
                        options = select_elem[0].find_elements(By.TAG_NAME, 'option')
                        for option in options:
                            text = option.text.strip()
                            # Skip empty and placeholder options
                            if text and text not in ['', '--', '-- Please Select --']:
                                finishes.append(text)

                        if finishes:
                            break
                except:
                    continue

            # Alternative: Get from swatch names
            if not finishes:
                swatch_elements = driver.find_elements(By.CSS_SELECTOR, '.amprot-swatch-option .amprot-name')
                for elem in swatch_elements:
                    text = elem.text.strip()
                    if text:
                        finishes.append(text)

            # Remove duplicates while preserving order
            finishes = list(dict.fromkeys(finishes))
            result['Finish'] = ', '.join(finishes)

            if result['Finish']:
                print(f"  ✓ Found {len(finishes)} finish options")
        except Exception as e:
            print(f"  ⚠ Finish error: {e}")

        print(f"  ✅ Scraping successful!")
        return result

    except Exception as e:
        print(f"  ✗ Fatal error: {str(e)}")
        return create_error_result(str(e))

def create_error_result(error_msg):
    """Create error result"""
    return {
        'SKU': 'ERROR',
        'Product Family Id': 'ERROR',
        'Description': f'Error: {error_msg}',
        'Dimension': 'ERROR',
        'Weight': 'ERROR',
        'Width': 'ERROR',
        'Height': 'ERROR',
        'Depth': 'ERROR',
        'Diameter': 'ERROR',
        'Length': 'ERROR',
        'Shade Details': 'ERROR',
        'Finish': 'ERROR'
    }

def main():
    """Main scraping function"""
    print("=" * 80)
    print("Collier Webb Scraper - FINAL VERSION with Correct Selectors")
    print("=" * 80)

    # Read Excel file
    input_file = 'collierwebb_Backplates.xlsx'

    try:
        df = pd.read_excel(input_file)
        print(f"\n✓ Loaded Excel: {len(df)} products")
        print(f"Columns: {', '.join(df.columns.tolist())}")
    except FileNotFoundError:
        print(f"\n✗ Error: '{input_file}' not found!")
        return
    except Exception as e:
        print(f"\n✗ Error: {e}")
        return

    # Find columns (case-insensitive)
    url_col = None
    image_col = None
    name_col = None

    for col in df.columns:
        col_lower = col.lower()
        if 'url' in col_lower and 'image' not in col_lower:
            url_col = col
        elif 'image' in col_lower:
            image_col = col
        elif 'name' in col_lower:
            name_col = col

    if not url_col:
        print("\n✗ Error: Could not find URL column!")
        return

    print(f"\n✓ URL Column: '{url_col}'")
    if image_col:
        print(f"✓ Image Column: '{image_col}'")
    if name_col:
        print(f"✓ Name Column: '{name_col}'")

    # Setup Chrome driver
    print(f"\n{'=' * 80}")
    driver = setup_driver()
    print(f"✓ Browser ready!")
    print(f"{'=' * 80}\n")

    print("⚠ NOTE: Each product may take 5-10 seconds (Cloudflare check)")
    print("Please be patient!\n")

    print(f"{'=' * 80}")
    print(f"Starting to scrape {len(df)} products...")
    print(f"{'=' * 80}\n")

    results = []

    try:
        for idx, row in df.iterrows():
            url = row[url_col]

            if pd.isna(url) or not url:
                print(f"[{idx + 1}/{len(df)}] ⊘ Skipping: No URL\n")
                continue

            print(f"[{idx + 1}/{len(df)}] {url}")

            # Scrape product
            scraped = get_product_details(driver, url)

            # Get Product Name from input (this becomes Product Family Id)
            product_name = row.get(name_col, '') if name_col else ''

            # Product Family Id = Product Name
            if not scraped['Product Family Id'] or scraped['Product Family Id'] == 'ERROR':
                scraped['Product Family Id'] = product_name

            # Combine with existing data
            combined = {
                'Product URL': url,
                'Image URL': row.get(image_col, '') if image_col else '',
                'Product Name': product_name,
                'SKU': scraped['SKU'],
                'Product Family Id': scraped['Product Family Id'],
                'Description': scraped['Description'],
                'Weight': scraped['Weight'],
                'Width': scraped['Width'],
                'Depth': scraped['Depth'],
                'Diameter': scraped['Diameter'],
                'Height': scraped['Height'],
                'Length': scraped['Length'],
                'Shade Details': scraped['Shade Details'],
                'Finish': scraped['Finish'],
                'Dimension': scraped['Dimension']  # Keep original full text at end
            }

            results.append(combined)

            # Random delay
            if idx < len(df) - 1:
                delay = random.uniform(4, 8)
                print(f"  💤 Waiting {delay:.1f}s...\n")
                time.sleep(delay)

    except KeyboardInterrupt:
        print("\n\n⚠ Stopped by user!")

    finally:
        print("\n🔒 Closing browser...")
        driver.quit()

    if not results:
        print("\n✗ No data scraped!")
        return

    # Create output
    output_df = pd.DataFrame(results)

    # Column order as specified
    columns = ['Product URL', 'Image URL', 'Product Name', 'SKU',
               'Product Family Id', 'Description', 'Weight', 'Width',
               'Depth', 'Diameter', 'Height', 'Length', 'Shade Details',
               'Finish', 'Dimension']
    output_df = output_df[columns]

    # Save to Excel
    output_file = 'collierwebb_Backplates_Final.xlsx'
    output_df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"\n{'=' * 80}")
    print(f"✅ SCRAPING COMPLETE!")
    print(f"✅ Saved to: {output_file}")
    print(f"{'=' * 80}\n")

    # Summary
    total = len(results)
    successful = len([r for r in results if r['SKU'] != 'ERROR'])
    errors = total - successful

    print("📊 SUMMARY:")
    print(f"   Total products: {total}")
    print(f"   ✓ Successful: {successful}")
    print(f"   ✗ Errors: {errors}")

    if successful > 0:
        success_rate = (successful / total) * 100
        print(f"   Success rate: {success_rate:.1f}%")

    print(f"\n{'=' * 80}")
    print("Done! Check the output file for results.")
    print(f"{'=' * 80}\n")

if __name__ == "__main__":
    main()
