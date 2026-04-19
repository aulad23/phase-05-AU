from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import re
import json

# Configure Chrome options
chrome_options = Options()
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")


def extract_product_family_id(product_name):
    """
    Extract Product Family ID from Product Name
    Example: "Zeppelin Wall, Light" -> "Zeppelin Wall"
    """
    if ',' in product_name:
        return product_name.split(',')[0].strip()
    return product_name


def extract_description(driver):
    """
    Extract description from product page
    """
    try:
        description_texts = []

        # Approach 1: Look for paragraphs in description sections
        desc_elements = driver.find_elements(By.XPATH,
                                             "//div[contains(@class, 'tw-text-base') or contains(@class, 'description')]/p")
        if desc_elements:
            for elem in desc_elements:
                text = elem.text.strip()
                if text:
                    description_texts.append(text)

        # Approach 2: Look for article or main content area
        if not description_texts:
            main_content = driver.find_elements(By.XPATH, "//main//p | //article//p")
            if main_content:
                for elem in main_content[:3]:
                    text = elem.text.strip()
                    if text and len(text) > 20:
                        description_texts.append(text)

        if description_texts:
            return " ".join(description_texts).strip()[:500]
    except:
        pass

    return "N/A"


def extract_specifications(driver):
    """
    Extract Weight, Width, Depth, Diameter, Height from product page.
    Tries multiple common patterns used on lighting/product sites.
    """
    specs = {
        'Weight': '',
        'Width': '',
        'Depth': '',
        'Diameter': '',
        'Height': ''
    }

    try:
        # Strategy 1: Look for table rows with label + value pattern
        rows = driver.find_elements(By.XPATH, "//table//tr | //dl//div | //ul[contains(@class,'spec')]//li")
        for row in rows:
            text = row.text.strip().lower()
            for key in specs:
                if key.lower() in text:
                    cells = row.find_elements(By.XPATH, ".//td | .//dd | .//span")
                    if len(cells) >= 2:
                        specs[key] = cells[-1].text.strip()
                    else:
                        match = re.search(
                            rf'{key.lower()}\s*[:\-]?\s*([\d.,]+\s*(?:kg|g|cm|mm|in|")?)',
                            text, re.IGNORECASE
                        )
                        if match:
                            specs[key] = match.group(1).strip()

        # Strategy 2: Search entire page text with regex (fallback)
        if all(v == 'N/A' for v in specs.values()):
            page_text = driver.find_element(By.TAG_NAME, 'body').text
            for key in specs:
                match = re.search(
                    rf'{key}\s*[:\-]?\s*([\d.,]+\s*(?:kg|g|cm|mm|in|")?)',
                    page_text, re.IGNORECASE
                )
                if match:
                    specs[key] = match.group(1).strip()

        # Strategy 3: Look for definition lists (dl/dt/dd pattern)
        if all(v == 'N/A' for v in specs.values()):
            dt_elements = driver.find_elements(By.TAG_NAME, 'dt')
            for dt in dt_elements:
                label = dt.text.strip().lower()
                for key in specs:
                    if key.lower() in label:
                        try:
                            dd = dt.find_element(By.XPATH, "following-sibling::dd[1]")
                            specs[key] = dd.text.strip()
                        except:
                            pass

    except Exception as e:
        print(f"    ⚠ Error extracting specs: {e}")

    return specs


def extract_finishes(driver):
    """
    Extract all available finishes — returns FULL NAMES.
    e.g. "Antique Brass, Brass Polished Lacquered, Bronze, Nickel Matte"
    """
    finishes_list = []

    # ── METHOD 1: Parse Alpine x-data JSON  →  use "title" field ─────────────
    try:
        xdata_containers = driver.find_elements(
            By.XPATH, "//*[contains(@x-data, '\"finishes\"')]"
        )
        for container in xdata_containers:
            xdata_raw = container.get_attribute('x-data')
            if not xdata_raw or '"finishes"' not in xdata_raw:
                continue
            match = re.search(r'"finishes"\s*:\s*(\[.*?\])', xdata_raw, re.DOTALL)
            if match:
                json_str = match.group(1).replace('&quot;', '"')
                for finish in json.loads(json_str):
                    title = finish.get('title', finish.get('sku', ''))
                    if title:
                        finishes_list.append(title)
            if finishes_list:
                print(f"    ✓ [Method 1] {len(finishes_list)} finishes from x-data JSON")
                break
    except Exception as e:
        print(f"    ⚠ Method 1 failed: {e}")

    # ── METHOD 2: JS-click the "+" button → overlay opens → read finish names ─
    if not finishes_list:
        try:
            clicked = driver.execute_script("""
                var els = document.querySelectorAll('*');
                for (var i = 0; i < els.length; i++) {
                    var attr = els[i].getAttribute('x-on:click.prevent');
                    if (attr && attr.trim() === 'overlay = true') {
                        els[i].click();
                        return true;
                    }
                }
                return false;
            """)

            if clicked:
                time.sleep(2)
                print("    → Overlay opened via JS click (Method 2)")

                name_divs = driver.find_elements(
                    By.CSS_SELECTOR,
                    "div.tw-text-xs.tw-uppercase.tw-font-sans-book.tw-text-center.tw-tracking-widest"
                )
                for div in name_divs:
                    title = div.text.strip()
                    if title:
                        finishes_list.append(title)

                driver.execute_script("""
                    var els = document.querySelectorAll('*');
                    for (var i = 0; i < els.length; i++) {
                        var attr = els[i].getAttribute('x-on:click.prevent');
                        if (attr && attr.trim() === 'overlay = false') {
                            els[i].click();
                            break;
                        }
                    }
                """)
                time.sleep(0.5)

                if finishes_list:
                    print(f"    ✓ [Method 2] {len(finishes_list)} finishes from overlay")

        except Exception as e:
            print(f"    ⚠ Method 2 failed: {e}")

    # ── METHOD 3: Read hidden DOM (no overlay click needed) ───────────────────
    if not finishes_list:
        try:
            name_divs = driver.find_elements(
                By.CSS_SELECTOR,
                "div.tw-text-xs.tw-uppercase.tw-font-sans-book.tw-text-center.tw-tracking-widest"
            )
            for div in name_divs:
                title = div.get_attribute('textContent').strip()
                if title:
                    finishes_list.append(title)

            if finishes_list:
                print(f"    ✓ [Method 3] {len(finishes_list)} finishes from hidden DOM")
        except Exception as e:
            print(f"    ⚠ Method 3 failed: {e}")

    if not finishes_list:
        return ""

    return ", ".join(finishes_list)


def extract_tearsheet_link(driver):
    """
    Extract the tearsheet PDF download link from product page.
    Looks for an <a> tag inside the tw-flex wrapper that links to a .pdf tearsheet.
    Example: <a href="https://cdn2.assets-servd.host/.../tearsheets/PC15S-Newport-Picture-Light-Small.pdf" ...>
    """
    try:
        # Strategy 1: Find anchor inside the tearsheet flex wrapper (most specific)
        tearsheet_anchors = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'tw-flex') and contains(@class,'tw-flex-wrap')]"
            "//a[contains(@href, 'tearsheets') and contains(@href, '.pdf')]"
        )
        if tearsheet_anchors:
            link = tearsheet_anchors[0].get_attribute('href')
            if link:
                print(f"    ✓ [Tearsheet Method 1] Found: {link}")
                return link

        # Strategy 2: Any anchor with tearsheets PDF href on the page
        tearsheet_anchors = driver.find_elements(
            By.XPATH,
            "//a[contains(@href, 'tearsheets') and contains(@href, '.pdf')]"
        )
        if tearsheet_anchors:
            link = tearsheet_anchors[0].get_attribute('href')
            if link:
                print(f"    ✓ [Tearsheet Method 2] Found: {link}")
                return link

        # Strategy 3: Any anchor whose text contains "Tearsheet" or "tearsheet"
        tearsheet_anchors = driver.find_elements(
            By.XPATH,
            "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tearsheet')]"
        )
        if tearsheet_anchors:
            link = tearsheet_anchors[0].get_attribute('href')
            if link:
                print(f"    ✓ [Tearsheet Method 3] Found: {link}")
                return link

    except Exception as e:
        print(f"    ⚠ Error extracting tearsheet link: {e}")

    print("    ⚠ No tearsheet link found")
    return ""


def scrape_product_details(driver, product_url, product_name, image_url, sku):
    """
    Scrape detailed information from a product URL
    """
    print(f"  📍 Scraping details...")

    try:
        driver.get(product_url)
        time.sleep(4)

        product_family_id = extract_product_family_id(product_name)
        description = extract_description(driver)
        specs = extract_specifications(driver)
        finishes = extract_finishes(driver)
        tearsheet_link = extract_tearsheet_link(driver)

        print(f"    ✓ Family ID: {product_family_id}")
        print(f"    ✓ Description: {description[:40]}...")
        print(f"    ✓ Height: {specs['Height']}")
        print(f"    ✓ Finishes: {finishes}")
        print(f"    ✓ Tearsheet Link: {tearsheet_link}")

        return {
            'Product URL': product_url,
            'Image URL': image_url,
            'Product Name': product_name,
            'Product Family ID': product_family_id,
            'Description': description,
            'SKU': sku,
            'Weight': specs['Weight'],
            'Width': specs['Width'],
            'Depth': specs['Depth'],
            'Diameter': specs['Diameter'],
            'Height': specs['Height'],
            'Finish': finishes,
            'Tearsheet Link': tearsheet_link
        }

    except Exception as e:
        print(f"    ✗ Error: {e}")
        return {
            'Product URL': product_url,
            'Image URL': image_url,
            'Product Name': product_name,
            'Product Family ID': extract_product_family_id(product_name),
            'Description': 'N/A',
            'SKU': sku,
            'Weight': 'N/A',
            'Width': 'N/A',
            'Depth': 'N/A',
            'Diameter': 'N/A',
            'Height': 'N/A',
            'Finish': 'N/A',
            'Tearsheet Link': ''
        }


def main():
    input_file = 'hector_finch_Table_Lamps.xlsx'
    driver = None
    all_final_data = []

    try:
        print("📖 Reading Excel file...")
        df = pd.read_excel(input_file)
        print(f"✓ Found {len(df)} products\n")

        print("🚀 Starting Chrome driver...")
        driver = webdriver.Chrome(options=chrome_options)

        batch_size = 10
        total_batches = (len(df) + batch_size - 1) // batch_size

        for batch_num in range(total_batches):
            start_idx = batch_num * batch_size
            end_idx = min(start_idx + batch_size, len(df))

            print(f"\n{'=' * 70}")
            print(f"📦 BATCH {batch_num + 1}/{total_batches} (Products {start_idx + 1}-{end_idx})")
            print(f"{'=' * 70}")

            for idx in range(start_idx, end_idx):
                row = df.iloc[idx]
                product_url = row['Product URL']
                image_url = row['Image URL']
                product_name = row['Product Name']
                sku = row['SKU']

                print(f"\n  [{idx + 1}/{len(df)}] {product_name}")

                try:
                    details = scrape_product_details(driver, product_url, product_name, image_url, sku)
                    all_final_data.append(details)
                except Exception as e:
                    print(f"    ✗ Error processing: {e}")
                    continue

            if batch_num < total_batches - 1:
                print(f"\n⏳ Waiting 5 seconds before next batch...")
                time.sleep(5)

        final_df = pd.DataFrame(all_final_data)

        columns_order = ['Product URL', 'Image URL', 'Product Name', 'Product Family ID',
                         'Description', 'SKU', 'Weight', 'Width', 'Depth', 'Diameter',
                         'Height', 'Finish', 'Tearsheet Link']
        final_df = final_df[columns_order]

        output_file = 'hector_finch_Table_Lamps_final_.xlsx'
        final_df.to_excel(output_file, index=False)

        print(f"\n{'=' * 70}")
        print(f"✅ SUCCESS - FINAL OUTPUT CREATED!")
        print(f"{'=' * 70}")
        print(f"📊 Total products processed: {len(all_final_data)}")
        print(f"💾 File saved: {output_file}")
        print(f"\n📈 Data preview (first 5 rows):")
        print(final_df.head())

    except FileNotFoundError:
        print(f"❌ Error: File '{input_file}' not found!")
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        if driver:
            print("\n🔌 Closing browser...")
            driver.quit()


if __name__ == "__main__":
    main()