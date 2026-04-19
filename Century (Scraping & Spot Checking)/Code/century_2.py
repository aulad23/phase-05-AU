# -*- coding: utf-8 -*-
# century_scraper_full_v9.py - Scraper with Com, Finish, Seat Height, Arm Height, List Price (clean, MSRP removed)

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
import re

# ================= CONFIG =================
INPUT_FILE = "century_Table_Lamps.xlsx"
OUTPUT_FILE = "century_Table_Lamps_final.xlsx"
# ==========================================

# Selenium driver
def create_driver():
    options = Options()
    options.headless = True
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(30)
    return driver


# Extract Product Family Id from Product Name
def extract_product_family_id(product_name):
    """
    'Andalusia Round- Dining Table'  → 'Andalusia Round'
    'Soho Bed- King'                 → 'Soho Bed'
    'Hampton Chair'                  → 'Hampton Chair' (no dash → full name)
    """
    if not product_name:
        return ""
    parts = re.split(r'\s*-\s+', product_name, maxsplit=1)
    return parts[0].strip()


# Extract dimensions + Seat Height + Arm Height + Length from accordion
def extract_dimensions_from_accordion(driver):
    dims = {
        "Width": "", "Depth": "", "Diameter": "", "Height": "",
        "Weight": "", "Seat": "", "Arm": "", "Length": ""
    }
    summary_elem = None
    try:
        summary_elem = driver.find_element(
            By.XPATH, "//summary[.//h2[contains(text(),'Dimensions')]]"
        )
        driver.execute_script("arguments[0].click();", summary_elem)
        time.sleep(0.5)

        content_div = summary_elem.find_element(
            By.XPATH, "./following-sibling::div[contains(@class,'accordion__content')]"
        )
        html = content_div.get_attribute("innerHTML")
        lines = re.split(r'<p>|</p>', html, flags=re.IGNORECASE)

        for line in lines:
            clean_line = re.sub(r'<.*?>', '', line).strip()
            if not clean_line:
                continue

            if m := re.search(r'(?:WEIGHT|OVERALL WEIGHT):?\s*([\d\.]+)', clean_line, re.IGNORECASE):
                dims["Weight"] = m.group(1)
            if m := re.search(r'(?:HEIGHT|OVERALL HEIGHT):?\s*([\d\.]+)', clean_line, re.IGNORECASE):
                dims["Height"] = m.group(1)
            if m := re.search(r'(?:WIDTH|OVERALL WIDTH):?\s*([\d\.]+)', clean_line, re.IGNORECASE):
                dims["Width"] = m.group(1)
            if m := re.search(r'(?:DEPTH|OVERALL DEPTH):?\s*([\d\.]+)', clean_line, re.IGNORECASE):
                dims["Depth"] = m.group(1)
            if m := re.search(r'DIAMETER:?\s*([\d\.]+)', clean_line, re.IGNORECASE):
                dims["Diameter"] = m.group(1)
            if m := re.search(r'(?:LENGTH|OVERALL LENGTH|\bL\b):?\s*([\d\.]+)', clean_line, re.IGNORECASE):
                dims["Length"] = m.group(1)
            if m := re.search(r'Seat Height:?\s*([\d\.]+\s*\w+)', clean_line, re.IGNORECASE):
                dims["Seat"] = m.group(1)
            if m := re.search(r'Arm Height:?\s*([\d\.]+\s*\w+)', clean_line, re.IGNORECASE):
                dims["Arm"] = m.group(1)

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"Error extracting dimensions: {e}")
    return dims, summary_elem


# Clean description extractor
def extract_clean_description(div_element):
    if not div_element:
        return ""
    try:
        html = div_element.get_attribute("innerHTML")
        parts = re.split(r'<br\s*/?>|</p>', html, flags=re.IGNORECASE)
        clean_lines = []
        for line in parts:
            text = re.sub(r'<.*?>', '', line).strip()
            if not text:
                continue
            if re.search(r'OVERALL\s+(HEIGHT|WIDTH|DEPTH)|HEIGHT|WIDTH|DEPTH|WEIGHT|DIAMETER', text, re.IGNORECASE):
                continue
            clean_lines.append(text)
        return " ".join(clean_lines)
    except Exception as e:
        print(f"Error extracting description: {e}")
        return ""


# Extract Finish
def extract_finish(div_element):
    if not div_element:
        return ""
    try:
        html = div_element.get_attribute("innerHTML")
        match = re.search(r'<br\s*/?>\s*Finish:\s*([^<\n\r]+)', html, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    except Exception as e:
        print(f"Error extracting Finish: {e}")
    return ""


# Extract COM
def extract_com(summary_elem):
    com = ""
    if not summary_elem:
        return com
    try:
        p_elements = summary_elem.find_elements(
            By.XPATH, "./following-sibling::div[contains(@class,'accordion__content')]//p"
        )
        for p in p_elements:
            text = p.text.strip()
            match = re.search(r'(?:COM|COM Fabric).*?([\d\.]+\s*\w+)', text, re.IGNORECASE)
            if match:
                com = match.group(1)
                break
    except Exception as e:
        print(f"Error extracting COM: {e}")
    return com


# List Price — removes MSRP, removes $ and USD, returns clean number like "3,597"
def extract_list_price(driver):
    price = ""
    try:
        price_elem = driver.find_element(By.CSS_SELECTOR, "div.price.price--large")
        html_text = price_elem.get_attribute("innerText").strip()

        # Remove any "MSRP $xxxx" portion
        clean_text = re.sub(r'MSRP\s*\$?\s*[\d,]+', '', html_text, flags=re.IGNORECASE)

        # Extract digits + commas only (e.g. 3,597 from "$3,597 USD")
        match = re.search(r'\$\s*([\d,]+(?:\.\d+)?)', clean_text)
        if match:
            price = match.group(1).strip()
        else:
            price = re.sub(r'[\$USDusd\s]', '', clean_text).strip()
    except Exception as e:
        print(f"Error extracting List Price: {e}")
    return price


# ================== MAIN ===================
def main():
    df_links = pd.read_excel(INPUT_FILE)
    driver = create_driver()
    final_data = []

    for idx, row in df_links.iterrows():
        product_link = row.get('Product URL', '')
        if not product_link:
            continue

        print(f"\n🔎 Scraping: {product_link}")
        try:
            driver.get(product_link)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(1)
        except TimeoutException:
            print("❌ Page did not load in time.")
            continue

        # Product Name (from input)
        product_name = row.get('Product Name', '')

        # Product Family Id
        product_family_id = extract_product_family_id(product_name)

        # SKU
        try:
            sku_elem = driver.find_element(By.CSS_SELECTOR, "div.hideAll span.sku")
            sku = sku_elem.text.strip()
        except NoSuchElementException:
            sku = ""

        # Description & Finish
        try:
            desc_div = driver.find_element(By.CSS_SELECTOR, "div.product__description.rte.quick-add-hidden")
            description = extract_clean_description(desc_div)
            finish = extract_finish(desc_div)
        except NoSuchElementException:
            description = ""
            finish = ""

        # Dimensions
        dims, summary_elem = extract_dimensions_from_accordion(driver)
        if not any(dims.values()):
            dims = {
                "Width": "", "Depth": "", "Diameter": "", "Height": "",
                "Weight": "", "Seat": "", "Arm": "", "Length": ""
            }

        # COM
        com = extract_com(summary_elem)

        # List Price
        list_price = extract_list_price(driver)

        # ✅ Final column order as requested
        product_data = {
            "Product URL":       product_link,
            "Image URL":         row.get("Image URL", ""),
            "Product Name":      product_name,
            "SKU":               sku,
            "Product Family Id": product_family_id,
            "Description":       description,
            "List Price":        list_price,
            "Weight":            dims["Weight"],
            "Width":             dims["Width"],
            "Depth":             dims["Depth"],
            "Diameter":          dims["Diameter"],
            "Length":            dims["Length"],
            "Height":            dims["Height"],
            "Finish":            finish,
            "Seat Height":       dims["Seat"],
            "Arm Height":        dims["Arm"]
        }

        final_data.append(product_data)

        print(f"✅ {product_name} | FamilyId:{product_family_id} | Price:{list_price} | SKU:{sku} | "
              f"H:{dims['Height']} W:{dims['Width']} D:{dims['Depth']} L:{dims['Length']} "
              f"Weight:{dims['Weight']} Dia:{dims['Diameter']} | "
              f"Finish:{finish} Seat Height:{dims['Seat']} Arm Height:{dims['Arm']}")

    driver.quit()

    # Save to Excel
    df_final = pd.DataFrame(final_data)
    df_final.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Scraping complete. File saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()