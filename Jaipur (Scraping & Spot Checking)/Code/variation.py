import time
import re
import os
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ----------------- SETUP -----------------
chrome_driver_path = "C:/chromedriver.exe"
service = Service(chrome_driver_path)
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
# chrome_options.add_argument("--headless=new")  # optional
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 15)

# ----------------- CONFIG -----------------
script_folder = os.path.dirname(os.path.abspath(__file__))
excel_file = os.path.join(script_folder, "jaipur_check.xlsx")
output_file = os.path.join(script_folder, "jaipur_check_final.xlsx")
max_retries = 3

# Load input Excel
df_input = pd.read_excel(excel_file)

# Load previous output if exists (resume support)
if os.path.exists(output_file):
    df_output = pd.read_excel(output_file)
else:
    df_output = pd.DataFrame(columns=[
        "Product URL", "Image URL", "Product Name", "SKU", "Product Family ID", "Description",
        "Construction", "Material", "Pile Height", "Backing", "Country of Origin",
        "Style", "Size", "Content", "Design", "Shape"
    ])

# ----------------- HELPERS -----------------
label_mapping = {
    "sku": "SKU",
    "construction": "Construction",
    "material": "Material",
    "content": "Material",
    "pile height": "Pile Height",
    "pile_height_inches": "Pile Height",
    "thickness (inches)": "Pile Height",
    "backing": "Backing",
    "origin": "Country of Origin",
    "country of origin": "Country of Origin",
    "style": "Style",
    "size": "Size",
    "design": "Design",
    "shape": "Shape"
}

DETAIL_SELECTORS = [
    "#product-attribute-specs-table .row",
    "#product-attribute-specs-table tr",
    ".product.attributes .row",
    ".product.attributes tr",
    ".product.attributes dl",
]

def clean_label(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip().lower().replace(":", "")

def extract_kv_from_row(row_elem):
    try:
        label_el = row_elem.find_element(By.CSS_SELECTOR, ".col.label")
        value_el = row_elem.find_element(By.CSS_SELECTOR, ".col.data")
        label_text = label_el.get_attribute("data-th-attribute") or label_el.text
        value_text = value_el.text
        if label_text.strip() and value_text.strip():
            return label_text.strip(), value_text.strip()
    except: pass
    try:
        th = row_elem.find_element(By.CSS_SELECTOR, "th")
        td = row_elem.find_element(By.CSS_SELECTOR, "td")
        if th.text.strip() and td.text.strip():
            return th.text.strip(), td.text.strip()
    except: pass
    try:
        dt = row_elem.find_element(By.CSS_SELECTOR, "dt")
        dd = row_elem.find_element(By.CSS_SELECTOR, "dd")
        if dt.text.strip() and dd.text.strip():
            return dt.text.strip(), dd.text.strip()
    except: pass
    return None, None

def grab_description():
    selectors = [
        "div.product.attribute.description > div.value",
        ".product.attribute.description .value",
        "#description .value",
        ".product.info.detailed .value",
    ]
    for sel in selectors:
        try:
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
            text = el.text.strip()
            if text:
                return text
        except: continue
    return ""

def grab_image_url():
    try:
        el = driver.find_element(By.CSS_SELECTOR, ".fotorama__stage__frame.fotorama__active img")
        src = el.get_attribute("src")
        if src and src.strip():
            return src
    except: pass
    try:
        el = driver.find_element(By.CSS_SELECTOR, ".product.media img")
        src = el.get_attribute("src")
        if src and src.strip():
            return src
    except: pass
    return ""

def grab_sku_from_specs():
    try:
        el = driver.find_element(By.CSS_SELECTOR, '.col.data[data-td-attribute="sku"]')
        txt = el.text.strip()
        if txt:
            return txt
    except: pass
    try:
        el = driver.find_element(By.CSS_SELECTOR, '#product-attribute-specs-table [data-th-attribute="sku"] + .col.data')
        txt = el.text.strip()
        if txt:
            return txt
    except: pass
    return ""

def wait_for_sku_update(prev_sku="", timeout=8):
    end = time.time() + timeout
    while time.time() < end:
        sku = grab_sku_from_specs()
        if sku and (not prev_sku or sku != prev_sku):
            return sku
        time.sleep(0.25)
    return grab_sku_from_specs()

def parse_details():
    result = {}
    for container in DETAIL_SELECTORS:
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, container)
            for r in rows:
                label, value = extract_kv_from_row(r)
                if label and value:
                    key = clean_label(label)
                    mapped = label_mapping.get(key)
                    if mapped:
                        result[mapped] = value.strip()
                    elif "pile" in key or "thickness" in key:
                        result["Pile Height"] = value.strip()
        except: continue
    return result

def get_family_id(name):
    return (name or "").split()[0].strip()

def requery_sizes():
    sw = driver.find_elements(By.CSS_SELECTOR, ".swatch-attribute.size .swatch-option.text")
    if sw:
        return "swatch", sw
    dds = driver.find_elements(By.CSS_SELECTOR, 'select.super-attribute-select')
    for sel in dds:
        name_id = (sel.get_attribute("name") or "") + " " + (sel.get_attribute("id") or "")
        if "size" in name_id.lower():
            opts = [o for o in sel.find_elements(By.TAG_NAME, "option")
                    if (o.get_attribute("value") or "").strip()
                    and o.text.strip().lower() not in ("choose an option", "select", "please select")]
            if opts:
                return "dropdown", (sel, opts)
    return "none", None

def wait_for_sizes(max_wait=8):
    end = time.time() + max_wait
    while time.time() < end:
        kind, elems = requery_sizes()
        if kind != "none":
            return kind, elems
        driver.execute_script("window.scrollBy(0, 300);")
        time.sleep(0.35)
    return "none", None

# ----------------- SCRAPING LOOP -----------------
print(f"\n🧭 Starting scraping for {len(df_input)} products...\n")

for index, row in tqdm(df_input.iterrows(), total=len(df_input), desc="Scraping Progress", unit="product"):
    product_url = str(row.get("Product URL", "") or "").strip()
    product_name = str(row.get("Product Name", "") or "").strip()
    if not product_url:
        continue

    # Skip if already scraped
    if not df_output[df_output["Product URL"] == product_url].empty:
        continue

    attempt = 0
    success = False
    while attempt < max_retries and not success:
        try:
            driver.get(product_url)
            time.sleep(2)

            family_id = get_family_id(product_name)
            prev_sku = grab_sku_from_specs()
            description = grab_description()
            details = parse_details()
            image_url = grab_image_url()

            # Default record
            default_record = {
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name,
                "SKU": prev_sku,
                "Product Family ID": family_id,
                "Description": description,
                "Construction": details.get("Construction", ""),
                "Material": details.get("Material", ""),
                "Pile Height": details.get("Pile Height", ""),
                "Backing": details.get("Backing", ""),
                "Country of Origin": details.get("Country of Origin", ""),
                "Style": details.get("Style", ""),
                "Size": details.get("Size", ""),
                "Content": details.get("Material", ""),
                "Design": details.get("Design", ""),
                "Shape": details.get("Shape", "")
            }
            df_output = pd.concat([df_output, pd.DataFrame([default_record])], ignore_index=True)
            df_output.to_excel(output_file, index=False)

            # ------------ VARIATIONS ------------
            color_blocks = driver.find_elements(By.CSS_SELECTOR, ".swatch-attribute.color .swatch-option.image")

            # If color exists
            if color_blocks:
                for c in color_blocks:
                    if "selected" in (c.get_attribute("class") or ""):
                        driver.execute_script("arguments[0].click();", c)
                        time.sleep(1.5)
                        break

                for color in color_blocks:
                    driver.execute_script("arguments[0].click();", color)
                    time.sleep(1.5)
                    kind, elems = wait_for_sizes()

                    if kind in ("swatch", "dropdown"):
                        size_elems = elems if kind == "swatch" else elems[1]
                        sel = None if kind == "swatch" else elems[0]

                        for item in size_elems:
                            size_label = item.text.strip()
                            if "custom" in size_label.lower():
                                continue
                            if kind == "swatch":
                                driver.execute_script("arguments[0].click();", item)
                            else:
                                driver.execute_script(
                                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                    sel, item.get_attribute("value")
                                )
                            time.sleep(0.8)

                            sku = wait_for_sku_update(prev_sku)
                            prev_sku = sku
                            description = grab_description()
                            details = parse_details()
                            image_url = grab_image_url()

                            record = {
                                "Product URL": product_url,
                                "Image URL": image_url,
                                "Product Name": product_name,
                                "SKU": sku,
                                "Product Family ID": family_id,
                                "Description": description,
                                "Construction": details.get("Construction", ""),
                                "Material": details.get("Material", ""),
                                "Pile Height": details.get("Pile Height", ""),
                                "Backing": details.get("Backing", ""),
                                "Country of Origin": details.get("Country of Origin", ""),
                                "Style": details.get("Style", ""),
                                "Size": size_label,
                                "Content": details.get("Material", ""),
                                "Design": details.get("Design", ""),
                                "Shape": details.get("Shape", "")
                            }
                            df_output = pd.concat([df_output, pd.DataFrame([record])], ignore_index=True)
                            df_output.to_excel(output_file, index=False)

            else:
                # If no color, handle only size variations
                kind, elems = wait_for_sizes()
                if kind in ("swatch", "dropdown"):
                    size_elems = elems if kind == "swatch" else elems[1]
                    sel = None if kind == "swatch" else elems[0]

                    for item in size_elems:
                        size_label = item.text.strip()
                        if "custom" in size_label.lower():
                            continue
                        if kind == "swatch":
                            driver.execute_script("arguments[0].click();", item)
                        else:
                            driver.execute_script(
                                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                sel, item.get_attribute("value")
                            )
                        time.sleep(0.8)

                        sku = wait_for_sku_update(prev_sku)
                        prev_sku = sku
                        description = grab_description()
                        details = parse_details()
                        image_url = grab_image_url()

                        record = {
                            "Product URL": product_url,
                            "Image URL": image_url,
                            "Product Name": product_name,
                            "SKU": sku,
                            "Product Family ID": family_id,
                            "Description": description,
                            "Construction": details.get("Construction", ""),
                            "Material": details.get("Material", ""),
                            "Pile Height": details.get("Pile Height", ""),
                            "Backing": details.get("Backing", ""),
                            "Country of Origin": details.get("Country of Origin", ""),
                            "Style": details.get("Style", ""),
                            "Size": size_label,
                            "Content": details.get("Material", ""),
                            "Design": details.get("Design", ""),
                            "Shape": details.get("Shape", "")
                        }
                        df_output = pd.concat([df_output, pd.DataFrame([record])], ignore_index=True)
                        df_output.to_excel(output_file, index=False)

            success = True

        except Exception as e:
            attempt += 1
            print(f"\n⚠️ Error for {product_name} (attempt {attempt}/{max_retries}): {e}")
            time.sleep(3)
            if attempt >= max_retries:
                print(f"❌ Failed to scrape: {product_name}. Skipping.")
                break

print(f"\n✅ All products & variations scraped successfully! Saved to: {output_file}")
driver.quit()
