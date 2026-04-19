import time
import re
import sys
import pandas as pd
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
# chrome_options.add_argument("--headless=new")  # uncomment if you want headless
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 15)

# ----------------- CONFIG -----------------
excel_file = "check_Jaipur.xlsx"
output_file = "jaipur_rugs_list_step2.xlsx"

# ----------------- LOAD EXCEL -----------------
df = pd.read_excel(excel_file)

# Ensure these columns exist
columns_to_add = [
    "SKU", "Description", "Construction", "Material", "Pile Height", "Backing",
    "Country of Origin", "Style", "Size", "Content", "Design", "Shape"
]
for col in columns_to_add:
    if col not in df.columns:
        df[col] = ""

# ----------------- HELPERS -----------------
def clean_label(text: str) -> str:
    if not text:
        return ""
    t = re.sub(r"\s+", " ", text).strip().lower()
    t = t.replace(":", "")
    return t

def set_if_empty(idx, col, value):
    if value and (pd.isna(df.at[idx, col]) or str(df.at[idx, col]).strip() == ""):
        df.at[idx, col] = value

# Canonical mapping (left = page label variations, right = dataframe column)
label_mapping = {
    "sku": "SKU",
    "construction": "Construction",
    "content": "Material",           # site may use "Content" as Material
    "material": "Material",
    "pile height": "Pile Height",
    "pile_height_inches": "Pile Height",
    "pile height (inches)": "Pile Height",
    "thickness (inches)": "Pile Height",   # some pages use Thickness
    "backing": "Backing",
    "origin": "Country of Origin",
    "country of origin": "Country of Origin",
    "style": "Style",
    "size": "Size",
    "design": "Design",
    "shape": "Shape"
}

# Some pages render details as Bootstrap rows, some as table/dl, some in bullets.
DETAIL_SELECTORS = [
    "#product-attribute-specs-table .row",                       # main (Magento style)
    "#product-attribute-specs-table tr",                         # table rows
    ".product.attributes .row",                                  # alt rows
    ".product.attributes tr",                                    # alt table rows
    ".product.attributes dl",                                    # definition list container
]

def extract_kv_from_row(row_elem):
    """
    Try to read label/value from different row structures.
    Returns (label, value) or (None, None) if not found.
    """
    # 1) Magento style: .col.label with data-th-attribute + .col.data
    try:
        label_el = row_elem.find_element(By.CSS_SELECTOR, ".col.label")
        value_el = row_elem.find_element(By.CSS_SELECTOR, ".col.data")
        label_attr = label_el.get_attribute("data-th-attribute")
        label_text = label_attr if label_attr else label_el.text
        value_text = value_el.text
        if label_text.strip() and value_text.strip():
            return label_text.strip(), value_text.strip()
    except Exception:
        pass

    # 2) Simple table tr > th/td
    try:
        th = row_elem.find_element(By.CSS_SELECTOR, "th")
        td = row_elem.find_element(By.CSS_SELECTOR, "td")
        if th.text.strip() and td.text.strip():
            return th.text.strip(), td.text.strip()
    except Exception:
        pass

    # 3) Generic two-column divs
    try:
        cells = row_elem.find_elements(By.CSS_SELECTOR, "div, span")
        if len(cells) >= 2:
            left = cells[0].text.strip()
            right = cells[-1].text.strip()
            if left and right:
                return left, right
    except Exception:
        pass

    # 4) Definition lists (dl/dt/dd)
    try:
        dt = row_elem.find_element(By.CSS_SELECTOR, "dt")
        dd = row_elem.find_element(By.CSS_SELECTOR, "dd")
        if dt.text.strip() and dd.text.strip():
            return dt.text.strip(), dd.text.strip()
    except Exception:
        pass

    return None, None

def map_and_store(idx, raw_label, raw_value):
    label_key = clean_label(raw_label)
    value = raw_value.strip()

    # Normalize some keys
    if label_key in ("content",):
        mapped = "Material"
    else:
        mapped = label_mapping.get(label_key)

    # Try softer matching for common variants
    if not mapped:
        if "pile" in label_key and "height" in label_key:
            mapped = "Pile Height"
        elif "thickness" in label_key:
            mapped = "Pile Height"
        elif "country" in label_key and "origin" in label_key:
            mapped = "Country of Origin"
        elif "origin" == label_key:
            mapped = "Country of Origin"
        elif "material" in label_key:
            mapped = "Material"

    if not mapped:
        return  # unknown field; ignore

    # “Material → Content” mirror rule
    if mapped == "Material":
        set_if_empty(idx, "Material", value)
        set_if_empty(idx, "Content", value)
    else:
        set_if_empty(idx, mapped, value)

def grab_description():
    # Try multiple description selectors
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
        except Exception:
            continue
    # Fallback: meta description
    try:
        meta = driver.find_element(By.CSS_SELECTOR, 'meta[name="description"]')
        cnt = meta.get_attribute("content")
        if cnt and cnt.strip():
            return cnt.strip()
    except Exception:
        pass
    return ""

def grab_sku():
    # Common places to find SKU
    sku_selectors = [
        '[itemprop="sku"]',
        'div.product-info-main .product.attribute.sku .value',
        '.product-info-stock-sku .value',
        'span[data-ui-id="page-product-sku"]',
        '#product-attribute-specs-table [data-th-attribute="SKU"] + .col.data',
        '#product-attribute-specs-table tr:has(th:contains("SKU")) td',
    ]
    for sel in sku_selectors:
        try:
            # Selenium doesn't support :has or :contains natively; skip those two
            if ":has" in sel or ":contains" in sel:
                continue
            el = driver.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip() if el.text else el.get_attribute("content") or el.get_attribute("value") or ""
            if txt:
                return txt
        except Exception:
            continue
    # Fallback: look for visible text like "SKU: XXXX"
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        m = re.search(r"\bSKU[:\s]+([A-Za-z0-9\-_/\.]+)", body_text, flags=re.I)
        if m:
            return m.group(1).strip()
    except Exception:
        pass
    return ""

def parse_details(idx):
    # Try all container patterns
    found_any = False
    for container in DETAIL_SELECTORS:
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, container)
            for r in rows:
                label, value = extract_kv_from_row(r)
                if label and value:
                    map_and_store(idx, label, value)
                    found_any = True
        except Exception:
            continue
    return found_any

# ----------------- SCRAPING LOOP -----------------
for index, row in df.iterrows():
    product_url = row.get("Product URL", "")
    product_name = row.get("Product Name", "")
    print(f"\n--- Scraping ({index + 1}/{len(df)}) --- {product_name}")
    print(f"URL: {product_url}")

    if not product_url or str(product_url).strip().lower() in ("nan", "none"):
        print("No URL found, skipping...")
        continue

    try:
        driver.get(product_url)

        # Wait for page ready (title or product wrapper)
        try:
            wait.until(EC.any_of(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".product-info-main")),
                EC.presence_of_element_located((By.CSS_SELECTOR, "#maincontent")),
                EC.title_contains("Jaipur")
            ))
        except Exception:
            # still try after a short nap
            time.sleep(2)

        # Gentle scroll to trigger lazy blocks
        for y in (300, 800, 1400):
            driver.execute_script(f"window.scrollTo(0, {y});")
            time.sleep(0.6)

        # -------- Description --------
        try:
            description = grab_description()
            df.at[index, "Description"] = description
        except Exception:
            df.at[index, "Description"] = ""

        # -------- SKU (from multiple places) --------
        try:
            sku = grab_sku()
            if sku:
                df.at[index, "SKU"] = sku
        except Exception:
            pass

        # -------- Specs / Details (robust parsers + fallbacks) --------
        got_details = False
        try:
            got_details = parse_details(index)
        except Exception:
            pass

        # If we still don't have Material→Content mirrored, do it now if Material exists
        if str(df.at[index, "Material"]).strip():
            set_if_empty(index, "Content", str(df.at[index, "Material"]).strip())

        # Light normalization: unify “Pile Height” inches if they gave “Thickness”
        if str(df.at[index, "Pile Height"]).strip():
            # Keep as-is; but if it looks like "0.50" just store text unchanged
            pass

        if not got_details:
            print("No details table found or no mappable rows — continuing.")

    except Exception as e:
        print(f"Error scraping {product_url}: {e}")

# ----------------- SAVE UPDATED EXCEL -----------------
df.to_excel(output_file, index=False)
print(f"\nStep 2 scraping finished! Data saved to {output_file}")

driver.quit()
