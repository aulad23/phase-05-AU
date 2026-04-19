import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

INPUT_FILE = "Chelsea_Wall_Covering.xlsx"
OUTPUT_FILE = "Chelsea_Wall_Covering_FINAL.xlsx"

df = pd.read_excel(INPUT_FILE)
final_rows = []

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def txt_num(token: str) -> str:
    t = (token or "").replace("″", "").replace('"', "").strip()
    m = re.search(r"(\d+(?:\.\d+)?)", t)
    return m.group(1) if m else ""

def pick_visible_size_from_dd(dd_el):
    spans = dd_el.find_elements(By.CSS_SELECTOR, "span.c-attrlist__unitdisplay")
    if not spans:
        return norm_space(dd_el.text)

    for sp in spans:
        style = (sp.get_attribute("style") or "").lower().replace(" ", "")
        if "display:none" not in style:
            return norm_space(sp.text)

    return norm_space(spans[0].text)

def clean_weight_value(val: str) -> str:
    if not val:
        return ""
    v = val.strip()
    v = re.sub(r"(?i)\b(lbs|lb|ib)\b", "", v)  # remove units
    v = re.sub(r"\s+", " ", v).strip()
    return v

def extract_details_size_weight(details_div):
    pairs = []
    size_val = ""
    weight_val = ""
    content_val = ""
    repeat_val = ""

    dls = details_div.find_elements(By.CSS_SELECTOR, "dl")
    for dl in dls:
        dts = dl.find_elements(By.TAG_NAME, "dt")
        dds = dl.find_elements(By.TAG_NAME, "dd")
        if not dts or not dds:
            continue

        key = norm_space(dts[0].text)
        if not key:
            continue

        key_l = key.lower()
        dd = dds[0]

        if key_l == "size":
            val = pick_visible_size_from_dd(dd)
            size_val = val
            if val:
                pairs.append(f"{key}: {val}")

        elif key_l == "weight":
            val = norm_space(dd.text)
            weight_val = clean_weight_value(val)
            if val:
                pairs.append(f"{key}: {val}")

        elif key_l == "fabric content":
            val = norm_space(dd.text)
            content_val = val
            if val:
                pairs.append(f"{key}: {val}")

        elif key_l == "pattern repeat":
            val = norm_space(dd.text)
            repeat_val = val
            if val:
                pairs.append(f"{key}: {val}")

        else:
            val = norm_space(dd.text)
            if val:
                pairs.append(f"{key}: {val}")

    return " | ".join(pairs), size_val, weight_val, content_val, repeat_val

def extract_size_from_dropdown(driver):
    """
    Fallback:
    select[name="sizeOptions"] option[data-disabled="true"]
    2nd option (index=1) text will be used.
    Example text:
      "GUS170D — H53.9 × W57.1 × D79.9″"
    We will extract the part after dash (—) if exists.
    """
    try:
        opts = driver.find_elements(By.CSS_SELECTOR, 'select[name="sizeOptions"] option[data-disabled="true"]')
        if len(opts) >= 2:
            raw = norm_space(opts[1].text)  # ✅ 2nd option
            # keep only size part after —
            if "—" in raw:
                raw = norm_space(raw.split("—", 1)[1])
            return raw
        elif len(opts) == 1:
            raw = norm_space(opts[0].text)
            if "—" in raw:
                raw = norm_space(raw.split("—", 1)[1])
            return raw
    except:
        pass
    return ""

def extract_list_price(driver):
    """
    Extract price from: div.o-grid --2cols@s > button span.current
    Returns only the numeric value (e.g., "122")
    """
    try:
        price_elem = driver.find_element(By.CSS_SELECTOR, "div.o-grid.--2cols\\@s button span.current")
        price_text = norm_space(price_elem.text)
        # Extract numeric value, removing $ and any other characters
        match = re.search(r"(\d+(?:\.\d+)?)", price_text)
        if match:
            return match.group(1)
    except:
        pass
    return ""

def split_size_parts(size_str: str):
    if not size_str:
        return []
    s = size_str.replace("×", "x")
    s = re.sub(r"\s+", " ", s).strip()
    return [p.strip() for p in re.split(r"\s*x\s*", s) if p.strip()]

def parse_size_dimensions(size_str: str):
    """
    Returns: Width, Depth, Diameter, Length, Height
    Rules:
      - token-start labels: H/W/D/L/Dia
      - Diameter first, then Depth
      - no labels + 2 values -> first=Length, last=Width
      - no labels + 3 values -> first=Length, second=Width, third=Depth
    """
    width = depth = diameter = length = height = ""
    parts = split_size_parts(size_str)
    if not parts:
        return width, depth, diameter, length, height

    def starts_with(token, prefixes):
        t = token.strip().lower()
        return any(t.startswith(p) for p in prefixes)

    has_label = any(
        starts_with(p, ["h", "w", "d", "l", "dia", "diam", "diameter"])
        for p in parts
    )

    if has_label:
        # Diameter first
        for p in parts:
            pl = p.strip().lower()
            if (pl.startswith("dia") or pl.startswith("diam") or pl.startswith("diameter")) and not diameter:
                diameter = txt_num(p)

        for p in parts:
            pl = p.strip().lower()
            if pl.startswith("h") and not height:
                height = txt_num(p)

        for p in parts:
            pl = p.strip().lower()
            if pl.startswith("w") and not width:
                width = txt_num(p)

        for p in parts:
            pl = p.strip().lower()
            if pl.startswith("l") and not length:
                length = txt_num(p)

        for p in parts:
            pl = p.strip().lower()
            if pl.startswith("d") and not (pl.startswith("dia") or pl.startswith("diam") or pl.startswith("diameter")) and not depth:
                depth = txt_num(p)

        return width, depth, diameter, length, height

    nums = [txt_num(p) for p in parts if txt_num(p)]
    if len(nums) == 2:
        length = nums[0]
        width = nums[1]
    elif len(nums) == 3:
        length, width, depth = nums[0], nums[1], nums[2]
    elif len(nums) == 1:
        length = nums[0]

    return width, depth, diameter, length, height

# -------------------
# Selenium setup
# -------------------
options = Options()
# options.add_argument("--headless=new")  # চাইলে headless চালাও
options.add_argument("--window-size=1400,900")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36"
)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 30)

try:
    for _, row in df.iterrows():
        product_url = str(row.get("Product URL", "")).strip()
        image_url = str(row.get("Image URL", "")).strip()
        product_name = str(row.get("Product Name", "")).strip()
        sku = str(row.get("SKU", "")).strip()

        if not product_url:
            continue

        print(f"Scraping: {product_url}")
        driver.get(product_url)

        try:
            wait.until(
                EC.any_of(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.o-text")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.c-attrlist")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'select[name="sizeOptions"]'))
                )
            )
        except:
            pass

        time.sleep(1.2)

        # Description: first div.o-text
        description = ""
        otexts = driver.find_elements(By.CSS_SELECTOR, "div.o-text")
        if otexts:
            description = norm_space(otexts[0].text)

        # Details + Size + Weight + Content + Repeat
        details = ""
        size = ""
        weight = ""
        content = ""
        repeat = ""
        attrlists = driver.find_elements(By.CSS_SELECTOR, "div.c-attrlist")
        if attrlists:
            details, size, weight, content, repeat = extract_details_size_weight(attrlists[0])

        # ✅ Size fallback from dropdown if missing
        if not size:
            size = extract_size_from_dropdown(driver)

        # ✅ Extract List Price
        list_price = extract_list_price(driver)

        # Split size
        width, depth, diameter, length, height = parse_size_dimensions(size)

        final_rows.append({
            "Product URL": product_url,
            "Image URL": image_url,
            "Product Name": product_name,
            "SKU": sku,
            "Product Family Id": product_name,
            "Description": description,
            "Weight": weight,
            "Width": width,
            "Depth": depth,
            "Diameter": diameter,
            "Length": length,
            "Height": height,
            "Details": details,
            "Size": size,
            "List Price": list_price,
            "Content": content,
            "Repeat": repeat
        })

finally:
    driver.quit()

final_df = pd.DataFrame(final_rows)

# ✅ exact column order
final_cols = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family Id",
    "Description", "Weight", "Width", "Depth", "Diameter", "Length", "Height",
    "Details", "Size", "List Price", "Content", "Repeat"
]
final_df = final_df.reindex(columns=final_cols)

final_df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ FINAL Excel created: {OUTPUT_FILE}")