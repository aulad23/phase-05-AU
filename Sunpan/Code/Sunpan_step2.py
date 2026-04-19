from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import time
import os
import re
import json

# =========================
# File Path
# =========================
script_path = os.path.dirname(os.path.abspath(__file__))
input_excel = os.path.join(script_path, "sunpan-pendants.xlsx")
output_excel = os.path.join(script_path, "sunpan-pendants_final.xlsx")

# Columns in final order
final_columns = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family Id",
    "Color", "Weight", "Width", "Depth", "Diameter", "Length", "Height",
    "Arm Height", "Seat Height", "Seat Width", "Seat Depth",
    "Finish", "Base", "Seat", "Cushion", "Wattage"  # ✅ Wattage added
]

# =========================
# Selenium Setup
# =========================
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.page_load_strategy = "eager"

driver = webdriver.Chrome(options=chrome_options)
driver.set_page_load_timeout(300)
wait = WebDriverWait(driver, 15)


# =========================
# Helper Functions
# =========================
def open_accordion(title_text):
    try:
        summary_tag = driver.find_element(
            By.XPATH,
            f"//h2[contains(text(),'{title_text}')]/ancestor::summary"
        )
        toggle_svg = summary_tag.find_element(By.CSS_SELECTOR, "svg.icon-caret")
        driver.execute_script("arguments[0].click();", toggle_svg)
        time.sleep(1)
    except:
        print(f"⚠ Could not open accordion: {title_text}")


def extract_table(table_css):
    try:
        table = driver.find_element(By.CSS_SELECTOR, table_css)
        rows = table.find_elements(By.TAG_NAME, "tr")
        out_lines = []
        for tr in rows:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if len(tds) == 2:
                label = tds[0].text.strip()
                value = tds[1].text.strip()
                out_lines.append(f"{label}: {value}")
        return "\n".join(out_lines)
    except:
        return ""


def parse_dimensions(dim_text):
    width = depth = height = diameter = length = ""
    parts = re.findall(r"([\d\.]+)\s*([WDHLDIA]+)", dim_text, re.IGNORECASE)
    for val, key in parts:
        key = key.lower()
        if key == "w" and not width:
            width = val
        elif key == "d" and not depth:
            depth = val
        elif key == "h" and not height:
            height = val
        elif "dia" in key and not diameter:
            diameter = val
        elif key == "l" and not length:
            length = val
    return width, depth, height, diameter, length


def extract_from_dimensions_table(dim_table_text):
    data = {}
    for line in dim_table_text.split("\n"):
        if ":" not in line:
            continue
        key, val = line.split(":", 1)
        key = key.strip().lower()
        val = val.strip()

        # weight line (Net / Carton / Gross)
        if "net weight" in key:
            data["Weight"] = val
        elif "carton weight" in key and "Weight" not in data:
            data["Weight"] = val
        elif "gross weight" in key and "Weight" not in data:
            data["Weight"] = val

        elif "overall dimensions" in key:
            w, d, h, dia, l = parse_dimensions(val)
            data.update({
                "Width": w, "Depth": d, "Height": h,
                "Diameter": dia, "Length": l
            })
        elif "arm height" in key:
            data["Arm Height"] = val
        elif "seat height" in key:
            data["Seat Height"] = val
        elif "seat width" in key:
            data["Seat Width"] = val
        elif "seat depth" in key:
            data["Seat Depth"] = val
    return data


def extract_from_construction_table(cons_table_text):
    data = {}
    for line in cons_table_text.split("\n"):
        if ":" not in line:
            continue
        key, val = line.split(":", 1)
        key = key.strip().lower()
        val = val.strip()
        # Finish
        if "material finish" in key:
            data["Finish"] = val
        # Base
        elif "base / legs" in key or "base finish" in key:
            if "Base" in data:
                data["Base"] += ", " + val
            else:
                data["Base"] = val
        # Seat
        elif "seat construction" in key or "seat cushion" in key or "cover fabric" in key:
            if "Seat" in data:
                data["Seat"] += ", " + val
            else:
                data["Seat"] = val
        # Cushion
        elif "back cushion" in key:
            data["Cushion"] = val
        # ✅ Wattage from construction: "Max Bulb Wattage"
        elif "max bulb wattage" in key:
            data["Wattage"] = val
    return data


def extract_color():
    # 1st priority: selected-option span (original logic)
    try:
        color_span = driver.find_element(By.CSS_SELECTOR, "span.selected-option-value--color")
        txt = color_span.text.strip()
        if txt:
            return txt
    except:
        pass

    # fallback: currently selected radio input
    try:
        selected_radio = driver.find_element(
            By.CSS_SELECTOR,
            "fieldset.js.product-form__input input[type='radio'][name='Color']:checked"
        )
        val = selected_radio.get_attribute("value")
        if val:
            return val.strip()
    except:
        pass

    return ""


def get_current_variant():
    """
    Return (variant_dict or None) for the currently selected variant.
    We try by ?variant=id first; fallback by color name (option1).
    """
    try:
        script_tag = driver.find_element(
            By.CSS_SELECTOR, "script[type='application/json'][data-product]"
        )
        prod_json = json.loads(script_tag.get_attribute("innerHTML"))
    except Exception as e:
        print(f"⚠ Could not load product JSON: {e}")
        return None

    variants = prod_json.get("variants", [])
    if not variants:
        return None

    # try via variant id from URL
    url = driver.current_url
    variant_id = None
    if "variant=" in url:
        try:
            variant_id = int(url.split("variant=")[1].split("&")[0])
        except:
            variant_id = None

    if variant_id:
        for v in variants:
            if v.get("id") == variant_id:
                return v

    # fallback: match by color / option1
    selected_color = extract_color()
    if selected_color:
        for v in variants:
            if v.get("option1") == selected_color or v.get("title") == selected_color:
                return v

    return None


def extract_image_url():
    """
    Always try to return CURRENT SELECTED VARIANT IMAGE.
    """
    # 1) Active media in gallery
    try:
        active_img = driver.find_element(
            By.CSS_SELECTOR,
            ".product__media.is-active img, .product__media-item.is-active img"
        )
        src = active_img.get_attribute("src")
        if src:
            if src.startswith("//"):
                src = "https:" + src
            return src
    except:
        pass

    # 2) Active thumbnail
    try:
        active_thumb = driver.find_element(
            By.CSS_SELECTOR,
            ".thumbnail.is-active img, .product__media-toggle.is-active img"
        )
        src = active_thumb.get_attribute("src")
        if src:
            if src.startswith("//"):
                src = "https:" + src
            return src
    except:
        pass

    # 3) From current variant JSON image
    try:
        variant = get_current_variant()
        if variant:
            image = variant.get("featured_image") or variant.get(
                "featured_media", {}
            ).get("preview_image", {})
            if isinstance(image, dict):
                src = image.get("src")
                if src:
                    if src.startswith("//"):
                        src = "https:" + src
                    return src
    except:
        pass

    # 4) Fallback: og:image
    try:
        meta = driver.find_element(By.CSS_SELECTOR, "meta[property='og:image']")
        content = meta.get_attribute("content")
        if content:
            return content.strip()
    except:
        pass

    return ""


def fill_from_variant_meta(row_index):
    """
    SKU + dimensions + construction window.variantMetafields থেকে আনার চেষ্টা।
    কিছু পেলে True return, নাহলে False.
    """
    filled_any = False

    variant = get_current_variant()
    if not variant:
        return False

    sku = variant.get("sku") or ""
    if sku:
        df.at[row_index, "SKU"] = sku
    else:
        return False

    # get window.variantMetafields from page
    try:
        vm = driver.execute_script("return window.variantMetafields || {};")
    except Exception as e:
        print(f"⚠ Could not read window.variantMetafields: {e}")
        return False

    meta = vm.get(sku)
    if not meta:
        return False

    details = meta.get("product_details", {})
    dimensions = details.get("dimensions", {}) or {}
    construction = details.get("construction", {}) or {}

    # Dimensions
    if dimensions:
        lines = [f"{k}: {v}" for k, v in dimensions.items() if v]
        dim_text = "\n".join(lines)
        dim_data = extract_from_dimensions_table(dim_text)
        for k, v in dim_data.items():
            if k in final_columns and v:
                df.at[row_index, k] = v
                filled_any = True

    # Construction
    if construction:
        lines = [f"{k}: {v}" for k, v in construction.items() if v]
        cons_text = "\n".join(lines)
        cons_data = extract_from_construction_table(cons_text)
        for k, v in cons_data.items():
            if k in final_columns and v:
                df.at[row_index, k] = v
                filled_any = True

    return filled_any


# =========================
# Load Excel
# =========================
df = pd.read_excel(input_excel)

# Ensure final columns exist
for col in final_columns:
    if col not in df.columns:
        df[col] = ""

# Copy Product Name to Product Family Id
df["Product Family Id"] = df["Product Name"]


# =========================
# Helper: scrape 1 configuration (1 selected color / variant)
# =========================
def scrape_configuration(row_index):
    # Color
    df.at[row_index, "Color"] = extract_color()

    # First try: variantMetafields (per-variant data)
    used_variant_meta = fill_from_variant_meta(row_index)

    # If variantMetafields did NOT give us anything, fallback to old HTML logic
    if not used_variant_meta:
        # Dimensions
        open_accordion("Dimensions")
        dim_table_text = extract_table("table.tab-dimensions__table")
        if dim_table_text:
            dim_data = extract_from_dimensions_table(dim_table_text)
            for k, v in dim_data.items():
                if k in final_columns and v:
                    df.at[row_index, k] = v

        # Construction
        open_accordion("Construction")
        cons_table_text = extract_table("table.tab-construction__table")
        if cons_table_text:
            cons_data = extract_from_construction_table(cons_table_text)
            for k, v in cons_data.items():
                if k in final_columns and v:
                    df.at[row_index, k] = v

    # Variant-wise URLs (always)
    df.at[row_index, "Product URL"] = driver.current_url
    df.at[row_index, "Image URL"] = extract_image_url()

    # Save
    df.to_excel(output_excel, index=False)
    print(f"Saved row {row_index} ({df.at[row_index, 'Color']}).")


# =========================
# MAIN LOOP (no duplicate processing)
# =========================
original_len = len(df)

for idx in range(original_len):
    row = df.iloc[idx]
    url = str(row["Product URL"]).strip()
    name = str(row["Product Name"])
    if not url or url.lower() == "nan":
        continue

    print(f"\nScraping: {name}")

    # Open Product URL
    try:
        driver.get(url)
        time.sleep(2)
    except:
        print("⚠ Page load failed.")
        continue

    # Find Color variations
    try:
        color_inputs = driver.find_elements(
            By.CSS_SELECTOR,
            "fieldset.js.product-form__input input[type='radio'][name='Color']"
        )
    except:
        color_inputs = []

    # Unique color values
    color_values = []
    for inp in color_inputs:
        val = inp.get_attribute("value")
        if val and val not in color_values:
            color_values.append(val)

    # CASE 1: No variation or only 1 color → single config
    if len(color_values) <= 1:
        scrape_configuration(idx)
        continue

    # CASE 2: Multiple variations → each color = separate row
    print(f"Found {len(color_values)} color variations.")
    base_row = df.loc[idx].copy()
    first = True

    for val in color_values:
        try:
            radio = driver.find_element(
                By.CSS_SELECTOR,
                f"fieldset.js.product-form__input input[type='radio'][name='Color'][value='{val}']"
            )
            driver.execute_script("arguments[0].click();", radio)
            time.sleep(1.5)
        except Exception as e:
            print(f"⚠ Could not click color '{val}': {e}")
            continue

        if first:
            target_index = idx
            first = False
        else:
            target_index = len(df)
            df.loc[target_index] = base_row

        scrape_configuration(target_index)


# =========================
# FINAL SAVE
# =========================
df = df[final_columns]
df.to_excel(output_excel, index=False)
print("\n✅ DONE — All data (including per-variant URL, image, SKU, dimensions, construction, wattage) saved!")

driver.quit()
