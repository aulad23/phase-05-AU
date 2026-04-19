"""
Eichholtz Step 2 - Product Detail Scraper (FIXED v4)
Usage: python eichholtz_step2_fixed.py
Input:  eichholtz_products.xlsx  (from Step 1)
Output: eichholtz_products_Final.xlsx

Column Order:
  Product URL, Image URL, Product Name, SKU, Product Family Id,
  Description, Weight, Width, Depth, Diameter, Length, Height,
  Seat Depth, Seat Width, Seat Height, Arm Height, Dimension,
  Finish, Wattage, Socket, Fabric, Specifications

Weight: শুধু LBS value নেয়, KG ignore করে
Finish: "General info:" prefix সরায়
Fabric: "Fabric composition:" prefix সরায়
Socket: "Lamp holder:" prefix সরায়
Wattage: "Max wattage:" prefix সরায়

FIX v4: get_dimension() now ALWAYS tries modal first to get full dimensions
         (SD, SW, SH, AH). dim_cm from page source is only a FALLBACK.
"""

import re
import json
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# ===================== CONFIG =====================
INPUT_FILE   = "eichholtz_Lounge_Chairs.xlsx"
OUTPUT_FILE  = "eichholtz_Lounge_Chairs_Final.xlsx"
WAIT_SECONDS = 8
MODAL_WAIT   = 3.0
MAX_RETRIES  = 3
# ==================================================

CM_TO_INCH = 0.393701

# ─── Final column order ───
FINAL_COLUMN_ORDER = [
    "Product URL",
    "Image URL",
    "Product Name",
    "SKU",
    "Product Family Id",
    "Description",
    "Weight",
    "Width",
    "Depth",
    "Diameter",
    "Length",
    "Height",
    "Seat Depth",
    "Seat Width",
    "Seat Height",
    "Arm Height",
    "Shade Details",
    "Dimension",
    "Finish",
    "Wattage",
    "Socket",
    "Fabric",
    "Specifications",
]

# Dimension abbreviation → column name
DIM_ABBR_MAP = {
    "W":  "Width",
    "D":  "Depth",
    "H":  "Height",
    "L":  "Length",
    "SD": "Seat Depth",
    "SH": "Seat Height",
    "SW": "Seat Width",
    "AH": "Arm Height",
}

# Specification target columns
SPEC_KEYS = ["Finish", "Weight", "Fabric", "Wattage", "Socket"]


def get_driver():
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--window-size=1400,900")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    prefs = {"profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)


def human_delay(min_s=0.5, max_s=1.5):
    time.sleep(random.uniform(min_s, max_s))


def js_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    human_delay(0.3, 0.6)
    driver.execute_script("arguments[0].click();", element)


def close_modal(driver):
    try:
        close_selectors = [
            "[aria-label='Close']", ".modal-close", "button.close",
            "[x-on\\:click*='close']", "button[type='button'][class*='close']",
            ".close-button",
        ]
        for sel in close_selectors:
            try:
                btns = driver.find_elements(By.CSS_SELECTOR, sel)
                for btn in btns:
                    if btn.is_displayed():
                        js_click(driver, btn)
                        human_delay(0.8, 1.2)
                        return
            except Exception:
                pass
    except Exception:
        pass
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        human_delay(0.8, 1.2)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════
# JSON-LD
# ═══════════════════════════════════════════════════════════

def extract_json_ld(driver):
    try:
        scripts = driver.find_elements(By.CSS_SELECTOR, "script[type='application/ld+json']")
        for script in scripts:
            try:
                data = json.loads(script.get_attribute("innerHTML"))
                if isinstance(data, list):
                    for item in data:
                        if isinstance(item, dict) and item.get("@type") in ["Product", "product"]:
                            return item
                elif isinstance(data, dict):
                    if data.get("@type") in ["Product", "product"]:
                        return data
                    for item in data.get("@graph", []):
                        if isinstance(item, dict) and item.get("@type") in ["Product", "product"]:
                            return item
            except Exception:
                pass
    except Exception:
        pass
    return {}


# ═══════════════════════════════════════════════════════════
# PAGE SOURCE EXTRACT
# ═══════════════════════════════════════════════════════════

def extract_from_page_source(driver):
    result = {
        "description": "",
        "family_id": "",
        "dimensions_cm": {},
    }
    try:
        src = driver.page_source

        for pattern in [
            r'"family_id"\s*:\s*"?(\w+)"?',
            r'"familyId"\s*:\s*"?(\w+)"?',
            r'data-family-id="([^"]+)"',
            r'"product_family_id"\s*:\s*"?(\w+)"?',
        ]:
            m = re.search(pattern, src, re.IGNORECASE)
            if m:
                result["family_id"] = m.group(1).strip()
                break

        for pattern in [
            r'"description"\s*:\s*"((?:[^"\\]|\\.)+)"',
            r'"short_description"\s*:\s*"((?:[^"\\]|\\.)+)"',
        ]:
            m = re.search(pattern, src, re.IGNORECASE)
            if m:
                raw = m.group(1)
                try:
                    raw = json.loads(f'"{raw}"')
                except Exception:
                    pass
                raw = re.sub(r'<[^>]+>', ' ', raw).strip()
                if raw and len(raw) > 20:
                    result["description"] = raw
                    break

        dim_map = {}
        for axis, patterns in {
            "W": [r'"width(?:_cm)?"\s*:\s*([\d.]+)', r'"breedte"\s*:\s*([\d.]+)'],
            "D": [r'"depth(?:_cm)?"\s*:\s*([\d.]+)', r'"diepte"\s*:\s*([\d.]+)'],
            "H": [r'"height(?:_cm)?"\s*:\s*([\d.]+)', r'"hoogte"\s*:\s*([\d.]+)'],
            "L": [r'"length(?:_cm)?"\s*:\s*([\d.]+)', r'"lengte"\s*:\s*([\d.]+)'],
        }.items():
            for pat in patterns:
                m = re.search(pat, src, re.IGNORECASE)
                if m:
                    dim_map[axis] = float(m.group(1))
                    break
        result["dimensions_cm"] = dim_map

    except Exception:
        pass
    return result


# ═══════════════════════════════════════════════════════════
# DESCRIPTION
# ═══════════════════════════════════════════════════════════

def get_description(driver, json_ld_data, page_data):
    try:
        for xpath in [
            "//span[@x-text][contains(text(),'Read more') or contains(text(),'read more')]",
            "//span[normalize-space(text())='Read more']",
            "//span[normalize-space(text())='Read More']",
            "//button[contains(text(),'Read more')]",
        ]:
            try:
                btn = driver.find_element(By.XPATH, xpath)
                if btn.is_displayed():
                    js_click(driver, btn)
                    human_delay(1.2, 1.8)
                    break
            except Exception:
                pass
    except Exception:
        pass

    js_get_text = """
        var els = document.querySelectorAll('[x-show]');
        for (var i = 0; i < els.length; i++) {
            var el = els[i];
            var xshow = el.getAttribute('x-show');
            if (xshow && xshow.trim() === 'expanded') {
                var t = el.innerText.trim();
                if (t && t.length > 20) return t;
            }
        }
        return '';
    """
    try:
        desc = driver.execute_script(js_get_text)
        if desc and len(desc.strip()) > 20:
            return desc.strip()
    except Exception:
        pass

    for sel in ["div[x-show='expanded']", 'div[x-show="expanded"]']:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                t = driver.execute_script("return arguments[0].innerText;", el).strip()
                if t and len(t) > 20:
                    return t
        except Exception:
            pass

    desc = str(json_ld_data.get("description", "")).strip()
    if desc and len(desc) > 20:
        return re.sub(r'<[^>]+>', ' ', desc).strip()

    desc = page_data.get("description", "").strip()
    if desc and len(desc) > 20:
        return desc

    js_short = """
        var els = document.querySelectorAll('[x-show]');
        for (var i = 0; i < els.length; i++) {
            var el = els[i];
            var xshow = el.getAttribute('x-show');
            if (xshow && xshow.trim() === '!expanded') {
                var t = el.innerText.trim();
                if (t && t.length > 10) return t;
            }
        }
        return '';
    """
    try:
        desc = driver.execute_script(js_short)
        if desc and len(desc.strip()) > 10:
            return desc.strip()
    except Exception:
        pass

    for sel in ["[itemprop='description']", ".product-description", ".description"]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                t = driver.execute_script("return arguments[0].innerText;", el).strip()
                if t and len(t) > 20:
                    return t
        except Exception:
            pass

    return ""


# ═══════════════════════════════════════════════════════════
# DIMENSION  (FIX v4: Modal FIRST, dim_cm as FALLBACK)
# ═══════════════════════════════════════════════════════════

def cm_to_inch_str(cm_val):
    inch = round(float(cm_val) * CM_TO_INCH, 2)
    return f"{inch}″"


def parse_shade_from_text(text):
    if not text:
        return ""
    m = re.search(
        r'[Ss]hade[:\s]*(.*?)(?:\n[A-Z]|\Z)',
        text, re.DOTALL
    )
    if not m:
        return ""
    shade_text = m.group(1).strip()
    numbers = re.findall(r'(\d+(?:\.\d+)?)(?:\s*[″"])', shade_text)
    if numbers:
        return ",".join(numbers)
    return ""


def _has_extended_dims(dim_str):
    """Check if dimension string contains SD/SW/SH/AH (extended dimensions)."""
    return bool(re.search(r'(?:SD|SH|SW|AH)\.', dim_str))


def _try_modal_dimension(driver):
    """
    Try to open Dimensions modal, click inch tab, and extract dimension string.
    Returns (dimension_str, shade_details) or (None, "") if modal fails.
    """
    shade_details = ""

    try:
        dim_btn = None
        for p in driver.find_elements(By.TAG_NAME, "p"):
            if p.text.strip() == "Dimensions":
                dim_btn = p
                break

        if dim_btn is None:
            for xpath in [
                "//button[normalize-space(text())='Dimensions']",
                "//*[contains(@class,'cursor-pointer')][.//text()[normalize-space()='Dimensions']]",
                "//*[@x-on:click][.//text()[normalize-space()='Dimensions']]",
                "//span[normalize-space(text())='Dimensions']/parent::*",
            ]:
                try:
                    el = driver.find_element(By.XPATH, xpath)
                    if el.is_displayed():
                        dim_btn = el
                        break
                except Exception:
                    pass

        if not dim_btn:
            return None, ""

        js_click(driver, dim_btn)
        time.sleep(MODAL_WAIT)

        # ── Capture full modal text FIRST for shade extraction ──
        full_modal_text = ""
        try:
            modal_els = driver.find_elements(
                By.CSS_SELECTOR,
                "[role='dialog'], [x-show]:not([x-show='false']), .modal-content"
            )
            for mel in modal_els:
                try:
                    t = mel.text.strip()
                    if t and len(t) > 10:
                        full_modal_text = t
                        break
                except Exception:
                    pass
        except Exception:
            pass

        # Click "inch" tab
        for xpath in [
            "//span[normalize-space(text())='inch']",
            "//button[normalize-space(text())='inch']",
            "//*[normalize-space(text())='inch' and contains(@class,'tab')]",
            "//*[normalize-space(text())='inch']",
        ]:
            try:
                inch_el = driver.find_element(By.XPATH, xpath)
                if inch_el.is_displayed():
                    js_click(driver, inch_el)
                    human_delay(1.0, 1.5)
                    break
            except Exception:
                pass

        # ── Re-capture modal text AFTER inch tab click ──
        try:
            modal_els = driver.find_elements(
                By.CSS_SELECTOR,
                "[role='dialog'], [x-show]:not([x-show='false']), .modal-content"
            )
            for mel in modal_els:
                try:
                    t = mel.text.strip()
                    if t and len(t) > 10:
                        full_modal_text = t
                        break
                except Exception:
                    pass
        except Exception:
            pass

        # Extract shade from full modal text
        shade_details = parse_shade_from_text(full_modal_text)

        # Method 1: x-show inch container
        for sel in ["[x-show*='inch']", ".measurement-inch", "div[data-unit='inch']"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    txt = el.text.strip()
                    if txt and re.search(r'[\d.]+', txt):
                        close_modal(driver)
                        return txt, shade_details
            except Exception:
                pass

        # Method 2: modal content regex
        try:
            modal_candidates = driver.find_elements(
                By.CSS_SELECTOR,
                "[role='dialog'], [x-show]:not([x-show='false']), .modal-content"
            )
            for mc in modal_candidates:
                try:
                    txt = mc.text
                    # Match standard (W. 29.92″), extended (SD. 23.62″) and diameter (Ø 24.80″)
                    matches = re.findall(
                        r'(?:(?:SD|SH|SW|AH|[WDHL])\.\s*[\d.]+[″\'"]|[ØøΦφ]\s*[\d.]+[″\'"])',
                        txt
                    )
                    if matches:
                        close_modal(driver)
                        return " | ".join(matches), shade_details
                    matches2 = re.findall(r'[\d.]+\s*[″]', txt)
                    if len(matches2) >= 2:
                        close_modal(driver)
                        return " | ".join(matches2), shade_details
                except Exception:
                    pass
        except Exception:
            pass

        close_modal(driver)

    except Exception:
        try:
            close_modal(driver)
        except Exception:
            pass

    return None, shade_details


def get_dimension(driver, page_data):
    """
    FIX v4: ALWAYS try modal first to get full dimensions (including SD, SW, SH, AH).
    Only fall back to dim_cm from page source if modal fails.
    """
    shade_details = ""

    # ═══ STEP 1: Always try modal FIRST (has SD, SW, SH, AH) ═══
    modal_dim, shade_details = _try_modal_dimension(driver)
    if modal_dim:
        return {"dimension": modal_dim, "shade_details": shade_details}

    # ═══ STEP 2: Fallback — dim_cm from page source (only W, D, H, L) ═══
    dim_cm = page_data.get("dimensions_cm", {})
    if len(dim_cm) >= 2:
        parts = []
        for axis in ["W", "D", "H", "L"]:
            if axis in dim_cm:
                parts.append(f"{axis}. {cm_to_inch_str(dim_cm[axis])}")
        if parts:
            return {"dimension": " | ".join(parts), "shade_details": shade_details}

    # ═══ STEP 3: Fallback — regex from page source ═══
    try:
        src = driver.page_source

        # Regex matches any abbreviation (W, D, H, L, SD, SH, SW, AH) AND Ø diameter
        _seg = r'(?:(?:SD|SH|SW|AH|[WDHL])\.\s*[\d.]+[″"]|[ØøΦφ]\s*[\d.]+[″"])'
        m = re.search(
            r'((?:' + _seg + r'\s*\|\s*)*' + _seg + r')',
            src
        )
        if m:
            return {"dimension": m.group(1).strip(), "shade_details": shade_details}

        m2 = re.search(r'([\d.]+)\s*[xX×]\s*([\d.]+)\s*[xX×]\s*([\d.]+)\s*cm', src)
        if m2:
            w = round(float(m2.group(1)) * CM_TO_INCH, 2)
            d = round(float(m2.group(2)) * CM_TO_INCH, 2)
            h = round(float(m2.group(3)) * CM_TO_INCH, 2)
            return {"dimension": f"W. {w}″ | D. {d}″ | H. {h}″", "shade_details": shade_details}
    except Exception:
        pass

    return {"dimension": "", "shade_details": shade_details}


# ═══════════════════════════════════════════════════════════
# SPECIFICATIONS (raw scraping)
# ═══════════════════════════════════════════════════════════

def get_specifications(driver):
    def parse_table_rows(rows):
        specs = []
        for row in rows:
            try:
                cells = row.find_elements(By.CSS_SELECTOR, "th, td")
                if len(cells) >= 2:
                    key = cells[0].text.strip()
                    val = cells[1].text.strip()
                    if key and val and key.lower() not in ["", "attribute", "value"]:
                        specs.append(f"{key}: {val}")
            except Exception:
                pass
        return specs

    spec_btn = None
    for p in driver.find_elements(By.TAG_NAME, "p"):
        if p.text.strip() == "Specifications":
            spec_btn = p
            break

    if spec_btn is None:
        for xpath in [
            "//button[normalize-space(text())='Specifications']",
            "//*[contains(@class,'cursor-pointer')][.//text()[normalize-space()='Specifications']]",
            "//*[normalize-space(text())='Specifications'][@x-on:click or @onclick]",
            "//span[normalize-space(text())='Specifications']/parent::*",
        ]:
            try:
                el = driver.find_element(By.XPATH, xpath)
                if el.is_displayed():
                    spec_btn = el
                    break
            except Exception:
                pass

    if spec_btn:
        try:
            js_click(driver, spec_btn)
            time.sleep(MODAL_WAIT)

            for sel in [
                "table.additional-attributes tr",
                "[role='dialog'] table tr",
                ".modal-content table tr",
                "table tr",
            ]:
                rows = driver.find_elements(By.CSS_SELECTOR, sel)
                if rows:
                    specs = parse_table_rows(rows)
                    if specs:
                        close_modal(driver)
                        return " | ".join(specs)

            try:
                modal = driver.find_element(By.CSS_SELECTOR, "[role='dialog'], .modal-content")
                dts = modal.find_elements(By.TAG_NAME, "dt")
                dds = modal.find_elements(By.TAG_NAME, "dd")
                specs = []
                for dt, dd in zip(dts, dds):
                    k = dt.text.strip()
                    v = dd.text.strip()
                    if k and v:
                        specs.append(f"{k}: {v}")
                if specs:
                    close_modal(driver)
                    return " | ".join(specs)
            except Exception:
                pass

            try:
                modal = driver.find_element(By.CSS_SELECTOR, "[role='dialog'], .modal-content")
                lis = modal.find_elements(By.TAG_NAME, "li")
                specs = [li.text.strip() for li in lis if li.text.strip()]
                if specs:
                    close_modal(driver)
                    return " | ".join(specs)
            except Exception:
                pass

            close_modal(driver)
        except Exception:
            close_modal(driver)

    rows = driver.find_elements(By.CSS_SELECTOR, "table tr")
    if rows:
        specs = parse_table_rows(rows)
        if specs:
            return " | ".join(specs)

    try:
        src = driver.page_source
        m = re.search(r'"attributes"\s*:\s*(\[.*?\])', src, re.DOTALL)
        if m:
            attrs = json.loads(m.group(1))
            specs = []
            for attr in attrs:
                if isinstance(attr, dict):
                    label = attr.get("label", attr.get("code", ""))
                    value = attr.get("value", "")
                    if label and value:
                        specs.append(f"{label}: {value}")
            if specs:
                return " | ".join(specs)
    except Exception:
        pass

    return ""


# ═══════════════════════════════════════════════════════════
# PRODUCT FAMILY ID
# ═══════════════════════════════════════════════════════════

def get_product_family_id(product_name):
    name = str(product_name).strip()
    if not name or name == "nan":
        return ""
    if "-" in name:
        return name.split("-")[0].strip()
    return name


# ═══════════════════════════════════════════════════════════
# PARSE DIMENSION STRING → SEPARATE COLUMNS
# ═══════════════════════════════════════════════════════════

def parse_dimension_string(dim_str):
    """
    Parse:
      "W. 29.92″ | D. 39.37″ | H. 30.71″ | SD. 23.62″ | SH. 17.32″ | SW. 22.24″ | AH. 22.44″"
      "Ø 31.50″ | D. 2.17″"
      "L. 157.48″ | W. 118.11″"

    → Width, Depth, Height, Diameter, Length, Seat Depth, Seat Width, Seat Height, Arm Height
    """
    result = {
        "Width": "", "Depth": "", "Height": "", "Diameter": "",
        "Length": "", "Seat Depth": "", "Seat Height": "",
        "Seat Width": "", "Arm Height": "",
    }

    if not dim_str or str(dim_str).strip() in ("", "nan"):
        return result

    dim_str = str(dim_str).strip()

    # Diameter: Ø 31.50″
    m_dia = re.search(r'[ØøΦφ]\s*([\d.]+)', dim_str)
    if m_dia:
        result["Diameter"] = str(round(float(m_dia.group(1)), 2))

    # Match longer abbreviations first (SD before D, SH before H, etc.)
    for abbr in ["SD", "SH", "SW", "AH", "W", "D", "H", "L"]:
        pattern = r'(?<![A-Za-z])' + abbr + r'\.\s*([\d.]+)'
        m = re.search(pattern, dim_str)
        if m:
            col_name = DIM_ABBR_MAP.get(abbr, "")
            if col_name:
                result[col_name] = str(round(float(m.group(1)), 2))

    return result


# ═══════════════════════════════════════════════════════════
# PARSE SPECIFICATIONS STRING → SEPARATE COLUMNS
# ═══════════════════════════════════════════════════════════

def clean_finish(raw_val):
    val = raw_val.strip()
    val = re.sub(r'^general\s*info\s*[:]\s*', '', val, flags=re.IGNORECASE).strip()
    return val


def clean_weight(raw_val):
    val = raw_val.strip()
    if re.search(r'\bKG\b', val, re.IGNORECASE) and not re.search(r'\bLBS\b', val, re.IGNORECASE):
        return ""
    m = re.search(r'LBS\s*[:]\s*([\d.]+)', val, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r'([\d.]+)\s*lbs', val, re.IGNORECASE)
    if m:
        return m.group(1)
    val = re.sub(r'^max\s*weight\s*load\s*(lbs|kg)?\s*[:]\s*', '', val, flags=re.IGNORECASE).strip()
    m = re.search(r'([\d.]+)', val)
    return m.group(1) if m else val


def clean_fabric(raw_val):
    val = raw_val.strip()
    val = re.sub(r'^fabric\s*composition\s*[:]\s*', '', val, flags=re.IGNORECASE).strip()
    val = re.sub(r'^fabric\s*[:]\s*', '', val, flags=re.IGNORECASE).strip()
    # Remove trailing " |" or "|" (leftover from split boundaries)
    val = re.sub(r'\s*\|\s*$', '', val).strip()
    return val


def clean_socket(raw_val):
    val = raw_val.strip()
    val = re.sub(r'^lamp\s*holder\s*[:]\s*', '', val, flags=re.IGNORECASE).strip()
    return val


def clean_wattage(raw_val):
    val = raw_val.strip()
    val = re.sub(r'^max\s*wattage\s*[:]\s*', '', val, flags=re.IGNORECASE).strip()
    m = re.search(r'([\d.]+)\s*(watt|w)\b', val, re.IGNORECASE)
    if m:
        return f"{m.group(1)} Watt"
    return val


SPEC_CLEANERS = {
    "Finish":  clean_finish,
    "Weight":  clean_weight,
    "Fabric":  clean_fabric,
    "Wattage": clean_wattage,
    "Socket":  clean_socket,
}

SPEC_KEY_ALIASES = {
    "Finish":  ["finish", "general info", "general_info"],
    "Weight":  ["weight", "max weight load", "max weight load lbs",
                "max_weight_load"],
    "Fabric":  ["fabric", "fabric composition", "fabric_composition",
                "material composition"],
    "Wattage": ["wattage", "max wattage", "max_wattage"],
    "Socket":  ["socket", "lamp holder", "lamp_holder", "lampholder"],
    "_skip":   ["extra info", "extra_info",
                "max weight load kg",
                "lamp holder qty", "lamp holder quantity",
                "light bulbs included", "light bulbs",
                "bulbs included", "bulb included",
                "light source included",
                "number of lights", "number of light sources",
                "dimmable", "dimbaar",
                "ip rating", "ip_rating",
                "cord length", "cable length",
                "switch type", "switch",
                "power source", "voltage",
                "color temperature", "colour temperature",
                "indoor/outdoor",
                ],
}


def _match_spec_key(raw_key):
    raw_lower = raw_key.strip().lower()
    for target, aliases in SPEC_KEY_ALIASES.items():
        for alias in aliases:
            if raw_lower == alias:
                return target
    return None


def parse_specifications_string(spec_str):
    result = {k: "" for k in SPEC_KEYS}

    if not spec_str or str(spec_str).strip() in ("", "nan"):
        return result

    raw_parts = [p.strip() for p in str(spec_str).strip().split(" | ")]

    parsed_pairs = []
    current_key = None
    current_val_parts = []

    for part in raw_parts:
        found_key = False

        if ":" in part:
            potential_key = part.split(":", 1)[0].strip()
            potential_val = part.split(":", 1)[1].strip()

            target = _match_spec_key(potential_key)
            if target:
                if current_key is not None:
                    parsed_pairs.append((current_key, " | ".join(current_val_parts)))
                current_key = target
                current_val_parts = [potential_val] if potential_val else []
                found_key = True

        if not found_key:
            if current_key is not None:
                current_val_parts.append(part)

    if current_key is not None:
        parsed_pairs.append((current_key, " | ".join(current_val_parts)))

    for target_key, raw_val in parsed_pairs:
        if target_key in SPEC_KEYS:
            cleaner = SPEC_CLEANERS.get(target_key, lambda x: x.strip())
            cleaned = cleaner(raw_val)
            if cleaned:
                result[target_key] = cleaned

    return result


# ═══════════════════════════════════════════════════════════
# MAIN SCRAPE FUNCTION
# ═══════════════════════════════════════════════════════════

def scrape_detail(driver, url):
    result = {
        "Description":    "",
        "Dimension":      "",
        "Shade Details":  "",
        "Specifications": "",
    }

    driver.get(url)

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
        )
    except Exception:
        pass

    time.sleep(WAIT_SECONDS)

    try:
        driver.execute_script("window.scrollTo(0, 300);")
        human_delay(0.5, 1.0)
        driver.execute_script("window.scrollTo(0, 600);")
        human_delay(0.5, 1.0)
        driver.execute_script("window.scrollTo(0, 0);")
        human_delay(0.3, 0.8)
    except Exception:
        pass

    json_ld = extract_json_ld(driver)
    page_data = extract_from_page_source(driver)

    result["Description"]    = get_description(driver, json_ld, page_data)
    dim_result               = get_dimension(driver, page_data)
    result["Dimension"]      = dim_result["dimension"]
    result["Shade Details"]  = dim_result["shade_details"]
    result["Specifications"] = get_specifications(driver)

    return result


# ═══════════════════════════════════════════════════════════
# REORDER COLUMNS
# ═══════════════════════════════════════════════════════════

def reorder_columns(df):
    ordered = []
    for col in FINAL_COLUMN_ORDER:
        if col in df.columns:
            ordered.append(col)
    for col in df.columns:
        if col not in ordered:
            ordered.append(col)
    return df[ordered]


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    import os

    if os.path.exists(OUTPUT_FILE):
        print(f"⚡ Resume mode: '{OUTPUT_FILE}' found — loading previous progress...")
        df = pd.read_excel(OUTPUT_FILE)
        done_count = df["Description"].apply(
            lambda x: bool(str(x).strip()) and str(x).strip() not in ("nan", "")
        ).sum()
        print(f"   Already scraped: {done_count} / {len(df)} rows — continuing from where it stopped.\n")
    else:
        print(f"Reading: {INPUT_FILE}")
        df = pd.read_excel(INPUT_FILE)

    if "Product URL" not in df.columns:
        print("ERROR: 'Product URL' column not found!")
        return

    for col in ["Product Family Id", "Description", "Dimension", "Shade Details", "Specifications"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).replace("nan", "")

    dim_cols = ["Width", "Depth", "Height", "Diameter", "Length",
                "Seat Depth", "Seat Width", "Seat Height", "Arm Height",
                "Shade Details"]
    for col in dim_cols:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].astype(str).replace("nan", "")

    for col in SPEC_KEYS:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].astype(str).replace("nan", "")

    total = len(df)
    print(f"Total rows: {total}")
    print("Browser window will open — do not close it!\n")

    driver = get_driver()
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )

    failed = []

    try:
        for idx, row in df.iterrows():
            url = str(row.get("Product URL", "")).strip()
            if not url or url == "nan":
                print(f"[{idx+1}/{total}] Skip — no URL")
                continue

            existing_desc = str(row.get("Description", "")).strip()
            if existing_desc and existing_desc not in ("nan", ""):
                dim_parsed = parse_dimension_string(row.get("Dimension", ""))
                for col_name, val in dim_parsed.items():
                    df.at[idx, col_name] = str(val) if val != "" else ""

                spec_parsed = parse_specifications_string(row.get("Specifications", ""))
                for col_name, val in spec_parsed.items():
                    df.at[idx, col_name] = str(val) if val != "" else ""

                print(f"[{idx+1}/{total}] Already done — parsed columns updated")
                continue

            print(f"[{idx+1}/{total}] {url}")

            success = False
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    detail = scrape_detail(driver, url)

                    df.at[idx, "Product Family Id"] = get_product_family_id(row.get("Product Name", ""))
                    df.at[idx, "Description"]       = str(detail["Description"])
                    df.at[idx, "Dimension"]         = str(detail["Dimension"])
                    df.at[idx, "Shade Details"]     = str(detail["Shade Details"])
                    df.at[idx, "Specifications"]    = str(detail["Specifications"])

                    dim_parsed = parse_dimension_string(detail["Dimension"])
                    for col_name, val in dim_parsed.items():
                        df.at[idx, col_name] = str(val) if val != "" else ""

                    spec_parsed = parse_specifications_string(detail["Specifications"])
                    for col_name, val in spec_parsed.items():
                        df.at[idx, col_name] = str(val) if val != "" else ""

                    fam = get_product_family_id(row.get("Product Name", ""))
                    print(f"  ✓ Family : {fam or '-'}")
                    print(f"    Desc   : {detail['Description'][:70] or '-'}...")
                    print(f"    Dim    : {detail['Dimension'] or '-'}")
                    print(f"    Shade  : {detail['Shade Details'] or '-'}")
                    print(f"    → W:{dim_parsed['Width']}  D:{dim_parsed['Depth']}  H:{dim_parsed['Height']}  "
                          f"L:{dim_parsed['Length']}  "
                          f"Ø:{dim_parsed['Diameter']}  SD:{dim_parsed['Seat Depth']}  "
                          f"SH:{dim_parsed['Seat Height']}  SW:{dim_parsed['Seat Width']}  AH:{dim_parsed['Arm Height']}")
                    print(f"    Spec   : {(detail['Specifications'] or '-')[:70]}")
                    print(f"    → Finish:{spec_parsed['Finish']}  Weight(LBS):{spec_parsed['Weight']}  "
                          f"Fabric:{spec_parsed['Fabric']}  Wattage:{spec_parsed['Wattage']}  Socket:{spec_parsed['Socket']}")
                    success = True
                    break

                except Exception as e:
                    print(f"  [Attempt {attempt}/{MAX_RETRIES}] Error: {type(e).__name__}: {str(e)[:80]}")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    time.sleep(random.uniform(3, 6))
                    try:
                        driver = get_driver()
                        driver.execute_script(
                            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                        )
                    except Exception:
                        break

            if not success:
                failed.append(url)
                print(f"  ✗ FAILED: {url}")

            reorder_columns(df).to_excel(OUTPUT_FILE, index=False)
            human_delay(1.5, 3.0)

    except KeyboardInterrupt:
        print("\nStopped by user. Saving progress...")
    except Exception as e:
        print(f"Fatal error: {e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    for col in dim_cols + SPEC_KEYS:
        df[col] = df[col].apply(lambda x: "" if str(x).strip() in ("", "nan", "0", "0.0") else x)

    df_final = reorder_columns(df)
    df_final.to_excel(OUTPUT_FILE, index=False)

    print(f"\n{'='*55}")
    print(f"Done! Saved: '{OUTPUT_FILE}'")
    print(f"Success: {total - len(failed)} / {total}")
    print(f"\nColumn order:")
    for i, col in enumerate(df_final.columns, 1):
        print(f"  {i}. {col}")
    if failed:
        print(f"\nFailed ({len(failed)}):")
        for u in failed:
            print(f"  - {u}")


if __name__ == "__main__":
    main()