# -*- coding: utf-8 -*-
# Rejuvenation Step-2 (Details + Variations) — FINAL (Rigdon + Hosford Dimensions FIXED + ORDER UPDATED)
# ✅ Description: HTML tag remove -> only clean text (Description not exported per your last order)
# ✅ Dimensions extraction: works even when "Dimensions" tab doesn't exist (Rigdon lighting specs)
# ✅ Dimension text থেকে extract:
#    Weight, Width, Depth, Diameter, Height,
#    Seat Width, Seat Height, Seat Depth, Arm Height
# ✅ NEW extract from Dimension:
#    Shade Details, Wattage (Max Wattage + Energy efficiency), Socket, Canopy
# ✅ Auto-save every 10 base products
# ✅ Column order EXACT as requested

import re
import time
import json
import html
import pandas as pd
from openpyxl.descriptors import Descriptor

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =========================
# CONFIG
# =========================
VENDOR_NAME = "Rejuvenation"
CATEGORY_NAME = "Ottomans"
INPUT_XLSX  = "rejuvenation_Ottomans.xlsx"
OUTPUT_XLSX = "rejuvenation_Ottomans_DETAILS.xlsx"

WAIT_TIMEOUT = 18
POLITE_DELAY = 0.6
AUTOSAVE_EVERY = 10


# =========================
# BASIC HELPERS
# =========================
def clean_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def sku_fallback(vendor: str, category: str, idx: int) -> str:
    v = re.sub(r"[^A-Za-z]", "", vendor or "").upper()[:3] or "VEN"
    c = re.sub(r"[^A-Za-z]", "", category or "").upper()[:2] or "CA"
    return f"{v}{c}{idx}"

def get_text_safe(driver, css: str, timeout: int = 2) -> str:
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css))
        )
        return clean_space(el.text)
    except Exception:
        return ""

def click_if_exists(driver, css: str, timeout: int = 2) -> bool:
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, css))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.2)
        el.click()
        return True
    except Exception:
        return False

def extract_current_sku(driver) -> str:
    t = get_text_safe(driver, '[data-test-id="sku-display"]', timeout=2)
    if not t:
        return ""
    m = re.search(r"SKU\s*:\s*([A-Za-z0-9\-_]+)", t)
    return m.group(1).strip() if m else ""

def html_to_text(s: str) -> str:
    """Remove HTML tags + decode entities; return clean text only."""
    s = (s or "").strip()
    if not s:
        return ""
    s = html.unescape(s)

    if "<" in s and ">" in s:
        try:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(s, "html.parser")
            for br in soup.find_all(["br", "p", "li", "div"]):
                br.append(" ")
            s = soup.get_text(" ", strip=True)
        except Exception:
            s = re.sub(r"<\s*br\s*/?\s*>", " ", s, flags=re.I)
            s = re.sub(r"</\s*(p|li|div|ul|ol|h\d)\s*>", " ", s, flags=re.I)
            s = re.sub(r"<[^>]+>", " ", s)
            s = re.sub(r"\s+", " ", s).strip()

    return clean_space(s)


# =========================
# DIMENSION PARSER
# =========================
def _norm_dim_text(t: str) -> str:
    t = (t or "")
    t = t.replace("″", '"').replace("”", '"').replace("“", '"')
    t = t.replace("’", "'").replace("–", "-").replace("—", "-")
    t = re.sub(r"[ \t]+", " ", t)
    return t.strip()

def _to_float_str(x: str) -> str:
    """
    Convert:
      "34-1/4" -> "34.25"
      "17-3/4" -> "17.75"
      "59 1/4" -> "59.25"
      "1/4" -> "0.25"
      "5.25" -> "5.25"
      "12" -> "12"
    """
    s = (x or "").strip()
    if not s:
        return ""

    s = re.sub(r"^(\d+)\s*-\s*(\d+)\s*/\s*(\d+)$", r"\1 \2/\3", s)

    m = re.match(r"^(\d+)\s+(\d+)\s*/\s*(\d+)$", s)
    if m:
        whole = int(m.group(1))
        num = int(m.group(2))
        den = int(m.group(3))
        return str(round(whole + (num / den), 4)).rstrip("0").rstrip(".")

    m = re.match(r"^(\d+)\s*/\s*(\d+)$", s)
    if m:
        num = int(m.group(1))
        den = int(m.group(2))
        return str(round(num / den, 4)).rstrip("0").rstrip(".")

    if re.match(r"^\d+(\.\d+)?$", s):
        return s

    return ""

def _find_weight(dim_text: str) -> str:
    t = dim_text or ""
    m = re.search(r"(?:item\s*)?weight\s*[:\-\s]*([0-9]+(?:\.[0-9]+)?)", t, flags=re.I)
    if m:
        return m.group(1)
    m2 = re.search(r"\bfixture\s*weight\b[^0-9]*([0-9]+(?:\.[0-9]+)?)", t, flags=re.I)
    return m2.group(1) if m2 else ""

def _kv_value(t: str, keys) -> str:
    num = r'([0-9]+(?:\.[0-9]+)?(?:\s+[0-9]+/\d+)?(?:-[0-9]+/\d+)?)'
    for k in keys:
        m = re.search(rf'\b{k}\b\s*:\s*{num}', t, flags=re.I)
        if m:
            return _to_float_str(m.group(1))
    return ""

def _kv_raw_value(t: str, keys) -> str:
    num = r'([0-9]+(?:\.[0-9]+)?(?:\s+[0-9]+/\d+)?(?:-[0-9]+/\d+)?)'
    for k in keys:
        m = re.search(rf'\b{k}\b\s*:\s*{num}', t, flags=re.I)
        if m:
            return clean_space(m.group(1))
    return ""

def _collect_shade_details(t: str) -> str:
    keys = [
        "Fixture Width with Shade",
        "Length / Height with Shade",
        "Length / Height without Shade",
        "Projection with Shade",
        "Fixture Width W/ Shade",
        "Fixture Width W/o Shade",
        "Projection W/ Shade",
    ]
    values = []
    for k in keys:
        v = _kv_raw_value(t, [k])
        if v:
            values.append(v.replace('"', '').strip())

    m = re.search(
        r"\bLength\s*/?\s*Height\s*W\s*/\s*Shade\b\s*:\s*(L\s*:\s*[^A-Za-z]*[0-9][^A-Za-z]*\"\s*H\s*:\s*[0-9][^A-Za-z]*\")",
        t, flags=re.I
    )
    if m:
        values.append(clean_space(m.group(1)).replace("”", '"').replace("“", '"'))

    seen = set()
    out = []
    for v in values:
        vv = clean_space(v).strip().strip(",")
        if vv and vv not in seen:
            seen.add(vv)
            out.append(vv)
    return ", ".join(out)

def _extract_wattage(t: str) -> str:
    parts = []
    m = re.search(r"\bmax\.?\s*wattage\s*(?:per\s*socket)?\b\s*:\s*([0-9]+(?:\.[0-9]+)?)\s*W\b", t, flags=re.I)
    if m:
        parts.append(f"{m.group(1)}W")
    else:
        m2 = re.search(r"\bmax\s*wattage\b[^0-9]*([0-9]+(?:\.[0-9]+)?)\s*W\b", t, flags=re.I)
        if m2:
            parts.append(f"{m2.group(1)}W")

    me = re.search(r"\benergy\s*efficiency\b\s*:\s*([0-9]+(?:\.[0-9]+)?)\s*CFM\s*/\s*Watt\b", t, flags=re.I)
    if me:
        parts.append(f"{me.group(1)} CFM/Watt")

    out = []
    seen = set()
    for p in parts:
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return ", ".join(out)

def _extract_socket(t: str) -> str:
    sock = ""
    num = ""
    m = re.search(r"\bsocket\s*type(?:/integrated\s*led)?\b\s*:\s*([A-Za-z0-9\-_/]+)", t, flags=re.I)
    if m:
        sock = clean_space(m.group(1)).strip(",")
    m2 = re.search(r"\bnumber\s*of\s*sockets\b\s*:\s*([0-9]+)", t, flags=re.I)
    if m2:
        num = m2.group(1)
    if sock and num:
        return f"{sock}, {num}"
    return sock or num or ""

def _extract_canopy(t: str) -> str:
    v = _kv_raw_value(t, ["Canopy/Base Diameter", "Canopy / Base Diameter", "Canopy Diameter", "Base Diameter"])
    return v.replace('"', '').strip() if v else ""

def parse_dimension_fields(dimension_text: str) -> dict:
    t = _norm_dim_text(dimension_text or "")

    result = {
        "Description": "",
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Height": "",
        "Seat Width": "",
        "Seat Height": "",
        "Seat Depth": "",
        "Arm Height": "",
        "Shade Details": "",
        "Wattage": "",
        "Socket": "",
        "Canopy": "",
        "Length": "",  # New field for Length
        "Arm Depth": "",  # New field for Arm Depth
        "Arm Width": ""   # New field for Arm Width
    }

    if not t:
        return result

    result["Weight"] = _find_weight(t)

    # KV style Width/Height/Depth/Diameter
    w = _kv_value(t, ["Width", "Fixture Width", "Fixture Width with Shade", "Fixture Width W/ Shade"])
    h = _kv_value(t, ["Height", "Length/Height", "Length / Height", "Length / Height with Shade", "Length/Height W/ Shade"])
    d = _kv_value(t, ["Depth", "Projection", "Projection with Shade", "Projection W/ Shade"])
    dia = _kv_value(t, ["Overall Diameter", "Diameter", "Canopy/Base Diameter", "Canopy Diameter", "Base Diameter"])

    length = _kv_value(t, ["Length", "Overall Length"])  # Extract Length
    arm_depth = _kv_value(t, ["Arm Depth", "Arm Length"])  # Extract Arm Depth
    arm_width = _kv_value(t, ["Arm Width", "Width with Arms"])  # Extract Arm Width

    if w: result["Width"] = w
    if h: result["Height"] = h
    if d: result["Depth"] = d
    if dia: result["Diameter"] = dia
    if length: result["Length"] = length  # Add Length
    if arm_depth: result["Arm Depth"] = arm_depth  # Add Arm Depth
    if arm_width: result["Arm Width"] = arm_width  # Add Arm Width

    # other extracted fields
    result["Shade Details"] = _collect_shade_details(t)
    result["Wattage"] = _extract_wattage(t)
    result["Socket"] = _extract_socket(t)
    result["Canopy"] = _extract_canopy(t)

    # Seat/Arm lines
    msh = re.search(r"\bseat\s*height\b\s*[:\-\s]*([0-9]+(?:\.[0-9]+)?(?:\s+[0-9]+/\d+)?(?:-[0-9]+/\d+)?)", t, flags=re.I)
    if msh:
        result["Seat Height"] = _to_float_str(msh.group(1))
    msw = re.search(r"\bseat\s*width\b\s*[:\-\s]*([0-9]+(?:\.[0-9]+)?(?:\s+[0-9]+/\d+)?(?:-[0-9]+/\d+)?)", t, flags=re.I)
    if msw:
        result["Seat Width"] = _to_float_str(msw.group(1))
    msd = re.search(r"\bseat\s*depth\b\s*[:\-\s]*([0-9]+(?:\.[0-9]+)?(?:\s+[0-9]+/\d+)?(?:-[0-9]+/\d+)?)", t, flags=re.I)
    if msd:
        result["Seat Depth"] = _to_float_str(msd.group(1))

    mah = re.search(r"\barm\s*height\b\s*[:\-\s]*([0-9]+(?:\.[0-9]+)?(?:\s+[0-9]+/\d+)?(?:-[0-9]+/\d+)?)", t, flags=re.I)
    if mah:
        result["Arm Height"] = _to_float_str(mah.group(1))

    return result


# =========================
# FINISH
# =========================
def extract_finish_from_legends(driver) -> str:
    try:
        legends = driver.find_elements(By.CSS_SELECTOR, 'legend[data-test-id^="prompt-for-product-attribute-"]')
    except Exception:
        legends = []

    items = []
    for leg in legends:
        try:
            full = clean_space(leg.text)
            label = full.split(":", 1)[0].strip() if full else ""
            label_l = label.lower()

            spans = leg.find_elements(By.CSS_SELECTOR, "span")
            sel_val = ""
            for sp in spans[::-1]:
                txt = clean_space(sp.text)
                if txt:
                    sel_val = txt
                    break

            if sel_val:
                items.append((label_l, sel_val))
        except Exception:
            continue

    if not items:
        return ""

    for label_l, val in items:
        if "wood finish" in label_l:
            return val
    for label_l, val in items:
        if "finish" in label_l:
            return val
    for label_l, val in items:
        if "material" in label_l:
            return val

    return items[-1][1]

def extract_finish(driver) -> str:
    v = extract_finish_from_legends(driver)
    if v:
        return v

    t = get_text_safe(driver, '[data-test-id="guided-accordion-header-attribute-value"]', timeout=1)
    if t:
        return t

    try:
        checked = driver.find_elements(By.CSS_SELECTOR, "ul[data-test-id^='guided-pip-step-'] input[type='radio']:checked")
        if checked:
            _id = checked[0].get_attribute("id") or ""
            if _id:
                lab = driver.find_elements(By.CSS_SELECTOR, f"label#accordion-label-{_id}")
                if lab:
                    t2 = clean_space(lab[0].text)
                    if t2:
                        return t2
    except Exception:
        pass

    return ""


# =========================
# DESCRIPTION (sanitized, not exported)
# =========================
def _jsonld_description(driver) -> str:
    try:
        scripts = driver.find_elements(By.CSS_SELECTOR, 'script[type="application/ld+json"]')
        for sc in scripts:
            raw = (sc.get_attribute("innerText") or "").strip()
            if not raw:
                continue
            try:
                data = json.loads(raw)
            except Exception:
                continue

            def walk(obj):
                if isinstance(obj, dict):
                    if "description" in obj and isinstance(obj["description"], str) and obj["description"].strip():
                        return html_to_text(obj["description"])
                    for v in obj.values():
                        r = walk(v)
                        if r:
                            return r
                elif isinstance(obj, list):
                    for it in obj:
                        r = walk(it)
                        if r:
                            return r
                return ""

            got = walk(data)
            if got:
                return got
    except Exception:
        pass
    return ""

def extract_overview_description(driver) -> str:
    targets = {"overview", "details", "description"}

    click_if_exists(driver, '[data-accordion-list="product"][data-accordion-item="0product"] button', timeout=2)
    time.sleep(0.25)
    try:
        els = driver.find_elements(By.CSS_SELECTOR, '#accordion-panel-0product span.rich-text')
        if els:
            txt = html_to_text(clean_space(els[0].text))
            if txt:
                return txt
    except Exception:
        pass

    try:
        items = driver.find_elements(By.CSS_SELECTOR, '[data-accordion-list="product"] [data-accordion-item]')
        for it in items:
            head = it.find_elements(By.CSS_SELECTOR, '[data-test-id="product-accordion-heading"]')
            if not head:
                continue
            h = clean_space(head[0].text).lower()
            if h in targets:
                btn = it.find_elements(By.CSS_SELECTOR, "button")
                if btn:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn[0])
                    time.sleep(0.2)
                    try:
                        btn[0].click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn[0])
                    time.sleep(0.35)

                rt = it.find_elements(By.CSS_SELECTOR, "span.rich-text")
                if rt:
                    txt = html_to_text(clean_space(rt[0].text))
                    if txt:
                        return txt
    except Exception:
        pass

    j = _jsonld_description(driver)
    if j:
        return j

    try:
        meta = driver.find_elements(By.CSS_SELECTOR, 'meta[name="description"]')
        if meta:
            c = html_to_text(meta[0].get_attribute("content") or "")
            if c:
                return c
    except Exception:
        pass

    return ""


# =========================
# DIMENSIONS / SPECS (RIGDON FIX)
# =========================
def extract_dimensions_text(driver) -> str:
    """
    Robust Dimensions/Specs extraction:
    - Tries accordion "Dimensions"
    - If missing, tries "Specifications" / "Product Specifications"
    - Final fallback: collects spec-like lines from visible page text
    """
    def looks_like_specs(txt: str) -> bool:
        t = (txt or "").lower()
        keys = [
            "width", "height", "depth", "diameter", "overall",
            "fixture width", "length / height", "projection",
            "canopy/base", "fixture weight", "socket type",
            "number of sockets", "max. wattage", "energy efficiency",
        ]
        return any(k in t for k in keys)

    # 1) try common accordion slot
    try:
        click_if_exists(driver, '[data-accordion-list="product"][data-accordion-item="1product"] button', timeout=2)
        time.sleep(0.25)
        els = driver.find_elements(By.CSS_SELECTOR, '#accordion-panel-1product span.rich-text')
        if els:
            txt = clean_space(els[0].text)
            if txt and looks_like_specs(txt):
                return txt
    except Exception:
        pass

    # 2) any accordion item named Dimensions/Specifications
    targets = {"dimensions", "specifications", "product specifications"}
    try:
        items = driver.find_elements(By.CSS_SELECTOR, '[data-accordion-list="product"] [data-accordion-item]')
        for it in items:
            head = it.find_elements(By.CSS_SELECTOR, '[data-test-id="product-accordion-heading"]')
            if not head:
                continue
            h = clean_space(head[0].text).lower()
            if h not in targets:
                continue

            btn = it.find_elements(By.CSS_SELECTOR, "button")
            if btn:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn[0])
                time.sleep(0.2)
                try:
                    btn[0].click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn[0])
                time.sleep(0.35)

            panel_id = ""
            try:
                panel_id = (btn[0].get_attribute("aria-controls") or "").strip() if btn else ""
            except Exception:
                panel_id = ""

            if panel_id:
                panel = driver.find_elements(By.CSS_SELECTOR, f"#{panel_id}")
                if panel:
                    txt = clean_space(panel[0].text)
                    if txt and looks_like_specs(txt):
                        return txt

            txt2 = clean_space(it.text)
            if txt2 and looks_like_specs(txt2):
                return txt2
    except Exception:
        pass

    # 3) find any element containing "specifications" or "dimensions"
    try:
        nodes = driver.find_elements(
            By.XPATH,
            "//*[self::button or self::h2 or self::h3 or self::span]"
            "[contains(translate(normalize-space(.),"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'specifications')"
            " or contains(translate(normalize-space(.),"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'dimensions')]"
        )
        for node in nodes[:8]:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", node)
                time.sleep(0.2)
                try:
                    node.click()
                    time.sleep(0.25)
                except Exception:
                    pass

                container = driver.execute_script(
                    "let el=arguments[0];"
                    "for(let i=0;i<7;i++){"
                    " if(!el) break;"
                    " if(el.innerText && el.innerText.length>40) return el;"
                    " el=el.parentElement;"
                    "}"
                    "return arguments[0];",
                    node
                )
                txt = clean_space(container.text if hasattr(container, "text") else "")
                if txt and looks_like_specs(txt):
                    return txt
            except Exception:
                continue
    except Exception:
        pass

    # 4) final fallback: scan body and keep only spec-like lines
    try:
        body_text = clean_space(driver.find_element(By.TAG_NAME, "body").text)
        if body_text:
            lines = [ln.strip() for ln in re.split(r"[\r\n]+", body_text) if ln.strip()]
            pick = []
            for ln in lines:
                low = ln.lower()
                if any(k in low for k in [
                    "fixture width", "length / height", "length/height",
                    "projection", "canopy/base", "fixture weight",
                    "socket type", "number of sockets",
                    "max. wattage", "energy efficiency",
                    "width:", "height:", "depth:", "diameter:"
                ]):
                    pick.append(ln)
            if pick:
                return "\n".join(pick[:80])
    except Exception:
        pass

    return ""


# =========================
# VARIATIONS + IMAGE
# =========================
def extract_product_family_id(product_name: str) -> str:
    n = clean_space(product_name).replace(",", " ")
    parts = [p for p in n.split(" ") if p.strip()]
    return " ".join(parts[:3]) if parts else ""

def get_variation_radio_inputs(driver):
    radios = []
    seen = set()
    try:
        inputs = driver.find_elements(By.CSS_SELECTOR, 'ul[data-test-id="product-attributes"] input[type="radio"]')
        for inp in inputs:
            _id = inp.get_attribute("id") or ""
            if _id and _id not in seen:
                seen.add(_id)
                radios.append(inp)
    except Exception:
        pass
    return radios

def click_radio_by_id(driver, rid: str) -> bool:
    try:
        lab = driver.find_elements(By.CSS_SELECTOR, f'label[for="{rid}"]')
        if lab:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", lab[0])
            time.sleep(0.2)
            try:
                lab[0].click()
            except Exception:
                driver.execute_script("arguments[0].click();", lab[0])
            return True
    except Exception:
        pass
    return False

def get_main_image_url(driver) -> str:
    try:
        og = driver.find_elements(By.CSS_SELECTOR, 'meta[property="og:image"]')
        if og:
            return clean_space(og[0].get_attribute("content") or "")
    except Exception:
        pass

    try:
        imgs = driver.find_elements(By.CSS_SELECTOR, 'img')
        for im in imgs:
            src = (im.get_attribute("src") or "").strip()
            if src and ("assets.rjimgs.com" in src or "rjimgs" in src):
                return src
    except Exception:
        pass

    return ""


# =========================
# DRIVER
# =========================
def connect_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    return webdriver.Chrome(options=opts)


# =========================
# OUTPUT COLS (EXACT ORDER)
# =========================
OUT_COLS = [
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
    "Height",
    "List Price",
    "Finish",
    "Dimension",
    "Seat Width",
    "Seat Height",
    "Seat Depth",
    "Arm Height",
    "Shade Details",
    "Wattage",
    "Socket",
    "Canopy",
    "Length",  # New column
    "Arm Depth",  # New column
    "Arm Width"   # New column
]


def save_now(out_rows):
    if not out_rows:
        return
    df_out = pd.DataFrame(out_rows, columns=OUT_COLS)
    df_out.to_excel(OUTPUT_XLSX, index=False)
    print(f"✅ Auto-saved rows: {len(df_out)} -> {OUTPUT_XLSX}")


# =========================
# MAIN
# =========================
def scrape_details_from_excel():
    df_in = pd.read_excel(INPUT_XLSX)
    if "Product URL" not in df_in.columns:
        raise Exception("Input Excel must contain 'Product URL' column")

    driver = connect_driver()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    out_rows = []
    global_index = 0
    base_product_count = 0

    try:
        for _, row in df_in.iterrows():
            base_url = str(row.get("Product URL") or "").strip()
            if not base_url:
                continue

            base_img = str(row.get("Image URL") or "").strip()
            product_name = clean_space(str(row.get("Product Name") or ""))
            list_price = clean_space(str(row.get("List Price") or ""))

            global_index += 1
            base_product_count += 1

            driver.get(base_url)
            time.sleep(POLITE_DELAY)

            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-test-id="product-title"], h1')))
            except Exception:
                pass

            title = get_text_safe(driver, '[data-test-id="product-title"]', timeout=1)
            if title:
                product_name = title

            product_family_id = extract_product_family_id(product_name)

            # keep (sanitized) description internally (not exported)
            desc = html_to_text(extract_overview_description(driver))

            dimension = extract_dimensions_text(driver)
            dim_fields = parse_dimension_fields(dimension)

            radios = get_variation_radio_inputs(driver)

            def build_row(sku_val, finish_val, img_val):
                return {
                    "Product URL": base_url,
                    "Image URL": img_val,
                    "Product Name": product_name,
                    "SKU": sku_val,
                    "Product Family Id": product_family_id,
                    "Description": desc,
                    "Weight": dim_fields["Weight"],
                    "Width": dim_fields["Width"],
                    "Depth": dim_fields["Depth"],
                    "Diameter": dim_fields["Diameter"],
                    "Height": dim_fields["Height"],
                    "List Price": list_price,
                    "Finish": finish_val,
                    "Dimension": dimension,
                    "Seat Width": dim_fields["Seat Width"],
                    "Seat Height": dim_fields["Seat Height"],
                    "Seat Depth": dim_fields["Seat Depth"],
                    "Arm Height": dim_fields["Arm Height"],
                    "Shade Details": dim_fields["Shade Details"],
                    "Wattage": dim_fields["Wattage"],
                    "Socket": dim_fields["Socket"],
                    "Canopy": dim_fields["Canopy"],
                    "Length": dim_fields["Length"],  # New field
                    "Arm Depth": dim_fields["Arm Depth"],  # New field
                    "Arm Width": dim_fields["Arm Width"],  # New field
                }

            if not radios:
                sku = extract_current_sku(driver) or sku_fallback(VENDOR_NAME, CATEGORY_NAME, global_index)
                finish = extract_finish(driver)
                img = get_main_image_url(driver) or base_img
                out_rows.append(build_row(sku, finish, img))
            else:
                for vi, inp in enumerate(radios, start=1):
                    rid = inp.get_attribute("id") or ""
                    if not rid:
                        continue

                    click_radio_by_id(driver, rid)
                    time.sleep(0.8)

                    sku = extract_current_sku(driver)
                    if not sku:
                        sku = sku_fallback(VENDOR_NAME, CATEGORY_NAME, int(f"{global_index}{vi}"))

                    finish = extract_finish(driver)
                    img = get_main_image_url(driver) or base_img
                    out_rows.append(build_row(sku, finish, img))

            if base_product_count % AUTOSAVE_EVERY == 0:
                save_now(out_rows)

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    save_now(out_rows)
    print(f"Done! Rows: {len(out_rows)}")
    print(f"Saved: {OUTPUT_XLSX}")


if __name__ == "__main__":
    scrape_details_from_excel()
