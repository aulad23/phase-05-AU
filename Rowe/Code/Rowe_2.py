# -*- coding: utf-8 -*-
import os, re, time
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ================== CONFIG ==================

BASE_URL = "https://rowefurniture.com"

# Script folder
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Input/Output (same folder)
INPUT_FILE_NAME  = "rowefurniture_Consoles.xlsx"
OUTPUT_FILE_NAME = "rowefurniture_Consoles_veriation.xlsx"

INPUT_FILE  = os.path.join(SCRIPT_DIR, INPUT_FILE_NAME)
OUTPUT_FILE = os.path.join(SCRIPT_DIR, OUTPUT_FILE_NAME)

HEADLESS = False          # True korle browser show hobe na
WAIT_TIMEOUT = 15
PAUSE_BETWEEN = 1.0

# ============================================

def make_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1400,1000")
    return webdriver.Chrome(options=opts)

def clean_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

def find_text_in_table_by_label(container: BeautifulSoup, label: str) -> str | None:
    if not container:
        return None
    target = label.lower().strip()
    for td in container.find_all("td"):
        if clean_text(td.get_text(" ")).lower() == target:
            sib = td.find_next_sibling("td")
            if sib:
                return clean_text(sib.get_text(" "))
    return None

def parse_description(soup: BeautifulSoup) -> str:
    tab = soup.find("div", id="quickTab-description")
    if tab:
        box = tab.find("div", class_="full-description")
        if box:
            text = clean_text(box.get_text(" "))
            if text:
                return text

    box = soup.find("div", class_="full-description")
    if box:
        text = clean_text(box.get_text(" "))
        if text:
            return text

    sbox = soup.find("div", class_="short-description")
    if sbox:
        text = clean_text(sbox.get_text(" "))
        if text:
            return text

    return ""

def parse_specs_block(soup: BeautifulSoup) -> BeautifulSoup | None:
    spec = soup.find("div", id="quickTab-specifications")
    if spec:
        return spec
    for tab in soup.select("div.productTabs-tab"):
        title = tab.find("div", class_="ui-tab-title")
        if title and "specification" in title.get_text(" ").lower():
            body = tab.find("div", class_="ui-tab-body")
            if body:
                return body
    return None

def extract_specs_text(spec: BeautifulSoup) -> str:
    if not spec:
        return ""

    parts = []
    for box in spec.select("div.product-specs-box"):
        for title_div in box.find_all("div", class_="title"):
            title_txt = clean_text(title_div.get_text(" "))
            if title_txt:
                parts.append(f"[{title_txt}]")

        for tr in box.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) >= 2:
                label = clean_text(tds[0].get_text(" "))
                value = clean_text(tds[1].get_text(" "))
                if label or value:
                    if value:
                        parts.append(f"{label}: {value}")
                    else:
                        parts.append(label)

    return "\n".join(parts)

def get_sku_from_page(soup: BeautifulSoup) -> str:
    # 1) <span class="label">SKU:</span> ... <span class="value" id="sku-XXXX">RR-...</span>
    lbl = soup.find("span", class_="label", string=re.compile(r"\bSKU\b", re.I))
    if lbl:
        parent = lbl.parent
        val = parent.find("span", class_="value")
        if val:
            txt = clean_text(val.get_text(" "))
            if txt:
                return txt

    # 2) any span id^="sku-"
    span_val = soup.find("span", id=re.compile(r"^sku-\d+"))
    if span_val:
        txt = clean_text(span_val.get_text(" "))
        if txt:
            return txt

    # 3) fallback: specifications table
    spec = parse_specs_block(soup)
    if spec:
        txt = find_text_in_table_by_label(spec, "SKU")
        if txt:
            return txt

    return ""

def get_main_image_from_page(soup: BeautifulSoup, fallback: str = "") -> str:
    """
    Main hero image:
      <img id="cloudZoomImage" src="..._1170.jpeg">
    """
    img = soup.find("img", id="cloudZoomImage")
    if img and img.get("src"):
        return img["src"].strip()

    wrapper = soup.find("div", class_="picture-wrapper")
    if wrapper:
        img2 = wrapper.find("img")
        if img2 and img2.get("src"):
            return img2["src"].strip()

    return fallback

def get_product_name_from_page(soup: BeautifulSoup, fallback: str = "") -> str:
    h1 = soup.find(["h1", "h2"], class_="product-name")
    if h1:
        txt = clean_text(h1.get_text(" "))
        if txt:
            return txt
    return fallback

def build_single_row_from_soup(base_row: dict, soup: BeautifulSoup) -> dict:
    p_name = get_product_name_from_page(soup, base_row.get("Product Name", ""))
    sku    = get_sku_from_page(soup) or base_row.get("SKU", "")
    img    = get_main_image_from_page(soup, base_row.get("Image URL", ""))

    description = parse_description(soup)
    spec_block  = parse_specs_block(soup)
    specs_text  = extract_specs_text(spec_block)

    return {
        "Product URL": base_row.get("Product URL", ""),
        "Image URL": img,
        "Product Name": p_name,
        "SKU": sku,
        "Finish": "",
        "Description": description,
        "Specifications": specs_text,
    }

# ============================
# ONLY "CHOOSE FINISH" VARIATIONS + CHANGE CHECK
# ============================

def scrape_variation_rows_with_driver(driver, url: str, base_row: dict) -> list[dict]:
    try:
        driver.get(url)
    except Exception as e:
        print(f"   ⚠️ driver.get error: {e}")
        return [{
            "Product URL": base_row.get("Product URL",""),
            "Image URL": base_row.get("Image URL",""),
            "Product Name": base_row.get("Product Name",""),
            "SKU": base_row.get("SKU",""),
            "Finish": "",
            "Description": "",
            "Specifications": "",
        }]

    # wait for product title
    try:
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1.product-name, h1.product-title"))
        )
    except Exception:
        pass

    # ---- baseline (before kono finish click) ----
    html0 = driver.page_source
    soup0 = BeautifulSoup(html0, "html.parser")
    base_img0 = get_main_image_from_page(soup0, base_row.get("Image URL", ""))
    base_sku0 = get_sku_from_page(soup0) or base_row.get("SKU", "")
    base_desc0 = parse_description(soup0)
    base_spec0 = extract_specs_text(parse_specs_block(soup0))
    baseline_signature = (base_img0, base_sku0, base_desc0, base_spec0)

    # -----------------------------
    # 1) "Choose Finish" accordion open korar try
    # -----------------------------
    try:
        finish_dt = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'attributes')]//dt[.//label[contains(normalize-space(),'Choose Finish')]]"
        )
        driver.execute_script("arguments[0].click();", finish_dt)
        time.sleep(0.8)
    except Exception:
        # jodi na pai, pore sudhu UL diye check korbo
        pass

    # -----------------------------
    # 2) Sudhu oi UL nibo jekhane data-att-textprompt="Choose Finish"
    # -----------------------------
    try:
        ul_elem = driver.find_element(
            By.CSS_SELECTOR,
            "div.attributes ul.attribute-squares[data-att-textprompt*='Choose Finish']"
        )
        li_elems = ul_elem.find_elements(By.TAG_NAME, "li")
        has_variation = len(li_elems) > 0
    except Exception:
        has_variation = False
        li_elems = []

    rows = []

    # Jodi "Choose Finish" er UL/LI na pai → single row (baseline)
    if not has_variation:
        row = build_single_row_from_soup(base_row, soup0)
        rows.append(row)
        return rows

    # -----------------------------
    # 3) Variation loop – sudhu "Choose Finish" li gula
    #    & sudhu jokhon data change hoy tokhon row nibo
    # -----------------------------
    seen_signatures = set()   # joto variation already new data ditaese

    for li in li_elems:
        try:
            # Finish text (Sesame / Godiva ...)
            try:
                label_p = li.find_element(By.CSS_SELECTOR, "p.attribute-image-ptag")
                finish_text = clean_text(label_p.text)
            except Exception:
                finish_text = ""

            # current hero image src (before click) – ikkhn sudhu info, use kortei pari
            try:
                hero_img_elem = driver.find_element(By.ID, "cloudZoomImage")
                old_src = hero_img_elem.get_attribute("src") or ""
            except Exception:
                old_src = ""

            # Finish radio click
            try:
                radio = li.find_element(By.CSS_SELECTOR, "input[type='radio']")
                driver.execute_script("arguments[0].scrollIntoView(true);", radio)
                driver.execute_script("arguments[0].click();", radio)
            except Exception as e:
                print(f"   ⚠️ radio click error: {e}")
                continue

            # wait hero image src change (optional hint)
            if old_src:
                try:
                    WebDriverWait(driver, 5).until(
                        lambda drv: drv.find_element(By.ID, "cloudZoomImage").get_attribute("src") != old_src
                    )
                except TimeoutException:
                    # change na hole o thakte pare – niche signature diye final check korbo
                    pass

            time.sleep(0.5)  # small extra wait

            # updated HTML
            html = driver.page_source
            soup = BeautifulSoup(html, "html.parser")

            p_name   = get_product_name_from_page(soup, base_row.get("Product Name", ""))
            sku      = get_sku_from_page(soup) or base_row.get("SKU", "")
            main_img = get_main_image_from_page(soup, base_row.get("Image URL", ""))

            description = parse_description(soup)
            spec_block  = parse_specs_block(soup)
            specs_text  = extract_specs_text(spec_block)

            # ---- signature diye check: data change hoise naki ----
            new_signature = (main_img, sku, description, specs_text)

            # 1) Jodi baseline-er sathe same → ei finish actually kichu change koreni → skip
            if new_signature == baseline_signature:
                print("   ℹ️ Finish clicked but data unchanged (same as baseline) → skipping this finish.")
                continue

            # 2) Jodi age already ei same signature dhore rakha → duplicate variation → skip
            if new_signature in seen_signatures:
                print("   ℹ️ Duplicate variation (same data as another finish) → skipping.")
                continue

            # Notun data, r flavour – store signature
            seen_signatures.add(new_signature)

            # display name + finish
            if finish_text and p_name:
                display_name = f"{p_name} - {finish_text}"
            elif finish_text:
                display_name = finish_text
            else:
                display_name = p_name

            rows.append({
                "Product URL": base_row.get("Product URL", ""),
                "Image URL": main_img,
                "Product Name": display_name,
                "SKU": sku,
                "Finish": finish_text,
                "Description": description,
                "Specifications": specs_text,
            })

        except Exception as e:
            print(f"   ⚠️ variation loop error: {e}")
            rows.append({
                "Product URL": base_row.get("Product URL",""),
                "Image URL": base_row.get("Image URL",""),
                "Product Name": base_row.get("Product Name",""),
                "SKU": base_row.get("SKU",""),
                "Finish": "",
                "Description": "",
                "Specifications": "",
            })

    # Jodi kono variation-eo data change na hoy (rows now empty) → at least baseline ekta row
    if not rows:
        rows.append(build_single_row_from_soup(base_row, soup0))

    return rows

# ============================
# MAIN
# ============================

def main():
    print(f"🔎 Looking for input file:\n   {INPUT_FILE}\n")

    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input not found: {INPUT_FILE}")
        return

    df_in = pd.read_excel(INPUT_FILE)

    def norm_col(c):
        c = str(c)
        c = re.sub(r"\s+", " ", c).strip()
        return c

    df_in.columns = [norm_col(c) for c in df_in.columns]

    if "Product URL" not in df_in.columns:
        print(f"❌ INPUT Excel-এ 'Product URL' কলাম নেই। কলামগুলো: {list(df_in.columns)}")
        return

    driver = make_driver()
    all_rows = []
    total = len(df_in)

    try:
        for i, rec in enumerate(df_in.itertuples(index=False, name=None), 1):
            row_values = dict(zip(df_in.columns, rec))

            base_row = {
                "Product URL": row_values.get("Product URL", "") or "",
                "Image URL": row_values.get("Image URL", "") if "Image URL" in df_in.columns else "",
                "Product Name": row_values.get("Product Name", "") if "Product Name" in df_in.columns else "",
                "SKU": row_values.get("SKU", "") if "SKU" in df_in.columns else "",
            }

            url = base_row["Product URL"]
            if not url:
                print(f"[{i}/{total}] ⚠️ Empty URL, skipping")
                continue

            print(f"[{i}/{total}] {url}")
            try:
                var_rows = scrape_variation_rows_with_driver(driver, url, base_row)
            except Exception as e:
                print(f"   ⚠️ Error: {e}")
                var_rows = [{
                    "Product URL": base_row.get("Product URL",""),
                    "Image URL": base_row.get("Image URL",""),
                    "Product Name": base_row.get("Product Name",""),
                    "SKU": base_row.get("SKU",""),
                    "Finish": "",
                    "Description": "",
                    "Specifications": "",
                }]

            all_rows.extend(var_rows)
            time.sleep(PAUSE_BETWEEN)
    finally:
        driver.quit()

    out_cols = [
        "Product URL",
        "Image URL",
        "Product Name",
        "SKU",
        "Finish",
        "Description",
        "Specifications",
    ]
    pd.DataFrame(all_rows, columns=out_cols).to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Saved variation output:\n   {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
