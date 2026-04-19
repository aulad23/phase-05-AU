from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re
import json

# ================== Paths ==================
INPUT_FILE = "berhardt missing.xlsx"
OUTPUT_FILE = "berhardt missing_final.xlsx"

# ================== Load Excel ==================
df = pd.read_excel(INPUT_FILE)

required_cols = ["Product URL", "Image URL", "Product Name", "SKU", "Product Family ID",
                 "Description", "Weight", "Width", "Depth", "Diameter", "Height", "Finish",
                 "Fabric Grade", "Content Details", "Cleaning Code"]

# Add missing columns
for col in required_cols:
    if col not in df.columns:
        df[col] = ""

# Always refresh image URLs
df["Image URL"] = ""

# ================== Selenium Setup ==================
def make_driver(headless=True):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 20.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    driver = webdriver.Chrome(options=opts)
    driver.implicitly_wait(0.5)
    return driver

driver = make_driver(headless=False)
wait = WebDriverWait(driver, 12)

# ================== Helpers ==================
PAT_W   = re.compile(r'W:\s*([\d.,]+(?:\s*(?:in|cm|mm))?)',  flags=re.I)
PAT_D   = re.compile(r'D:\s*([\d.,]+(?:\s*(?:in|cm|mm))?)',  flags=re.I)
PAT_H   = re.compile(r'H:\s*([\d.,]+(?:\s*(?:in|cm|mm))?)',  flags=re.I)
PAT_DIA = re.compile(r'Dia\.?:\s*([\d.,]+(?:\s*(?:in|cm|mm))?)', flags=re.I)

def normalize_ws(s: str) -> str:
    if s is None:
        return ""
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    return re.sub(r"[ \t]+", " ", s).strip()

def safe_match(pattern, text):
    m = pattern.search(text)
    if m:
        val = m.group(1).strip()
        if re.search(r"\d", val):
            return val[:23]
    return ""

def _first_in_value(elems):
    for el in elems:
        t = (el.text or "").strip()
        if " in" in t.lower():
            return t
    for el in elems:
        t = (el.text or "").strip()
        if re.search(r"\d", t):
            return t
    return ""

def scrape_dimensions_panel_all(driver, wait):
    results = {"Width": "", "Depth": "", "Height": "", "Diameter": ""}
    try:
        try:
            btn = driver.find_element(By.XPATH, "//panel[@panel-title='Dimensions']//button")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(0.3)
        except Exception:
            pass

        panel_body = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//panel[@panel-title='Dimensions']//div[contains(@class,'panel-body')]"
            ))
        )

        for ul in panel_body.find_elements(By.CSS_SELECTOR, "ul.no-bullets"):
            items = ul.find_elements(By.TAG_NAME, "li")
            if not items:
                continue
            label = (items[0].text or "").strip().upper()
            val = _first_in_value(items[1:])
            val = normalize_ws(val)
            if not val:
                continue

            if "WIDTH" in label:
                results["Width"] = val
            elif "DEPTH" in label or "PRODUCT DEPTH" in label:
                results["Depth"] = val
            elif "HEIGHT" in label:
                results["Height"] = val
            elif "DIAMETER" in label or "DIA" in label:
                results["Diameter"] = val
    except Exception:
        pass
    return results

def get_image_url(driver) -> str:
    selectors = [
        "img.main-product-image",
        "img.primary-image",
        "img.gallery-main-image",
        "img.product-main-image",
        "div.product-image img",
        "div.product-gallery img.active",
        "div.carousel img.active",
        "div.carousel img.slick-current",
        "div.slick-slide.slick-current img",
        "picture img",
    ]
    for sel in selectors:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            src = el.get_attribute("src") or el.get_attribute("data-src") or el.get_attribute("srcset")
            if src:
                return src.strip()
        except Exception:
            pass

    gallery_selectors = [
        "ul.thumbnails img",
        "div.thumbnails img",
        "div.product-thumbnails img",
        "div.gallery-thumbs img",
        "div.slick-track img",
    ]
    for gsel in gallery_selectors:
        try:
            imgs = driver.find_elements(By.CSS_SELECTOR, gsel)
            for el in imgs:
                if el.get_attribute("class") and ("active" in el.get_attribute("class") or "selected" in el.get_attribute("class")):
                    src = el.get_attribute("data-large") or el.get_attribute("data-zoom") or el.get_attribute("src") or el.get_attribute("data-src")
                    if src:
                        return src.strip()
            if imgs:
                el = imgs[0]
                src = el.get_attribute("data-large") or el.get_attribute("data-zoom") or el.get_attribute("src") or el.get_attribute("data-src")
                if src:
                    return src.strip()
        except Exception:
            pass

    try:
        og = driver.find_element(By.XPATH, "//meta[@property='og:image' or @name='og:image']")
        content = og.get_attribute("content")
        if content:
            return content.strip()
    except Exception:
        pass

    try:
        link_el = driver.find_element(By.XPATH, "//link[@rel='image_src']")
        href = link_el.get_attribute("href")
        if href:
            return href.strip()
    except Exception:
        pass

    try:
        scripts = driver.find_elements(By.XPATH, "//script[@type='application/ld+json']")
        for sc in scripts:
            try:
                txt = sc.get_attribute("innerText") or sc.get_attribute("textContent") or ""
                if not txt.strip():
                    continue
                data = json.loads(txt)
                candidates = data if isinstance(data, list) else [data]
                for obj in candidates:
                    if not isinstance(obj, dict):
                        continue
                    typ = obj.get("@type") or obj.get("@graph") or ""
                    if isinstance(typ, str) and "Product" in typ:
                        img = obj.get("image")
                        if isinstance(img, str) and img.strip():
                            return img.strip()
                        if isinstance(img, list) and img:
                            return str(img[0]).strip()
                    if "@graph" in obj and isinstance(obj["@graph"], list):
                        for node in obj["@graph"]:
                            if isinstance(node, dict) and node.get("@type") == "Product":
                                img = node.get("image")
                                if isinstance(img, str) and img.strip():
                                    return img.strip()
                                if isinstance(img, list) and img:
                                    return str(img[0]).strip()
            except Exception:
                continue
    except Exception:
        pass

    try:
        imgs = driver.find_elements(By.TAG_NAME, "img")
        best = ""
        best_w = 0
        for im in imgs:
            src = im.get_attribute("src") or im.get_attribute("data-src") or ""
            if not src:
                continue
            w = im.get_attribute("width") or im.get_attribute("data-width") or ""
            h = im.get_attribute("height") or im.get_attribute("data-height") or ""
            try:
                wv = int(re.sub(r"\D", "", str(w))) if w else 0
                hv = int(re.sub(r"\D", "", str(h))) if h else 0
                size_score = max(wv, hv)
            except Exception:
                size_score = 0
            if re.search(r"(product|large|zoom|gallery|main)", src, flags=re.I):
                size_score += 500
            if size_score > best_w:
                best_w = size_score
                best = src
        if best:
            return best.strip()
    except Exception:
        pass

    return ""

# ================== Main Loop ==================
BATCH_SIZE = 5
try:
    for idx, row in df.iterrows():
        url = str(row.get("Product URL", "")).strip()
        if not url or url.lower().startswith("nan"):
            print(f"[-] Skipping row {idx+1}: empty Product URL")
            continue

        try:
            driver.get(url)
            time.sleep(1)

            # ====== DESCRIPTION ======
            description = ""
            try:
                desc_panel = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//panel[@panel-title='Description']"))
                )
                panel_body = desc_panel.find_element(By.XPATH, ".//div[contains(@class,'panel-body')]")
                description = normalize_ws(panel_body.get_attribute("innerText") or panel_body.text)
            except Exception:
                try:
                    desc_el = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "p.one-up-long-desc"))
                    )
                    description = normalize_ws(desc_el.text)
                except Exception:
                    description = ""
            df.at[idx, "Description"] = description

            # ====== WEIGHT ======
            weight = ""
            try:
                weight_el = driver.find_element(By.XPATH, "//div[contains(text(),'Weight')]/following-sibling::div")
                weight = normalize_ws(weight_el.text)
            except Exception:
                weight = ""
            df.at[idx, "Weight"] = weight

            # ====== IMAGE URL ======
            image_url = ""
            try:
                image_url = get_image_url(driver)
            except Exception:
                image_url = ""
            df.at[idx, "Image URL"] = image_url

            # ====== PRODUCT FAMILY ID ======
            product_family_id = ""
            try:
                name_el = driver.find_element(By.CSS_SELECTOR, "h1.product-name, div.product-title, div.one-up-title")
                product_family_id = normalize_ws(name_el.text)
            except Exception:
                product_family_id = normalize_ws(row.get("Product Name", ""))
            df.at[idx, "Product Family ID"] = product_family_id

            # ====== SKU ======
            sku = row.get("SKU", "")
            df.at[idx, "SKU"] = sku

            # ====== FINISH ======
            finish_list = []
            try:
                finish_elements = driver.find_elements(By.CSS_SELECTOR, "div.items.with-images div.text-center.ng-binding")
                for el in finish_elements:
                    text = normalize_ws(el.text)
                    if text:
                        finish_list.append(text)
            except Exception:
                pass
            df.at[idx, "Finish"] = ", ".join(finish_list)

            # ====== FABRIC DETAILS ======
            fabric_grade = ""
            content_details = ""
            cleaning_code = ""

            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.col-xs-12 div.fabric-specs")))
                specs_blocks = driver.find_elements(By.CSS_SELECTOR, "div.col-xs-12 div.fabric-specs div.ng-scope")

                for block in specs_blocks:
                    try:
                        label_el = block.find_element(By.CSS_SELECTOR, "span.text-regular, span.text-regular.ng-binding")
                        value_el = block.find_element(By.CSS_SELECTOR, "span.text-muted.ng-binding")
                        label = normalize_ws(label_el.text).lower()
                        value = normalize_ws(value_el.text)

                        if "fabric grade" in label:
                            fabric_grade = value
                        elif "content details" in label:
                            content_details = value
                        elif "cleaning code" in label:
                            cleaning_code = value
                    except Exception:
                        continue
            except Exception:
                pass

            df.at[idx, "Fabric Grade"] = fabric_grade
            df.at[idx, "Content Details"] = content_details
            df.at[idx, "Cleaning Code"] = cleaning_code

            # ====== DIMENSIONS ======
            dims = scrape_dimensions_panel_all(driver, wait)
            df.at[idx, "Width"] = dims["Width"]
            df.at[idx, "Depth"] = dims["Depth"]
            df.at[idx, "Height"] = dims["Height"]
            df.at[idx, "Diameter"] = dims["Diameter"]

            if not any(dims.values()):
                try:
                    dim_el = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.dimensions"))
                    )
                    dim_text = normalize_ws(dim_el.text)
                except Exception:
                    dim_text = normalize_ws(driver.page_source)

                if not df.at[idx, "Width"]:
                    df.at[idx, "Width"] = safe_match(PAT_W, dim_text)
                if not df.at[idx, "Depth"]:
                    df.at[idx, "Depth"] = safe_match(PAT_D, dim_text)
                if not df.at[idx, "Height"]:
                    df.at[idx, "Height"] = safe_match(PAT_H, dim_text)
                if not df.at[idx, "Diameter"]:
                    df.at[idx, "Diameter"] = safe_match(PAT_DIA, dim_text)

            print(
                f"[+] Updated {idx+1}/{len(df)}: {row.get('Product Name', '')} | "
                f"W:{df.at[idx,'Width']} D:{df.at[idx,'Depth']} "
                f"H:{df.at[idx,'Height']} Dia:{df.at[idx,'Diameter']} | "
                f"Weight:{weight} Img:{('yes' if image_url else 'no')} "
                f"Desc len:{len(description)} Finish:{df.at[idx, 'Finish']} "
                f"Fabric Grade:{fabric_grade} Content:{content_details} Cleaning:{cleaning_code}"
            )

        except Exception as e:
            print(f"[!] Error on row {idx+1}: {e}")
            for c in required_cols:
                val = df.at[idx, c]
                if pd.isna(val) or val is None:
                    df.at[idx, c] = ""

        if (idx + 1) % BATCH_SIZE == 0:
            df[required_cols].to_excel(OUTPUT_FILE, index=False)
            print(f"[✓] Batch {idx+1//BATCH_SIZE} saved to {OUTPUT_FILE}")

finally:
    try:
        driver.quit()
    except Exception:
        pass

    df[required_cols].to_excel(OUTPUT_FILE, index=False)
    print(f"[✓] Done! Final updated file saved at {OUTPUT_FILE}")
