from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from pathlib import Path
import os
import re
import json

# ================== Paths ==================
INPUT_FILE = "bernhardt_Desk Chairs.xlsx"
OUTPUT_FILE = "bernhardt_Desk_Chairs_Final.xlsx"

# ================== Load Excel ==================
df = pd.read_excel(INPUT_FILE)

required_cols = ["Product URL", "Image URL", "Product Name", "SKU", "Product Family ID",
                 "Description", "Weight", "Width", "Depth", "Diameter", "Height", "Finish",
                 "Seat Width", "Seat Depth", "Seat Height",
                 "Arm Width", "Arm Depth", "Arm Height", "COM"]

# Add missing columns
for col in required_cols:
    if col not in df.columns:
        df[col] = ""

# Always clear existing Image URLs
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

# ================== DIMENSIONS PANEL ==================
def scrape_dimensions_panel_all(driver, wait):
    results = {
        "Width": "", "Depth": "", "Height": "", "Diameter": "",
        "Seat Width": "", "Seat Depth": "", "Seat Height": "",
        "Arm Width": "", "Arm Depth": "", "Arm Height": ""
    }
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

            if "SEAT WIDTH" in label:
                results["Seat Width"] = val
            elif "SEAT DEPTH" in label:
                results["Seat Depth"] = val
            elif "SEAT HEIGHT" in label:
                results["Seat Height"] = val
            elif "ARM WIDTH" in label:
                results["Arm Width"] = val
            elif "ARM DEPTH" in label:
                results["Arm Depth"] = val
            elif "ARM HEIGHT" in label:
                results["Arm Height"] = val
            elif "WIDTH" in label:
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

# ================== COM Yardage ==================
def scrape_com_yardage(driver, wait):
    try:
        fabrics_btn = driver.find_elements(By.XPATH, "//panel[@panel-title='Fabrics']//button")
        if fabrics_btn:
            driver.execute_script("arguments[0].click();", fabrics_btn[0])
            time.sleep(0.5)
        panels = driver.find_elements(By.XPATH, "//panel[@panel-title='Fabrics']//ul[li[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'COM YARDAGE')]]")
        for ul in panels:
            body_fabric = ul.find_elements(By.XPATH, ".//div[contains(text(),'Body Fabric:')]")
            for el in body_fabric:
                text = el.text.strip()
                m = re.search(r"Body Fabric:\s*(.+)", text)
                if m:
                    return m.group(1).strip()
        return ""
    except Exception:
        return ""

# ================== IMAGE SCRAPER ==================
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
            txt = sc.get_attribute("innerText") or ""
            if not txt.strip():
                continue
            data = json.loads(txt)
            candidates = data if isinstance(data, list) else [data]
            for obj in candidates:
                if not isinstance(obj, dict):
                    continue
                if obj.get("@type") == "Product":
                    img = obj.get("image")
                    if isinstance(img, str) and img.strip():
                        return img.strip()
                    if isinstance(img, list) and img:
                        return str(img[0]).strip()
        return ""
    except Exception:
        pass
    return ""

# ================== MAIN LOOP ==================
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

            # ====== IMAGE ======
            try:
                image_url = get_image_url(driver)
            except Exception:
                image_url = ""
            df.at[idx, "Image URL"] = image_url

            # ====== PRODUCT FAMILY ID ======
            try:
                name_el = driver.find_element(By.CSS_SELECTOR, "h1.product-name, div.product-title, div.one-up-title")
                product_family_id = normalize_ws(name_el.text)
            except Exception:
                product_family_id = normalize_ws(row.get("Product Name", ""))
            df.at[idx, "Product Family ID"] = product_family_id

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

            # ====== DIMENSIONS ======
            dims = scrape_dimensions_panel_all(driver, wait)
            for k, v in dims.items():
                df.at[idx, k] = v

            # ====== COM (Body Fabric Yardage) ======
            com_yard = scrape_com_yardage(driver, wait)
            df.at[idx, "COM"] = com_yard

            print(
                f"[+] {idx+1}/{len(df)} | {row.get('Product Name', '')} | "
                f"W:{df.at[idx,'Width']} D:{df.at[idx,'Depth']} H:{df.at[idx,'Height']} "
                f"SW:{df.at[idx,'Seat Width']} SH:{df.at[idx,'Seat Height']} "
                f"AW:{df.at[idx,'Arm Width']} AH:{df.at[idx,'Arm Height']} "
                f"COM:{com_yard} Img:{'yes' if image_url else 'no'}"
            )

        except Exception as e:
            print(f"[!] Error on row {idx+1}: {e}")
            for c in required_cols:
                if pd.isna(df.at[idx, c]) or df.at[idx, c] is None:
                    df.at[idx, c] = ""

        if (idx + 1) % BATCH_SIZE == 0:
            df[required_cols].to_excel(OUTPUT_FILE, index=False)
            print(f"[✓] Batch {idx+1//BATCH_SIZE} saved.")

finally:
    try:
        driver.quit()
    except Exception:
        pass
    df[required_cols].to_excel(OUTPUT_FILE, index=False)
    print(f"[✓] Done! Final updated file saved at {OUTPUT_FILE}")
