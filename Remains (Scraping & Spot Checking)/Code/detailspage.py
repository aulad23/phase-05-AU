import os, time, random
from typing import Dict, List, Set
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# =============================
# CONFIG
# =============================
INPUT_XLSX   = "remains_sconces.xlsx"
OUTPUT_XLSX  = "remains_sconces_detils.xlsx"
CHROMEDRIVER_PATH = r"C:/chromedriver.exe"

HEADLESS = False
BASE_DELAY = 0.6
JITTER = 0.6
SECTION_DELAY = 0.25
CHECKPOINT_EVERY = 5
WAIT_TIMEOUT = 25
PAGE_LOAD_TIMEOUT = 45
SCROLL_PAUSE = 0.2
MAX_RETRIES = 3
BACKOFF_BASE = 0.8

IGNORE_SECTIONS = {"contact us", "contact"}

# =============================
# HELPERS
# =============================
def robust_find(driver, by, selector):
    return WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located((by, selector))
    )

def sanitize_label(txt: str) -> str:
    return " ".join((txt or "").split()).strip()

def get_inner_text_js(driver, el) -> str:
    try:
        return (driver.execute_script("return arguments[0].innerText;", el) or "").strip()
    except Exception:
        return (el.text or "").strip()

def is_content_open(driver, content_el) -> bool:
    try:
        cls = (content_el.get_attribute("class") or "").lower()
        if "is-open" in cls:
            return True
        return bool(driver.execute_script("return arguments[0].offsetHeight > 0;", content_el))
    except Exception:
        return False

def ensure_open(driver, button, content_el):
    if not is_content_open(driver, content_el):
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", button)
        time.sleep(SECTION_DELAY + random.uniform(0, JITTER))
        driver.execute_script("arguments[0].click();", button)
        try:
            WebDriverWait(driver, WAIT_TIMEOUT).until(lambda d: is_content_open(d, content_el))
        except Exception:
            pass
    time.sleep(SECTION_DELAY / 2 + random.uniform(0, JITTER / 2))

def find_content_for_button(driver, button):
    aria_id = (button.get_attribute("aria-controls") or "").strip()
    if aria_id:
        try:
            return driver.find_element(By.ID, aria_id)
        except Exception:
            pass
    try:
        return button.find_element(By.XPATH, "following-sibling::div[contains(@class,'collapsible-content')][1]")
    except Exception:
        pass
    try:
        wrapper = button.find_element(By.XPATH, "ancestor::div[contains(@class,'collapsibles-wrapper')]")
        return wrapper.find_element(By.CSS_SELECTOR, "div.collapsible-content")
    except Exception:
        return None

# =============================
# extract_details_from_product_page
# =============================
def extract_details_from_product_page(driver) -> Dict[str, str]:
    data = {"Price": "", "SKU": "", "Descriptions": "", "Finish": ""}

    # Price
    try:
        meta = robust_find(driver, By.CSS_SELECTOR, "div.product-single__meta")
        price_el = meta.find_element(By.CSS_SELECTOR, "span.product__price")
        data["Price"] = get_inner_text_js(driver, price_el)
    except Exception:
        pass

    # SKU
    try:
        meta = robust_find(driver, By.CSS_SELECTOR, "div.product-single__meta")
        inners = meta.find_elements(By.CSS_SELECTOR, "div > div.inner") or meta.find_elements(By.CSS_SELECTOR, ".inner")
        if inners:
            data["SKU"] = get_inner_text_js(driver, inners[0])
    except Exception:
        pass

    # Descriptions
    try:
        desc_el = robust_find(driver, By.CSS_SELECTOR, "div.product-single__description")
        data["Descriptions"] = get_inner_text_js(driver, desc_el)
    except Exception:
        pass

    # Finish
    try:
        finish_wrappers = driver.find_elements(By.CSS_SELECTOR, "div.variant-wrapper.variant-wrapper--button.js")
        for wrap in finish_wrappers:
            lbl = wrap.find_element(By.CSS_SELECTOR, "label.variant__label")
            if "finish" in lbl.text.strip().lower():
                span = lbl.find_element(By.CSS_SELECTOR, "span.variant__label-info span[id^='VariantColorLabel']")
                data["Finish"] = span.text.strip()
                break
    except Exception:
        pass

    # Collapsible sections (existing logic)
    try:
        wrappers = driver.find_elements(By.CSS_SELECTOR, "div.collapsibles-wrapper")
        for wrap in wrappers:
            buttons = wrap.find_elements(By.CSS_SELECTOR, "button[aria-controls]")
            for btn in buttons:
                label = sanitize_label(btn.get_attribute("textContent"))
                if not label or label.lower() in IGNORE_SECTIONS:
                    continue

                content_el = find_content_for_button(driver, btn)
                if not content_el:
                    continue

                ensure_open(driver, btn, content_el)
                try:
                    inner = content_el.find_element(By.CSS_SELECTOR, ".collapsible-content__inner")
                    section_text = get_inner_text_js(driver, inner)
                except Exception:
                    section_text = get_inner_text_js(driver, content_el)

                if label in data and data[label]:
                    data[label] = f"{data[label]}\n\n{section_text}"
                else:
                    data[label] = section_text
    except Exception:
        pass

    return data

# =============================
# OUTPUT HELPERS
# =============================
def write_checkpoint(rows: List[Dict[str, str]], dynamic_headers: Set[str], base_cols: List[str], path: str):
    extra_cols = sorted([h for h in dynamic_headers if h not in base_cols and h.lower() not in IGNORE_SECTIONS])
    normalized = [{col: r.get(col, "") for col in (base_cols + extra_cols)} for r in rows]
    df_out = pd.DataFrame(normalized, columns=base_cols + extra_cols)
    df_out.to_excel(path, index=False)
    print(f"[Checkpoint] Wrote {len(df_out)} rows → {path}")

def safe_get(driver, url: str):
    try:
        driver.get(url)
    except TimeoutException:
        try:
            driver.execute_script("window.stop();")
        except Exception:
            pass

def fetch_with_retries(driver, url: str) -> bool:
    for attempt in range(1, MAX_RETRIES + 1):
        safe_get(driver, url)
        try:
            robust_find(driver, By.CSS_SELECTOR, "div.product-single__meta")
            return True
        except Exception as e:
            wait = BACKOFF_BASE * attempt + random.uniform(0, 0.8)
            print(f"  - attempt {attempt}/{MAX_RETRIES} failed ({e}); waiting {wait:.1f}s")
            time.sleep(wait)
    return False

# =============================
# MAIN
# =============================
def main():
    df_in = pd.read_excel(INPUT_XLSX)
    required = {"Product URL", "Image URL", "Product Name"}
    if not required.issubset(df_in.columns):
        raise ValueError(f"Input must contain columns: {required}")

    links = df_in["Product URL"].dropna().tolist()  # 🔥 process ALL rows

    print(f"Total products: {len(links)}")
    print(f"Output file: {OUTPUT_XLSX}")

    options = webdriver.ChromeOptions()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.page_load_strategy = "eager"
    options.add_argument("--disable-features=PaintHolding")
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.fonts": 2,
        "profile.managed_default_content_settings.media_stream": 2,
        "profile.managed_default_content_settings.plugins": 2
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)

    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setBlockedURLs", {
            "urls": ["*.mp4","*.webm","*.gif","*.avi","*.mov","*.m4v","*.svg","*.woff","*.woff2","*.ttf"]
        })
    except Exception:
        pass

    out_rows: List[Dict[str, str]] = []
    dynamic_headers: Set[str] = set()
    base_cols = ["Product URL", "Image URL", "Product Name", "SKU", "Price", "Descriptions", "Finish"]

    try:
        for idx, link in enumerate(links, start=1):
            try:
                ok = fetch_with_retries(driver, link)
                if not ok:
                    print(f"[{idx}/{len(links)}] ERROR: could not stabilize DOM")
                    continue

                time.sleep(SCROLL_PAUSE)

                details = extract_details_from_product_page(driver)
                for k in details.keys():
                    if k not in base_cols and k.lower() not in IGNORE_SECTIONS:
                        dynamic_headers.add(k)

                base_row = df_in.iloc[idx - 1].to_dict()
                combined = {
                    "Product URL": base_row.get("Product URL", link),
                    "Product Name": base_row.get("Product Name", ""),
                    "Image URL": base_row.get("Image URL", "")
                }
                combined.update(details)
                out_rows.append(combined)

                print(f"[{idx}/{len(links)}] OK")

                if idx % CHECKPOINT_EVERY == 0:
                    write_checkpoint(out_rows, dynamic_headers, base_cols, OUTPUT_XLSX)

                time.sleep(BASE_DELAY + random.uniform(0, JITTER))

            except Exception as e:
                print(f"[{idx}/{len(links)}] ERROR: {e}")
                if idx % CHECKPOINT_EVERY == 0:
                    write_checkpoint(out_rows, dynamic_headers, base_cols, OUTPUT_XLSX)
    finally:
        driver.quit()

    write_checkpoint(out_rows, dynamic_headers, base_cols, OUTPUT_XLSX)
    print(f"✅ Done. Saved {len(out_rows)} rows to {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
