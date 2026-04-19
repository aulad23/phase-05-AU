# -*- coding: utf-8 -*-
"""
Loloi Rugs - Step 2 (Details Page Enrichment)
✅ Final Version (Headless ON/OFF + Resume + One Excel + Live Save)
- Reads:  loloi_rugs_list.xlsx
- Writes: loloi_rugs_details.xlsx  (auto resume)
"""

import re
import time
from pathlib import Path
from typing import Dict, Tuple
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException
)

# =================== Config ===================
HEADLESS = False          # 👈 True = hide browser | False = show browser
PAGELOAD_TIMEOUT = 45
WAIT_SEC = 20

SLEEP_EACH  = 0.5         # polite delay
SAVE_EVERY  = 5           # save every 5 processed rows

BASE_DIR     = Path(__file__).parent
INPUT_FILE   = BASE_DIR / "demo_Rugs.xlsx"
MASTER_XLSX  = BASE_DIR / "demo_Rugs_Update.xlsx"
# ==============================================


# ---------- Helper: Clean SKU ----------
def clean_sku(sku: str) -> str:
    """
    Normalize SKU casing: capitalize only the alphabetic prefix before first '-'
    keep the rest (numbers) exactly as-is.
    Examples:
      'ABI-01'  → 'Abi-01'
      'ABi-01'  → 'Abi-01'
      'ZUP-02'  → 'Zup-02'
      'LAYLA-03'→ 'Layla-03'
    """
    if not sku:
        return sku
    parts = sku.split("-", 1)          # split only on first '-'
    parts[0] = parts[0].capitalize()   # ✅ "ABI" → "Abi", "ABi" → "Abi"
    return "-".join(parts)


# ---------- Helper: Extract Product Family ID ----------
def extract_family_id(sku: str) -> str:
    """
    Extract Product Family ID = part before the first '-' from (already cleaned) SKU.
    Examples:
      'Abi-01'   → 'Abi'
      'Zup-02'   → 'Zup'
      'Layla-03' → 'Layla'
    """
    if not sku:
        return ""
    if "-" in sku:
        return sku.split("-")[0].strip()
    return sku.split()[0].strip() if sku else ""


# ---------- Helper: Clean Pile Height ----------
def clean_pile_height(value: str) -> str:
    """
    Remove inch symbol (") from Pile Height value.
    Examples:
      '0.125"'  → '0.125'
      '0.5"'    → '0.5'
    """
    return value.replace('"', '').strip()


# ---------- Chrome Setup ----------
def make_driver(headless: bool = HEADLESS):
    opts = Options()
    opts.page_load_strategy = "eager"
    if headless:
        opts.add_argument("--headless=new")
    else:
        print("🪟 Browser window will open (Headless=False)")

    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,1000")
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(PAGELOAD_TIMEOUT)
    return driver


def go_to(driver, url: str, wait_css: str = "body", attempts: int = 3) -> None:
    url = (url or "").strip()
    if not url:
        raise ValueError("Empty URL")
    last_err = None
    for _ in range(attempts):
        try:
            driver.get(url)
            WebDriverWait(driver, WAIT_SEC).until(EC.presence_of_element_located((By.CSS_SELECTOR, wait_css)))
            return
        except Exception as e:
            last_err = e
            time.sleep(1.5)
    raise last_err


def safe_click(driver, el) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.15)
        el.click()
        return True
    except (ElementClickInterceptedException, StaleElementReferenceException):
        try:
            driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            return False
    except Exception:
        return False


def close_popups(driver):
    for sel in [
        "button[aria-label='Close']",
        "button#onetrust-accept-btn-handler",
        "button.cookie-accept",
        "button.klaviyo-close-form",
        "div#klaviyo-bis-close",
    ]:
        try:
            for e in driver.find_elements(By.CSS_SELECTOR, sel):
                if e.is_displayed():
                    safe_click(driver, e); time.sleep(0.2)
        except Exception:
            pass


# ---------- Extraction ----------
def get_description(driver) -> str:
    try:
        acc = WebDriverWait(driver, WAIT_SEC).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "div.accordion.js-accordion-open-all.spacer-4--mt[data-category='Product Detail'][data-action='Info']"
        )))
        try:
            desc_item = acc.find_element(
                By.XPATH,
                ".//div[contains(@class,'accordion__item')][.//div[contains(@class,'accordion__title') and contains(translate(.,'DESCRIPTION','description'),'description')]]"
            )
            title_btn = desc_item.find_element(By.CSS_SELECTOR, "div.accordion__title")
            safe_click(driver, title_btn)
            time.sleep(0.2)
            content = desc_item.find_element(By.CSS_SELECTOR, "div.accordion__content")
            return content.text.strip()
        except NoSuchElementException:
            content = acc.find_element(By.CSS_SELECTOR, "div.accordion__content")
            return content.text.strip()
    except Exception:
        return ""


def parse_details(driver) -> Dict[str, str]:
    data = {"Construction":"", "Material":"", "Pile Height":"", "Backing":"", "Country of Origin":""}
    try:
        drawer = WebDriverWait(driver, WAIT_SEC).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "div.accordion.js-accordion-open-all.spacer-4--mt[data-category='Product Detail'][data-action='Info']"
        )))
        try:
            dl = drawer.find_element(By.CSS_SELECTOR, "div#details-drawer > div.accordion__content dl")
        except NoSuchElementException:
            dl = drawer.find_element(By.CSS_SELECTOR, "div.accordion__content dl")
    except Exception:
        try:
            content = drawer.find_element(By.CSS_SELECTOR, "div.accordion__content")
            lines = [li.text.strip() for li in content.find_elements(By.CSS_SELECTOR,"li") if li.text.strip()]
            if not lines:
                lines = content.text.strip().split("\n")
            for line in lines:
                m = re.match(r"\s*([^:]+):\s*(.+)\s*$", line)
                if not m: continue
                tkey, tval = m.group(1).strip().lower(), m.group(2).strip()
                if tkey == "construction":        data["Construction"] = tval
                elif tkey == "material":          data["Material"] = tval
                elif tkey == "pile height":       data["Pile Height"] = clean_pile_height(tval)
                elif tkey == "backing":           data["Backing"] = tval
                elif tkey == "country of origin": data["Country of Origin"] = tval
            return data
        except Exception:
            return data

    groups = dl.find_elements(By.CSS_SELECTOR, ":scope > div")

    def dd_text(dd):
        try:
            lis = dd.find_elements(By.CSS_SELECTOR, "li")
            if lis:
                return ", ".join(li.text.strip() for li in lis if li.text.strip())
        except Exception:
            pass
        return dd.text.strip()

    for g in groups:
        try:
            dt = g.find_element(By.CSS_SELECTOR, "dt").text.strip()
            dd = g.find_element(By.CSS_SELECTOR, "dd")
            val = dd_text(dd)
        except Exception:
            continue
        key = dt.lower()
        if key == "construction":
            data["Construction"] = val
        elif key == "material":
            data["Material"] = val
        elif key == "country of origin":
            data["Country of Origin"] = val
        elif key == "technical specs":
            try:
                items = [li.text.strip() for li in dd.find_elements(By.CSS_SELECTOR, "li") if li.text.strip()]
            except Exception:
                items = re.split(r"[;\n]+", dd.text.strip())
            for item in items:
                m = re.match(r"\s*([^:]+):\s*(.+)\s*$", item)
                if not m: continue
                tkey, tval = m.group(1).strip().lower(), m.group(2).strip()
                if tkey == "pile height": data["Pile Height"] = clean_pile_height(tval)
                elif tkey == "backing":   data["Backing"] = tval
    return data


def process_row(driver, url: str) -> Tuple[str, Dict[str, str]]:
    if not url:
        return "", {"Construction":"", "Material":"", "Pile Height":"", "Backing":"", "Country of Origin":""}
    try:
        go_to(driver, url, wait_css="div.accordion.js-accordion-open-all.spacer-4--mt[data-category='Product Detail'][data-action='Info']")
    except Exception:
        return "", {"Construction":"", "Material":"", "Pile Height":"", "Backing":"", "Country of Origin":""}
    close_popups(driver)
    return get_description(driver), parse_details(driver)


# ---------- Main ----------
def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Missing {INPUT_FILE}. Run Step 1 first.")

    df_all = pd.read_excel(INPUT_FILE)
    df_all = df_all.drop_duplicates(subset=["Product URL"], keep="first").reset_index(drop=True)

    # Load previous progress
    if MASTER_XLSX.exists():
        done_df = pd.read_excel(MASTER_XLSX)
        done_urls = set(done_df["Product URL"].dropna().tolist())
        print(f"🔁 Resuming from saved progress ({len(done_urls)} done)")
    else:
        done_df = pd.DataFrame()
        done_urls = set()

    df = df_all[~df_all["Product URL"].isin(done_urls)].reset_index(drop=True)
    print(f"🟢 Remaining: {len(df)} URLs")

    if df.empty:
        print("✅ All done!")
        return

    STEP1_COLS = ["Product URL", "Image URL", "Product Name", "SKU"]
    NEW_COLS   = ["Product Family ID", "Description", "Construction", "Material", "Pile Height", "Backing", "Country of Origin"]
    ALL_COLS   = STEP1_COLS + NEW_COLS

    driver = make_driver()
    new_data = []

    try:
        for i, row in df.iterrows():
            url = str(row["Product URL"]).strip()
            if not url:
                continue

            print(f"[{i+1}/{len(df)}] 🔗 {url}")
            desc, det = process_row(driver, url)

            product_name = str(row.get("Product Name", "")).strip()

            # ✅ SKU: normalize casing → "ABI-01" or "ABi-01" → "Abi-01"
            sku = clean_sku(str(row.get("SKU", "")).strip())

            # ✅ Product Family ID: taken from cleaned SKU → "Abi-01" → "Abi"
            family_id = extract_family_id(sku)

            new_data.append({
                "Product URL":       url,
                "Image URL":         row.get("Image URL", ""),
                "Product Name":      product_name,
                "SKU":               sku,
                "Product Family ID": family_id,
                "Description":       desc,
                "Construction":      det.get("Construction", ""),
                "Material":          det.get("Material", ""),
                "Pile Height":       det.get("Pile Height", ""),
                "Backing":           det.get("Backing", ""),
                "Country of Origin": det.get("Country of Origin", ""),
            })

            print(f"   ✔ Done | SKU: {sku} | Family ID: {family_id} ({len(done_urls)+len(new_data)} total)")

            # save every N rows
            if len(new_data) % SAVE_EVERY == 0:
                temp_df = pd.DataFrame(new_data, columns=ALL_COLS)
                merged = pd.concat([done_df, temp_df], ignore_index=True)
                merged.to_excel(MASTER_XLSX, index=False)
                print(f"💾 Progress saved ({len(merged)} rows)")
                time.sleep(1)

            time.sleep(SLEEP_EACH)

        # Final save
        if new_data:
            final_df = pd.concat([done_df, pd.DataFrame(new_data, columns=ALL_COLS)], ignore_index=True)
            final_df.to_excel(MASTER_XLSX, index=False)
            print(f"\n✅ Final saved: {MASTER_XLSX.name} ({len(final_df)} total rows)")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()