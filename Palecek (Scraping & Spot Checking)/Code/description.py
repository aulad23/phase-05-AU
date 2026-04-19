import os
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ----------------- CONFIG -----------------
CHROMEDRIVER_PATH = r"C:\chromedriver.exe"
INPUT_XLSX = "palecek_bar_stools.xlsx"
OUTPUT_XLSX = "palecek_bar_stools_Final.xlsx"
UNMATCHED_LOG = "unmatched_labels.txt"
WAIT_TIMEOUT = 25
PAGE_PAUSE = 0.8
# ------------------------------------------

# Base columns
BASE_COLUMNS = ["Width", "Depth", "Height", "Diameter", "Description", "Finish", "Cushion", "Seat"]

# New specific ProductSpecifications columns
SPEC_COLUMNS = [
    "Canopy Dimensions",
    "Shade Dimension",
    "Shade Detail",
    "Canopy Finish",
    "Chain Length",
    "Chain Finish",
    "Adjustable Length - Min",
    "Adjustable Length - Max",
    "Socket Qty",
    "Socket Type",
    "Bulb Type",
    "Bulb Wattage",
    "Voltage",
    "Cord Length",
    "Cord Color"
]

ALL_COLUMNS = BASE_COLUMNS + SPEC_COLUMNS


# ----------------- FUNCTIONS -----------------
def connect_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-infobars")
    opts.add_argument("--disable-extensions")
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=opts)
    return driver


def clean_quotes(s: str) -> str:
    if not s:
        return s
    return s.replace('“', '"').replace('”', '"').replace("′", "'").replace("’", "'")


def extract_dimensions(text: str):
    res = {"Depth": "", "Width": "", "Height": "", "Diameter": ""}
    if not text:
        return res
    t = clean_quotes(text)
    num = r'(\d+(?:\.\d+)?(?:\s*\d+/\d+)?)'
    patterns = {
        "Depth": rf'{num}\s*(?:["]|in)?\s*(?:L\b|D\b|DP\b)',
        "Width": rf'{num}\s*(?:["]|in)?\s*[Ww]\b',
        "Height": rf'{num}\s*(?:["]|in)?\s*[Hh]\b',
        "Diameter": rf'{num}\s*(?:["]|in)?\s*(?:Dia|Diameter)\b',
    }
    for key, pat in patterns.items():
        m = re.search(pat, t, flags=re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if not val.endswith('"'):
                val = val + '"'
            res[key] = val
    return res


def extract_description(driver):
    try:
        desc_container = driver.find_element(By.CSS_SELECTOR, "div#item-info-short-description")
    except:
        return ""
    ps = desc_container.find_elements(By.CSS_SELECTOR, "p")
    chunks = []
    for p in ps:
        try:
            label = p.find_element(By.CSS_SELECTOR, "span.additionalSpecRepeaterAttributeName").text.strip()
        except:
            label = ""
        full = p.text.strip()
        if not full:
            continue
        if label:
            lab_esc = re.escape(label)
            cleaned = re.sub(rf'^\s*{lab_esc}\s*:\s*', '', full, flags=re.IGNORECASE)
            cleaned = re.sub(rf'^\s*{lab_esc}\s*', '', cleaned, flags=re.IGNORECASE)
        else:
            cleaned = full
        if cleaned:
            chunks.append(cleaned.strip())
    return " ".join(chunks).strip()


def extract_finish_and_cushion_from_selected_options(driver):
    finish_val = ""
    cushion_val = ""
    try:
        panel = driver.find_element(By.CSS_SELECTOR, "#UsedCoversListPanel #selected-options")
        lis = panel.find_elements(By.CSS_SELECTOR, "li")
        for li in lis:
            try:
                label_raw = li.find_element(By.CSS_SELECTOR, "span:first-child").text.strip()
            except:
                label_raw = ""
            label = label_raw.rstrip(":").strip().lower()
            avail_txt = ""
            try:
                avail_txt = li.find_element(By.CSS_SELECTOR, "span.selected-item-availability").text.strip()
            except:
                pass
            full = li.text.strip()
            if label_raw:
                full = re.sub(rf'^{re.escape(label_raw)}\s*', '', full, flags=re.IGNORECASE).strip()
            if avail_txt:
                full = full.replace(avail_txt, "").strip()
            if label.startswith("finish") and not finish_val:
                finish_val = full
            elif label.startswith("cushion") and not cushion_val:
                cushion_val = full
    except:
        pass
    return finish_val, cushion_val


def extract_seat(driver):
    """Extract Seat info from ProductSpecifications tab."""
    seat_val = ""
    try:
        ps = driver.find_elements(By.CSS_SELECTOR, "div.tab-content[data-tab-class='ProductSpecifications'] p")
        for p in ps:
            try:
                span = p.find_element(By.CSS_SELECTOR, "span.additionalSpecRepeaterAttributeName")
                label = span.text.strip()
            except:
                continue
            if label.lower().startswith("seat"):
                full_text = p.text.strip()
                seat_val = re.sub(r"^Seat\s*:\s*", "", full_text, flags=re.IGNORECASE).strip()
                break
    except:
        pass
    return seat_val


def extract_spec_fields_specific(driver):
    """
    Extract only the predefined SPEC_COLUMNS from ProductSpecifications tab.
    Uses fuzzy/partial matching to handle label variations, with Shade fixes.
    """
    specs = {col: "" for col in SPEC_COLUMNS}
    unmatched_labels = []
    try:
        ps = driver.find_elements(
            By.CSS_SELECTOR,
            "div.product-tabs div.tab-content[data-tab-class='ProductSpecifications'] p"
        )
        for p in ps:
            try:
                label = p.find_element(
                    By.CSS_SELECTOR, "span.additionalSpecRepeaterAttributeName"
                ).text.strip().rstrip(":")
                value = p.text.strip()
                value = re.sub(
                    rf'^{re.escape(label)}\s*:\s*', '', value, flags=re.IGNORECASE
                ).strip()

                matched = False
                for col in SPEC_COLUMNS:
                    # ---- Specific handling for Shade fields ----
                    if "shade detail" in label.lower() and col.lower() == "shade detail":
                        specs["Shade Detail"] = value
                        matched = True
                        break
                    elif "shade dimension" in label.lower() and col.lower() == "shade dimension":
                        specs["Shade Dimension"] = value
                        matched = True
                        break
                    # ---- General fuzzy match ----
                    elif (
                        col.lower() in label.lower()
                        or label.lower() in col.lower()
                        or col.lower().split()[0] in label.lower()
                    ):
                        specs[col] = value
                        matched = True
                        break

                if not matched:
                    unmatched_labels.append(label)
            except:
                continue
    except:
        pass

    # log unmatched
    if unmatched_labels:
        with open(UNMATCHED_LOG, "a", encoding="utf-8") as f:
            for lbl in unmatched_labels:
                f.write(lbl + "\n")

    return specs


def extract_finish_fallback_from_specs(driver):
    """Fallback Finish extraction (only exact 'Finish:' or 'Finish' label)."""
    finish_val = ""
    try:
        ps = driver.find_elements(By.CSS_SELECTOR, "div.tab-content[data-tab-class='ProductSpecifications'] p")
        for p in ps:
            try:
                label = p.find_element(By.CSS_SELECTOR, "span.additionalSpecRepeaterAttributeName").text.strip()
            except:
                continue
            if label.lower() == "finish":
                txt = p.text.strip()
                finish_val = re.sub(r"^Finish\s*:\s*", "", txt, flags=re.IGNORECASE).strip()
                break
    except:
        pass
    return finish_val


def ensure_output_columns(df):
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df


# ----------------- MAIN -----------------
def main():
    if not os.path.exists(INPUT_XLSX):
        print(f"❌ ERROR: Input file not found: {INPUT_XLSX}")
        return

    df = pd.read_excel(INPUT_XLSX)
    if "Product URL" not in df.columns:
        raise RuntimeError("Input Excel must contain a 'Product URL' column.")

    df = ensure_output_columns(df)
    urls = [u for u in df["Product URL"].astype(str).tolist() if u and u.lower() != "nan"]

    driver = connect_driver()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    total = len(urls)

    # clear unmatched log
    open(UNMATCHED_LOG, "w", encoding="utf-8").close()

    for idx, url in enumerate(urls, start=1):
        try:
            driver.get(url)
            wait.until(
                EC.any_of(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div#prod-desc-dim-container p.prod-desc-dim")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div#item-info-short-description")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.product-tabs div.tab-content[data-tab-class='ProductSpecifications']"))
                )
            )
            time.sleep(PAGE_PAUSE)

            # --- Dimensions ---
            dims_text = ""
            try:
                dim_p = driver.find_element(By.CSS_SELECTOR, "div#prod-desc-dim-container p.prod-desc-dim")
                dims_text = clean_quotes(dim_p.text.strip())
                dims_text = re.sub(r'^\s*Overall\s+Dimensions\s*:\s*', '', dims_text, flags=re.IGNORECASE).strip()
            except:
                pass
            dims = extract_dimensions(dims_text)

            # --- Description ---
            description = extract_description(driver)

            # --- Finish & Cushion ---
            finish, cushion = extract_finish_and_cushion_from_selected_options(driver)
            if not finish:
                finish = extract_finish_fallback_from_specs(driver)

            # --- Seat ---
            seat_value = extract_seat(driver)

            # --- ProductSpecifications fields ---
            specs = extract_spec_fields_specific(driver)

            # --- Update DataFrame row ---
            row_idx = df.index[df["Product URL"] == url]
            if len(row_idx) > 0:
                i = row_idx[0]
                for key, val in dims.items():
                    if val:
                        df.at[i, key] = val
                if description:
                    df.at[i, "Description"] = description
                if finish:
                    df.at[i, "Finish"] = finish
                if cushion:
                    df.at[i, "Cushion"] = cushion
                if seat_value:
                    df.at[i, "Seat"] = seat_value
                for k, v in specs.items():
                    if v:
                        df.at[i, k] = v

            filled_count = sum(1 for v in specs.values() if v.strip())
            print(f"[{idx}/{total}] ✅ {url} | {filled_count}/{len(specs)} specs filled")

        except Exception as e:
            print(f"[{idx}/{total}] ⚠️ ERROR on {url}: {e}")

    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"\n✅ Saved results to: {OUTPUT_XLSX} (rows: {len(df)})")
    print(f"⚠️ Unmatched labels (if any) saved to: {UNMATCHED_LOG}")


if __name__ == "__main__":
    main()
