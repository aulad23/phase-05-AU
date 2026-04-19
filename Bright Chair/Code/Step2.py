import re
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

INPUT_XLSX = "brightchair_Desk_Chairs.xlsx"
OUTPUT_XLSX = "brightchair_Desk_Chairs_step2.xlsx"

# ✅ Final column order
FINAL_COLS = [
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
    "Com",
    "Col",
    "Arm Height",
    "Seat Height",
    "Finish",
]


def make_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    # opts.add_argument("--headless=new")
    return webdriver.Chrome(options=opts)


def wait_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def clean_text(s: str) -> str:
    return " ".join((s or "").replace("\xa0", " ").split()).strip()


def key_norm(k: str) -> str:
    k = clean_text(k).lower()
    k = k.replace(":", "")
    return k.strip()


def extract_number(text: str) -> str:
    t = (text or "").strip()
    m = re.match(r"([0-9]+(?:\.[0-9]+)?)", t)
    return m.group(1) if m else ""


def contains_weight_units(text: str) -> bool:
    t = (text or "").lower()
    return bool(re.search(r"\b(lbs|lb|ibs|ib|kg|kgs)\b", t))


def parse_dimensions(raw: str):
    t = (raw or "").replace("\xa0", " ")
    t = re.sub(r"\s+", " ", t).strip()

    def grab(labels):
        for lab in labels:
            # ✅ FIX: colon optional করা এবং আরো flexible pattern
            m = re.search(rf"{lab}\s*:?\s*([0-9]+(?:\.[0-9]+)?)", t, re.I)
            if m:
                return m.group(1)
        return ""

    width = grab(["W", "WIDTH"])
    depth = grab(["D", "DEPTH"])
    height = grab(["H", "HEIGHT"])
    diameter = grab(["DIA", "DIAM", "DIAMETER"])

    # ✅ FIX: Arm Height আর Seat Height extract করা
    arm_height = grab(["ARM HEIGHT", "Arm Height"])
    seat_height = grab(["SEAT HEIGHT", "Seat Height"])

    return width, depth, diameter, height, arm_height, seat_height


def safe_click(driver, css, timeout=20):
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.5)  # scroll settle হওয়ার জন্য
    driver.execute_script("arguments[0].click();", el)
    return el


# ─────────────────────────────────────────────
# Finishes scrape — এখন আগে করা হবে
# ─────────────────────────────────────────────
def get_finishes(driver):
    try:
        safe_click(driver, '.text-nav li[info-view="product-finishes-shell"]')
        time.sleep(1.5)  # বাড়ানো: 0.8 → 1.5
        print("  ✅ Clicked Finishes tab")
    except Exception as e:
        print(f"  ⚠️  Could not click Finishes tab: {e}")
        return ""

    try:
        fin_root = WebDriverWait(driver, 25).until(  # বাড়ানো: 20 → 25
            EC.presence_of_element_located((By.CSS_SELECTOR, "#detailed-specs-shell .product-finishes-shell"))
        )
    except Exception as e:
        print(f"  ⚠️  Finishes container not found: {e}")
        return ""

    items = fin_root.find_elements(By.CSS_SELECTOR, ".finish")

    finishes = []
    seen = set()
    for it in items:
        txt = clean_text(it.text)
        if txt and txt not in seen:
            seen.add(txt)
            finishes.append(txt)

    return ", ".join(finishes)


# ─────────────────────────────────────────────
# Specifications scrape — এখন দ্বিতীয়তে
# ─────────────────────────────────────────────
def extract_specs_map(driver):
    try:
        safe_click(driver, '.text-nav li[info-view="product-details"]')
        time.sleep(1.5)  # বাড়ানো: 1.0 → 1.5
        print("  ✅ Clicked Specifications tab")
    except Exception as e:
        print(f"  ⚠️  Could not click Specifications tab: {e}")

    try:
        root = WebDriverWait(driver, 30).until(  # বাড়ানো: 25 → 30
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#detailed-specs-shell .product-details.desktop-only")
            )
        )
    except Exception as e:
        print(f"  ⚠️  Specs container not found: {e}")
        return {}

    specs = {}
    cols = root.find_elements(By.CSS_SELECTOR, ".col")

    for col in cols:
        titles = col.find_elements(By.CSS_SELECTOR, ".title p")
        details_box = col.find_elements(By.CSS_SELECTOR, ".details")

        for i, t in enumerate(titles):
            k = key_norm(t.text)
            if not k:
                continue

            # ✅ FIX: পুরো .details-এর text নেওয়া, শুধু <p> tag না
            if details_box:
                v = clean_text(details_box[0].text)
            else:
                v = ""

            specs[k] = v

        # row blocks (Not Shown / Designed by)
        rows = col.find_elements(By.CSS_SELECTOR, ".row")
        for r in rows:
            try:
                k = key_norm(r.find_element(By.CSS_SELECTOR, ".title").text)
                v = clean_text(r.find_element(By.CSS_SELECTOR, ".details").text)
                if k:
                    specs[k] = v
            except Exception:
                pass

    return specs


def get_description(driver):
    """Product description extract করা"""
    try:
        desc_el = driver.find_element(
            By.CSS_SELECTOR,
            ".product-description, .description, .detailed-description"
        )
        return clean_text(desc_el.text)
    except Exception:
        pass

    # Fallback: meta description থেকে
    try:
        meta = driver.find_element(By.CSS_SELECTOR, 'meta[name="description"]')
        return clean_text(meta.get_attribute("content"))
    except Exception:
        return ""


def extract_family_id(product_name: str) -> str:
    name = (product_name or "").strip()
    if "-" in name:
        return name.split("-")[0]
    return name


def build_columns(specs: dict):
    sku = specs.get("model no", "")

    com = extract_number(specs.get("com", ""))
    col = extract_number(specs.get("col", ""))

    dimensions_raw = specs.get("dimensions", "")
    # ✅ FIX: parse_dimensions এখন 6টা value return করে
    width, depth, diameter, height, arm_height, seat_height = parse_dimensions(dimensions_raw)

    weight_val = ""
    if "weight" in specs:
        candidate = specs.get("weight", "")
        if contains_weight_units(candidate):
            weight_val = extract_number(candidate)
        else:
            weight_val = candidate

    return {
        "SKU": sku,
        "Com": com,
        "Col": col,
        "Width": width,
        "Depth": depth,
        "Diameter": diameter,
        "Height": height,
        "Arm Height": arm_height,
        "Seat Height": seat_height,
        "Weight": weight_val,
    }


# ─────────────────────────────────────────────
# মূল scrape logic — Finishes আগে, তারপর Specs
# ─────────────────────────────────────────────
def scrape_one(driver, url):
    driver.get(url)
    wait_ready(driver, 35)  # বাড়ানো: 30 → 35

    WebDriverWait(driver, 30).until(  # বাড়ানো: 25 → 30
        EC.presence_of_element_located((By.CSS_SELECTOR, ".detailed-pre-footer, #detailed-specs-shell"))
    )
    time.sleep(1.2)  # বাড়ানো: 0.8 → 1.2 — page settle হওয়ার সময় দেওয়া

    # Description আগে নেওয়া (tab click ছাড়াই পাওয়া যায়)
    description = get_description(driver)

    # ✅ Finishes আগে — কারণ কিছু product এ এটা আগে ready হয়
    finish = get_finishes(driver)

    # ✅ তারপর Specifications
    specs_map = extract_specs_map(driver)
    cols = build_columns(specs_map)

    cols["Description"] = description
    cols["Finish"] = finish

    return cols


def main():
    df = pd.read_excel(INPUT_XLSX)

    driver = make_driver()
    try:
        new_rows = []
        for i, row in df.iterrows():
            url = str(row["Product URL"]).strip()
            name = str(row["Product Name"]).strip()
            print(f"[{i + 1}/{len(df)}] {name}")

            base = dict(row)
            base["Product Family Id"] = extract_family_id(name)

            try:
                extra = scrape_one(driver, url)
            except Exception as e:
                print(f"  ❌ Failed: {e}")
                print("  🔄 Retry করছি...")
                time.sleep(3.0)  # বাড়ানো: 2.0 → 3.0
                try:
                    extra = scrape_one(driver, url)
                    print("  ✅ Retry successful")
                except Exception as e2:
                    print(f"  ❌ Retry-ও fail: {e2}")
                    extra = {
                        "SKU": "", "Com": "", "Col": "",
                        "Width": "", "Depth": "", "Diameter": "", "Height": "",
                        "Arm Height": "", "Seat Height": "", "Weight": "",
                        "Description": "", "Finish": "",
                    }

            base.update(extra)
            new_rows.append(base)
            time.sleep(0.8)  # বাড়ানো: 0.6 → 0.8

        out_df = pd.DataFrame(new_rows)

        rest = [c for c in out_df.columns if c not in FINAL_COLS]
        out_df = out_df[FINAL_COLS + rest]

        out_df.to_excel(OUTPUT_XLSX, index=False)
        print(f"\n✅ DONE — saved: {OUTPUT_XLSX} | Rows: {len(out_df)}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()