import os
import time
import re
import pandas as pd
from urllib.parse import urldefrag

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait


# =========================
# CONFIG
# =========================
INPUT_FILE  = "janusetcie_Objects.xlsx"          # Step-1
OUTPUT_FILE = "janusetcie_Objects_final.xlsx"   # Final output (append style)
CHECKPOINT_FILE = "janusetcie_checkpoint.txt"    # last processed base URL (optional)
BATCH_SIZE = 10

# =========================
# Attach to existing Chrome
# =========================
options = Options()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 45)


# =========================
# Helpers
# =========================
def clean(txt):
    if not txt:
        return ""
    txt = txt.replace("\r", "\n")
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r"\n+", "\n", txt)
    return txt.strip()

def page_ready():
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(0.8)

def get_text(css):
    try:
        return clean(driver.find_element(By.CSS_SELECTOR, css).get_attribute("textContent"))
    except:
        return ""

def get_sku():
    metas = driver.find_elements(By.CSS_SELECTOR, ".entry-info .item-detail-meta")
    for m in metas:
        t = clean(m.get_attribute("textContent") or "")
        if "Product Number" in t:
            t = t.replace("Product Number", "").strip()
            m1 = re.search(r"\b\d+(?:-\d+){2,}\b", t)
            if m1:
                return m1.group(0)
    return ""

def get_description():
    metas = driver.find_elements(By.CSS_SELECTOR, ".entry-info .item-detail-meta")
    for m in metas:
        t = clean(m.text)
        if "Finish Shown" in t:
            continue
        if m.find_elements(By.CSS_SELECTOR, "table.item-dimensions"):
            continue
        ps = m.find_elements(By.CSS_SELECTOR, "p")
        if ps:
            return clean(ps[0].text)
    return ""

def get_finish():
    try:
        return clean(driver.find_element(By.CSS_SELECTOR, "p.top-color-label").text)
    except:
        try:
            return clean(driver.find_element(By.CSS_SELECTOR, ".color-label p").text)
        except:
            return ""

def get_main_image():
    try:
        img = driver.find_element(By.CSS_SELECTOR, ".carousel-image-wrapper img.carousel-image")
        src = img.get_attribute("src")
        if src:
            return src
    except:
        pass

    try:
        img = driver.find_element(By.CSS_SELECTOR, "img.carousel-image")
        src = img.get_attribute("src")
        if src:
            return src
    except:
        pass

    imgs = driver.find_elements(By.CSS_SELECTOR, "img")
    for im in imgs:
        src = im.get_attribute("src") or ""
        if "wp-content/uploads" in src:
            return src
    return ""

def parse_dimensions():
    data = {
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Length": "",
        "Height": "",
        "Seat Height": "",
        "Arm Height": ""
    }

    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.item-dimensions")
        rows = table.find_elements(By.CSS_SELECTOR, "tr")

        for r in rows:
            tds = r.find_elements(By.CSS_SELECTOR, "td")
            if len(tds) < 2:
                continue

            key = clean(tds[0].text).upper()
            raw_val = clean(tds[1].text)

            m = re.search(r"\d+(\.\d+)?", raw_val)
            val = m.group(0) if m else ""

            if key == "WEIGHT":
                data["Weight"] = val
            elif key == "W":
                data["Width"] = val
            elif key == "DIAM":          # priority
                data["Diameter"] = val
            elif key == "D":
                data["Depth"] = val
            elif key == "L":
                data["Length"] = val
            elif key == "H":
                data["Height"] = val
            elif key == "SH":
                data["Seat Height"] = val
            elif key == "AH":
                data["Arm Height"] = val

    except:
        pass

    return data

def get_finish_links():
    links = []
    for a in driver.find_elements(By.CSS_SELECTOR, "a.filter.product-color"):
        href = a.get_attribute("href")
        if href:
            href, _ = urldefrag(href)
            links.append(href)
    # unique ordered
    uniq = []
    seen = set()
    for u in links:
        if u not in seen:
            seen.add(u)
            uniq.append(u)
    return uniq

def save_batch(rows, mode="append"):
    """
    mode:
      - "append": read old file + append + save
      - "write":  save fresh
    """
    if not rows:
        return

    batch_df = pd.DataFrame(rows)

    final_cols = [
        "Product URL","Image URL","Product Name","SKU","Product Family Id",
        "Description","Weight","Width","Depth","Diameter","Length","Height",
        "Seat Height","Arm Height","Finish"
    ]
    for c in final_cols:
        if c not in batch_df.columns:
            batch_df[c] = ""
    batch_df = batch_df[final_cols]

    if mode == "append" and os.path.exists(OUTPUT_FILE):
        old = pd.read_excel(OUTPUT_FILE)
        out = pd.concat([old, batch_df], ignore_index=True)
        # keep first occurrence per Product URL
        out = out.drop_duplicates(subset=["Product URL"], keep="first")
        out.to_excel(OUTPUT_FILE, index=False)
    else:
        batch_df.to_excel(OUTPUT_FILE, index=False)

    print(f"✅ Batch saved: {len(rows)} rows -> {OUTPUT_FILE}")

def load_done_product_urls():
    if os.path.exists(OUTPUT_FILE):
        try:
            old = pd.read_excel(OUTPUT_FILE)
            if "Product URL" in old.columns:
                return set(str(x).strip() for x in old["Product URL"].dropna().tolist())
        except:
            pass
    return set()

def write_checkpoint(base_url):
    try:
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            f.write(base_url or "")
    except:
        pass


# =========================
# RUN (Resume)
# =========================
df_in = pd.read_excel(INPUT_FILE)
df_in.columns = [c.strip() for c in df_in.columns]

done_urls = load_done_product_urls()   # already scraped variation Product URLs
print("Already scraped Product URLs:", len(done_urls))

buffer_rows = []
processed_since_save = 0

try:
    for i, row in df_in.iterrows():
        base_url = str(row.get("Product URL", "")).strip()
        if not base_url:
            continue

        # checkpoint (optional)
        write_checkpoint(base_url)

        driver.get(base_url)
        page_ready()

        finish_links = get_finish_links()
        if not finish_links:
            finish_links = [base_url]

        for url in finish_links:
            url = (url or "").strip()
            if not url:
                continue

            # ✅ RESUME: skip if already in output
            if url in done_urls:
                continue

            driver.get(url)
            page_ready()

            pname = get_text(".entry-title .notranslate")
            dims = parse_dimensions()

            buffer_rows.append({
                "Product URL": url,
                "Image URL": get_main_image(),
                "Product Name": pname,
                "SKU": get_sku(),
                "Product Family Id": pname,
                "Description": get_description(),
                "Weight": dims["Weight"],
                "Width": dims["Width"],
                "Depth": dims["Depth"],
                "Diameter": dims["Diameter"],
                "Length": dims["Length"],
                "Height": dims["Height"],
                "Seat Height": dims["Seat Height"],
                "Arm Height": dims["Arm Height"],
                "Finish": get_finish()
            })

            done_urls.add(url)  # ✅ so even before save, duplicates skip in same run
            processed_since_save += 1

            # ✅ batch save every 10
            if processed_since_save >= BATCH_SIZE:
                save_batch(buffer_rows, mode="append")
                buffer_rows = []
                processed_since_save = 0

        print(f"[{i+1}/{len(df_in)}] Done base → {base_url} | variations: {len(finish_links)}")

    # save remaining
    if buffer_rows:
        save_batch(buffer_rows, mode="append")

    print("✅ Completed.")

finally:
    # attach mode এ driver.quit() দেবেন না
    pass
