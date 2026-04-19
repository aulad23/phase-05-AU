import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# ---------- SETUP ----------
chrome_driver_path = "C:/chromedriver.exe"

options = Options()
options.add_argument("--start-maximized")

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

# ---------- LOAD STEP 1 DATA ----------
df = pd.read_excel("Hocker_Ottomans.xlsx")

# Ensure 'Product Name' column exists
if "Product Name" not in df.columns:
    df["Product Name"] = ""

# Remove duplicates by Product URL
df = df.drop_duplicates(subset=["Product URL"]).reset_index(drop=True)

# ---------- ADD COLUMNS ----------
df["Description"] = ""
df["Weight"] = ""
df["Width"] = ""
df["Depth"] = ""
df["Diameter"] = ""
df["Height"] = ""
df["Wood Species"] = ""
df["Volume"] = ""
df["Product Family ID"] = df["Product Name"]
df["COM"] = ""
df["COL"] = ""
df["SeatType"] = ""
df["Seat Fill"] = ""
df["BackType"] = ""
df["Back Fill"] = ""

# New fields
df["Finish"] = ""
df["Fabric"] = ""
df["Cushion"] = ""
df["Seat Cushion Fill"] = ""
df["Back Cushion Fill"] = ""

# NEW: extra two columns
df["Seat Height"] = ""
df["Arm Height"]  = ""

# ---------- SCRAPE DETAILS ----------
for idx, row in df.iterrows():
    url = row["Product URL"]
    driver.get(url)
    time.sleep(2)

    print(f"\nScraping {idx+1}/{len(df)}: {url}")

    # --- Description (no change) ---
    try:
        desc = driver.find_element(By.ID, "product-description").text.strip()
    except:
        desc = ""
    df.at[idx, "Description"] = desc

    # --- Product Name (no change) ---
    try:
        name_el = driver.find_element(By.CSS_SELECTOR, "h1.product-name")
        name_text = name_el.text.strip()
        if row["Product Name"] == "":
            df.at[idx, "Product Name"] = name_text
    except:
        pass

    # ===============================
    # 1) CLICK "AS SHOWN" → get Finish/Fabric/Cushion/Seat Cushion Fill/Back Cushion Fill
    # ===============================
    try:
        # try both locators
        try:
            as_shown_tab = driver.find_element(By.CSS_SELECTOR, 'li#asShown a.ui-tabs-anchor')
        except:
            as_shown_tab = driver.find_element(By.CSS_SELECTOR, 'a#ui-id-1')
        driver.execute_script("arguments[0].click();", as_shown_tab)
        time.sleep(0.7)
    except:
        pass

    # collect values from As Shown panel
    finish_values = []
    try:
        as_panel = driver.find_element(By.ID, "asShownAs")

        # Case A: simple lines like "Fabric: XXX" / "Finish: YYY"
        for div in as_panel.find_elements(By.XPATH, ".//div"):
            t = (div.text or "").strip()
            if not t:
                continue
            if t.lower().startswith("finish"):
                v = t.split(":", 1)[-1].strip() if ":" in t else ""
                if v:
                    finish_values.append(v)
            elif t.lower().startswith("fabric"):
                if not df.at[idx, "Fabric"]:
                    df.at[idx, "Fabric"] = t.split(":", 1)[-1].strip() if ":" in t else ""
            elif t.lower().startswith("cushion"):
                if not df.at[idx, "Cushion"]:
                    df.at[idx, "Cushion"] = t.split(":", 1)[-1].strip() if ":" in t else ""
            elif t.lower().startswith("seat cushion fill"):
                if not df.at[idx, "Seat Cushion Fill"]:
                    df.at[idx, "Seat Cushion Fill"] = t.split(":", 1)[-1].strip() if ":" in t else ""
            elif t.lower().startswith("back cushion fill"):
                if not df.at[idx, "Back Cushion Fill"]:
                    df.at[idx, "Back Cushion Fill"] = t.split(":", 1)[-1].strip() if ":" in t else ""

        # Case B: span.asaDesc + span.asaVal like "Finish 1:" + "Ecru"
        rows = as_panel.find_elements(By.XPATH, './/*[contains(@class,"configWrapper")]//div | .//div')
        for r in rows:
            try:
                label_el = r.find_element(By.XPATH, './/span[contains(@class,"asaDesc")]')
                val_el   = r.find_element(By.XPATH, './/span[contains(@class,"asaVal")]')
                label = (label_el.text or "").strip().lower()
                value = (val_el.text or "").strip()
                if not value:
                    continue

                if "finish" in label:
                    finish_values.append(value)
                elif "fabric" in label and not df.at[idx, "Fabric"]:
                    df.at[idx, "Fabric"] = value
                elif "seat cushion fill" in label and not df.at[idx, "Seat Cushion Fill"]:
                    df.at[idx, "Seat Cushion Fill"] = value
                elif "back cushion fill" in label and not df.at[idx, "Back Cushion Fill"]:
                    df.at[idx, "Back Cushion Fill"] = value
                elif "cushion" in label and not df.at[idx, "Cushion"]:
                    df.at[idx, "Cushion"] = value
            except:
                continue

        # Case C: raw text regex sweep (handles "Finish 1: Ecru")
        raw_txt = as_panel.text
        for m in re.finditer(r'Finish\s*\d*\s*:\s*([^\n\r|]+)', raw_txt, re.IGNORECASE):
            v = (m.group(1) or "").strip()
            if v:
                finish_values.append(v)
        if not df.at[idx, "Fabric"]:
            m = re.search(r'Fabric\s*:\s*([^\n\r|]+)', raw_txt, re.IGNORECASE)
            if m: df.at[idx, "Fabric"] = m.group(1).strip()
        if not df.at[idx, "Seat Cushion Fill"]:
            m = re.search(r'Seat\s+Cushion\s+Fill\s*:\s*([^\n\r|]+)', raw_txt, re.IGNORECASE)
            if m: df.at[idx, "Seat Cushion Fill"] = m.group(1).strip()
        if not df.at[idx, "Back Cushion Fill"]:
            m = re.search(r'Back\s+Cushion\s+Fill\s*:\s*([^\n\r|]+)', raw_txt, re.IGNORECASE)
            if m: df.at[idx, "Back Cushion Fill"] = m.group(1).strip()
        if not df.at[idx, "Cushion"]:
            m = re.search(r'(?<!Seat\s)(?<!Back\s)Cushion\s*:\s*([^\n\r|]+)', raw_txt, re.IGNORECASE)
            if m: df.at[idx, "Cushion"] = m.group(1).strip()

        # dedup & join finishes
        if finish_values:
            seen = set()
            final_fin = []
            for v in finish_values:
                key = v.lower().strip()
                if key and key not in seen:
                    final_fin.append(v.strip())
                    seen.add(key)
            df.at[idx, "Finish"] = " | ".join(final_fin)
    except:
        pass

    # ===============================
    # 2) CLICK "DIMENSIONS" → then get Seat Height & Arm Height
    # ===============================
    try:
        try:
            dim_tab = driver.find_element(By.CSS_SELECTOR, 'li#dimtab a.ui-tabs-anchor')
        except:
            dim_tab = driver.find_element(By.CSS_SELECTOR, 'a#ui-id-2')
        driver.execute_script("arguments[0].click();", dim_tab)
        time.sleep(0.7)
    except:
        pass

    # Parse #dims seat/arm rows
    try:
        dims_panel = driver.find_element(By.ID, "dims")
        trs = dims_panel.find_elements(By.XPATH, ".//table//tr[td]")

        seat_h = ""
        arm_h  = ""

        for tr in trs:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            label = (tds[0].text or "").strip().lower()

            if label.startswith("seat"):
                vals = [td.text.strip() for td in tds[1:]]
                val = next((v for v in reversed(vals) if v), "")
                if val:
                    m = re.search(r'([\d\.]+)\s*in', val, re.I)
                    seat_h = m.group(1) if m else (re.search(r'[\d\.]+', val).group(0) if re.search(r'[\d\.]+', val) else "")

            if label == "arm" or "arm height" in label:
                vals = [td.text.strip() for td in tds[1:]]
                val = next((v for v in reversed(vals) if v), "")
                if val:
                    m = re.search(r'([\d\.]+)\s*in', val, re.I)
                    arm_h = m.group(1) if m else (re.search(r'[\d\.]+', val).group(0) if re.search(r'[\d\.]+', val) else "")

        if seat_h:
            df.at[idx, "Seat Height"] = seat_h
        if arm_h:
            df.at[idx, "Arm Height"] = arm_h
    except:
        pass

    # --- (your original) Dimensions text block (unchanged) ---
    try:
        dims = driver.find_element(By.ID, "product-dims").text.strip()
    except:
        dims = ""

    width, depth, height, diameter = "", "", "", ""
    if dims:
        w_match = re.search(r'W\s*([\d\.]+)', dims, re.I)
        d_match = re.search(r'D\s*([\d\.]+)', dims, re.I)
        h_match = re.search(r'H\s*([\d\.]+)', dims, re.I)
        dia_match = re.search(r'([\d\.]+)\s*(?:dia|Dia|DIAMETER)', dims, re.I)
        high_match = re.search(r'([\d\.]+)\s*(?:h|H|high|High)', dims, re.I)

        if w_match: width = w_match.group(1)
        if d_match: depth = d_match.group(1)
        if h_match: height = h_match.group(1)
        if dia_match: diameter = dia_match.group(1)
        if high_match: height = high_match.group(1)

    df.at[idx, "Width"] = width
    df.at[idx, "Depth"] = depth
    df.at[idx, "Height"] = height
    df.at[idx, "Diameter"] = diameter

    # --- Accordion Details (unchanged) ---
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "div.accordion-section table#prodDetailTable tr")
        for r in rows:
            try:
                cols = r.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 2:
                    label = cols[0].text.strip()
                    value = cols[1].text.strip()

                    if "Weight" in label:
                        weight_num = re.search(r'[\d\.]+', value)
                        df.at[idx, "Weight"] = weight_num.group(0) if weight_num else ""
                    elif "Volume" in label:
                        df.at[idx, "Volume"] = value
                    elif "Wood Species" in label:
                        df.at[idx, "Wood Species"] = value
                    elif "COM" in label:
                        df.at[idx, "COM"] = value
                    elif "COL" in label:
                        df.at[idx, "COL"] = value
                    elif "SeatType" in label:
                        df.at[idx, "SeatType"] = value
                    elif "Seat Fill" in label:
                        df.at[idx, "Seat Fill"] = value
                    elif "BackType" in label:
                        df.at[idx, "BackType"] = value
                    elif "Back Fill" in label:
                        df.at[idx, "Back Fill"] = value
            except:
                pass
    except:
        pass

    # --- PRINT TERMINAL OUTPUT ---
    print("Description:", desc[:100] + "..." if len(desc) > 100 else desc)
    print("Finish:", df.at[idx, "Finish"], "| Fabric:", df.at[idx, "Fabric"])
    print("Cushion:", df.at[idx, "Cushion"])
    print("Seat Cushion Fill:", df.at[idx, "Seat Cushion Fill"], "| Back Cushion Fill:", df.at[idx, "Back Cushion Fill"])
    print("Seat Height:", df.at[idx, "Seat Height"], "| Arm Height:", df.at[idx, "Arm Height"])
    print("Width:", df.at[idx, "Width"], "Depth:", df.at[idx, "Depth"], "Height:", df.at[idx, "Height"], "Diameter:", df.at[idx, "Diameter"])
    print("Weight:", df.at[idx, "Weight"], "Volume:", df.at[idx, "Volume"])
    print("-" * 50)

driver.quit()

# ---------- SAVE STEP 2 EXCEL ----------
df = df.drop_duplicates(subset=["Product URL"]).reset_index(drop=True)
step2_file = "Hocker_Ottomans.xlsx"
df.to_excel(step2_file, index=False)
#print(f"\n✅ Step 2 complete! {len(df)} unique products saved to {step2_file}")

# ---------- STEP 3: FINAL FILE ----------
final_columns = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family ID",
    "Description", "Weight", "Width", "Depth", "Diameter", "Height",
    "COM", "COL", "SeatType", "BackType",
    "Finish", "Fabric", "Cushion", "Seat Cushion Fill", "Back Cushion Fill",
    "Seat Height", "Arm Height"
]

df_final = df.loc[:, [c for c in final_columns if c in df.columns]]
df_final = df_final.drop_duplicates(subset=["Product URL"]).reset_index(drop=True)

final_file = "Hocker_Ottomans_final.xlsx"
df_final.to_excel(final_file, index=False)
print(f"\n✅ Step 3 complete! {len(df_final)} unique products saved to {final_file}")
