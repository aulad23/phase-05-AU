import time
import os
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests

# ---------- CONFIG ----------
category_url = "https://highlandhousefurniture.com/Consumer/ShowItems.aspx?TypeID=81"
output_file = r"C:\Users\ATS\Downloads\HighlandHouse_Loungchair_Final2.xlsx"
batch_size = 5
# ----------------------------

os.makedirs(os.path.dirname(output_file), exist_ok=True)

# ---------- Selenium Setup ----------
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)

# ---------- Step 1: Collect all product URLs ----------
driver.get(category_url)
time.sleep(2)

all_product_urls = set()
page_num = 1

while True:
    print(f"Scraping page {page_num} for product URLs ...")
    time.sleep(2)

    product_elements = driver.find_elements(By.CSS_SELECTOR, "li.prodListingDiv div.prodSearchDiv a[href]")
    for elem in product_elements:
        href = elem.get_attribute("href")
        if href:
            all_product_urls.add(href)

    try:
        view_all = driver.find_element(By.CSS_SELECTOR, "span.viewAll.prodPageNavItem")
        if view_all.is_displayed():
            driver.execute_script("arguments[0].scrollIntoView(true);", view_all)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", view_all)
            time.sleep(4)
            print("✔ 'View All' clicked.")
            product_elements = driver.find_elements(By.CSS_SELECTOR, "li.prodListingDiv div.prodSearchDiv a[href]")
            for elem in product_elements:
                href = elem.get_attribute("href")
                if href:
                    all_product_urls.add(href)
            break
    except:
        pass

    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, "span.nextPage.prodPageNavItem")
        style = next_btn.get_attribute("style")
        if "hidden" in style or not next_btn.is_displayed():
            print("No more pages.")
            break

        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", next_btn)
        page_num += 1
        time.sleep(3)
    except:
        print("Next not found.")
        break

print(f"\nTotal URLs collected: {len(all_product_urls)}")
driver.quit()

# ---------- Step 2: Scrape product details ----------
all_data = []

def extract_dimension_value(soup, span_id):
    tag = soup.find('span', id=span_id)
    if tag and tag.text.strip():
        return tag.text.strip().replace("in", "").strip()
    return ""

# STRICT FINISH BLOCK -----------------------------------------------------------------

BAD_WORDS = [
    "welt", "cushion", "text", "down", "banding", "border",
    "fabric", "leather", "cloth", "contrast", "trim",
    "skirt", "pillow", "seat", "back", "arm", "micro", "regular"
]

def is_valid_finish(val):
    v = val.strip()

    if not v:
        return False
    if any(bad in v.lower() for bad in BAD_WORDS):
        return False
    if v.isdigit():
        return False
    if re.search(r"\d", v):
        return False
    if "/" in v:
        return False
    if len(v) > 25:
        return False
    if len(v.split()) > 3:
        return False
    if v.upper().startswith("STR"):
        return False
    if "select" in v.lower():
        return False

    return True

# ----------------- START SCRAPING -----------------
for idx, url in enumerate(all_product_urls, start=1):
    try:
        print(f"[{idx}] Processing: {url}")

        response = requests.get(url, timeout=15)
        soup = BeautifulSoup(response.text, "html.parser")

        # ---------- BASIC FIELDS ----------
        product_name = soup.find("span", id="ItemName")
        product_name = product_name.text.strip() if product_name else ""

        sku_tag = soup.find("span", id="ItemNumber")
        sku = sku_tag.text.strip() if sku_tag else ""

        sku_hidden = soup.find("input", id="hSKU")
        if sku_hidden and sku_hidden.get("value"):
            sku = sku_hidden["value"].strip()

        # ---------- IMAGE URL ----------
        image_url = ""
        img_lg = soup.find("img", id="lgImg")
        if img_lg and img_lg.get("src"):
            image_url = img_lg["src"]

        if not image_url:
            img_lg2 = soup.find("img", class_="lgImg")
            if img_lg2 and img_lg2.get("src"):
                image_url = img_lg2["src"]

        if not image_url:
            container = soup.find("a", id="lnkLargeImg")
            if container:
                img_tag = container.find("img")
                if img_tag and img_tag.get("src"):
                    image_url = img_tag["src"]

        if not image_url:
            img_def = soup.find("img", class_="prodSearchImage")
            if img_def and img_def.get("src"):
                image_url = img_def["src"]

        if not image_url:
            img_cont = soup.find("div", id="imgContainer")
            if img_cont:
                img_tag = img_cont.find("img")
                if img_tag and img_tag.get("src"):
                    image_url = img_tag.get("src")

        if not image_url:
            for im in soup.find_all("img"):
                src = im.get("src", "")
                if "prod-images" in src or "productcatalog" in src:
                    image_url = src
                    break

        if image_url and image_url.startswith("/"):
            image_url = "https://highlandhousefurniture.com" + image_url

        # ---------- DIMENSIONS ----------
        width = extract_dimension_value(soup, "width")
        depth = extract_dimension_value(soup, "depth")
        height = extract_dimension_value(soup, "height")
        diameter = extract_dimension_value(soup, "diameter")

        weight = ""
        weight_tag = soup.find("tr", id="weightRow")
        if weight_tag:
            td_tag = weight_tag.find("td", id="weight")
            if td_tag:
                weight = td_tag.text.strip()

        description = ""

        # ---------- FIXED FINISH + CUSHION EXTRACTION ----------
        finish_values = []
        cushion_values = []
        seen_finish = set()

        for row in soup.find_all("div", class_="configWrapper"):
            for d in row.find_all("div"):
                desc_tag = d.find("span", class_="asaDesc")
                val_tag = d.find("span", class_="asaVal")

                if not desc_tag or not val_tag:
                    continue

                desc = desc_tag.get_text(strip=True).lower()
                val = val_tag.get_text(strip=True)
                val = re.sub(r"\s+", " ", val)

                if "cushion" in desc:
                    cushion_values.append(val)
                    continue

                if "finish" in desc:
                    if is_valid_finish(val) and val not in seen_finish:
                        seen_finish.add(val)
                        finish_values.append(val)
                    continue

        for div in soup.find_all("div"):
            txt = div.get_text(" ", strip=True)
            if txt.lower().startswith("finish"):
                after = txt.split(":", 1)[1].strip() if ":" in txt else ""
                after = re.split(r"[.,;/]", after)[0].strip()
                after = re.sub(r"\s+", " ", after)
                if is_valid_finish(after) and after not in seen_finish:
                    seen_finish.add(after)
                    finish_values.append(after)

        if not finish_values and sku:
            suffix = sku.split("-")[-1].upper()
            finish_map = {
                "EB": "Ebony",
                "CT": "City Light",
                "BC": "Blonde Cerused",
                "DW": "Dark Walnut",
                "OY": "Oyster",
                "WA": "Washed Almond",
                "PT": "Parchment",
                "WH": "White",
                "BK": "Black"
            }
            if suffix in finish_map:
                finish_values.append(finish_map[suffix])

        finish = ", ".join(finish_values)
        color = finish
        cushion = ", ".join(cushion_values)

        # ---------- DIMENSION STRING ----------
        dimensions = ""
        dimension_div = soup.find("div", id="dimensionDiv")
        if dimension_div:
            dimensions = dimension_div.get_text(" ", strip=True)

        length = depth

        # ---------- RESTORED FIELDS ----------
        def cell(id_name):
            td = soup.find("td", id=id_name)
            return td.get_text(strip=True) if td else ""

        seat_number = cell("seatNumber")
        com_available = cell("comAvail")
        com = cell("COM")
        col = cell("COL")
        cot = cell("COT")
        arm_height = cell("armHeight")
        seat_height = cell("seatHeight")
        seat_depth = cell("seatDepth")

        # BASE FOOT TYPE FROM NOTES
        base_foot_type = ""
        notes_td = soup.find("td", id="iNotes")
        if notes_td:
            notes = notes_td.get_text(" ", strip=True)
            matches = re.findall(r"([\w\s\-]+?) on Base", notes)
            cleaned = []
            for m in matches:
                m = m.strip().rstrip(",.")
                if m and m not in cleaned:
                    cleaned.append(m)
            base_foot_type = ", ".join(cleaned)

        # ---------- SAVE ----------
        all_data.append({
            "Product URL": url,
            "Product Name": product_name,
            "SKU": sku,
            "Image URL": image_url,
            "Description": description,
            "Finish": finish,
            "Color": color,
            "Cushion": cushion,
            "Dimensions": dimensions,
            "Length": length,
            "Weight": weight,
            "Width": width,
            "Depth": depth,
            "Diameter": diameter,
            "Height": height,

            # RESTORED FIELDS
            "Seat Number": seat_number,
            "Base/Foot Type": base_foot_type,
            "COM Available": com_available,
            "COM": com,
            "COL": col,
            "COT": cot,
            "Arm Height": arm_height,
            "Seat Height": seat_height,
            "Seat Depth": seat_depth
        })

        print(f" → FIN: {finish} | CUSH: {cushion} | BASE: {base_foot_type}")

        if idx % batch_size == 0 or idx == len(all_product_urls):
            df_out = pd.DataFrame(all_data)
            df_out.to_excel(output_file, index=False)
            print("Saved batch.")

        time.sleep(1)

    except Exception as e:
        print(f"Error scraping {url}: {e}")

print("\nAll product details scraped successfully!")
