# -*- coding: utf-8 -*-
"""
Google Maps scraper — Asphalt & Paving (single-sheet, sequential per-origin)
- Processes each target address sequentially (finish one, then the next)
- Computes driving distance in miles (shortest route)
- NO optional radius filter
- Output: ONE Excel sheet ("Master") only
"""

import os
import re
import time
import random
import urllib3
import pandas as pd
import requests
from urllib.parse import urlparse
from geopy.geocoders import Nominatim

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================== CONFIG ==================
chrome_path = os.getenv("CHROMEDRIVER_PATH", r"C:\chromedriver.exe")

# Process these origins sequentially — one finishes, then the next starts.
TARGET_ADDRESSES = [
"Texarkana, TX 75503 USA"

]

KEYWORD = "land development projects"
RADIUS_MILES = 50                # only used in search text, NOT for filtering
HEADLESS = False                 # True = headless browser

OUTPUT_XLSX = f"asphalt_paving_master_land_development_projects.xlsx"

# Optional: Fixed coordinates per origin (skip geocoding if provided)
ORIGIN_LATLNG_OVERRIDES = {
    # "Texarkana, TX 75503 (USA)": (33.4633, -94.0792),
}

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
start_time = time.time()

# ================== UTILITIES ==================
def extract_domain(url: str) -> str:
    try:
        netloc = urlparse(url).netloc or ""
        return netloc.replace("www.", "")
    except Exception:
        return ""

def check_website_status(url: str) -> str:
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        }
        r = requests.get(url, headers=headers, allow_redirects=True, timeout=10, verify=False)
        return "Active" if r.status_code < 400 else "Inactive"
    except Exception:
        return "Inactive"

# Bengali → English digits
_BN2EN = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")

def bn_to_en(text: str) -> str:
    if not isinstance(text, str):
        return text
    return text.translate(_BN2EN)

def normalize(text: str) -> str:
    if not text:
        return ""
    return bn_to_en(text).strip()

def create_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(service=Service(chrome_path), options=opts)

# ================== GEO ==================
_geocode_cache = {}

def geocode_address(addr: str):
    addr_norm = normalize(addr)
    if not addr_norm:
        return None
    if addr_norm in _geocode_cache:
        return _geocode_cache[addr_norm]
    # override?
    normalized_overrides = {normalize(k): v for k, v in ORIGIN_LATLNG_OVERRIDES.items()}
    if addr_norm in normalized_overrides:
        _geocode_cache[addr_norm] = normalized_overrides[addr_norm]
        return _geocode_cache[addr_norm]

    try:
        geolocator = Nominatim(user_agent="maps_scraper_texarkana_single_sheet")
        loc = geolocator.geocode(addr_norm, timeout=12)
        if loc:
            _geocode_cache[addr_norm] = (loc.latitude, loc.longitude)
            return _geocode_cache[addr_norm]
    except Exception:
        pass
    return None

# ================== Directions distance (robust) ==================
def get_shortest_driving_distance_miles(driver, origin_q: str, dest_q: str,
                                        wait_secs: float = 7.5, max_wait: float = 18.0):
    """
    Opens Google Maps Directions in a TEMP TAB and parses route distances.
    Returns shortest distance (float, miles) or None.
    """
    try:
        maps_url = f"https://www.google.com/maps/dir/{origin_q}/{dest_q}/?hl=en&gl=us&language=en"
        driver.execute_script("window.open(arguments[0], '_blank');", maps_url)
        driver.switch_to.window(driver.window_handles[-1])

        t0 = time.time()
        distances = []

        def _bn_en(txt: str) -> str:
            return bn_to_en(txt) if isinstance(txt, str) else txt

        def try_collect_distances():
            found = []
            selectors = [
                "div.UgZKXd.clearfix.yYG3jf div.ivN21e",
                "div.UgZKXd.clearfix.yYG3jf span.ivN21e",
                "div.ivN21e",
                "div.MespJc", "span.MespJc",
                "div.Fk3sm", "span.Fk3sm",
                "div.LJKBpe", "span.LJKBpe",
            ]
            for css in selectors:
                try:
                    for el in driver.find_elements(By.CSS_SELECTOR, css):
                        txt = el.text.strip()
                        if not txt:
                            continue
                        txt = _bn_en(txt).lower()
                        m = re.search(r'(\d+(?:\.\d+)?)\s*(?:mi|mile|miles)\b', txt)
                        if m:
                            found.append(float(m.group(1)))
                except Exception:
                    pass

            if not found:
                try:
                    body_text = driver.execute_script("return document.body ? document.body.innerText : ''") or ""
                    body_text = _bn_en(body_text).lower()
                    for m in re.finditer(r'(\d+(?:\.\d+)?)\s*(?:mi|mile|miles)\b', body_text):
                        try:
                            found.append(float(m.group(1)))
                        except Exception:
                            continue
                except Exception:
                    pass
            return found

        time.sleep(wait_secs)
        distances.extend(try_collect_distances())
        while not distances and (time.time() - t0) < max_wait:
            time.sleep(2.0)
            distances.extend(try_collect_distances())

        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        if distances:
            distances = sorted(distances)[:3]
            return min(distances)

    except Exception:
        try:
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass
    return None

# ================== Core scrape for ONE origin ==================
def scrape_for_origin(origin_address: str, keyword: str, radius_miles: int = 50):
    driver = create_driver()
    rows = []
    try:
        origin_latlng = geocode_address(origin_address)
        if not origin_latlng:
            raise RuntimeError(f"Could not geocode origin: {origin_address}")
        origin_query = f"{origin_latlng[0]},{origin_latlng[1]}".replace(" ", "+")

        driver.get("https://www.google.com/maps?hl=en")
        wait = WebDriverWait(driver, 25)

        # Search box text (for context only; no filtering in code)
        search_text = f'{keyword} near "{origin_address}" within {radius_miles} miles'
        search_box = wait.until(EC.presence_of_element_located((By.ID, "searchboxinput")))
        search_box.clear()
        search_box.send_keys(search_text)
        search_box.send_keys(Keys.ENTER)

        # Wait for list feed and scroll progressively to load many cards
        scrollable = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="feed"]')))
        prev_count, stalls = 0, 0
        for _ in range(40):
            driver.execute_script("arguments[0].scrollBy(0, 1400);", scrollable)
            time.sleep(random.uniform(1.2, 2.0))
            listings = driver.find_elements(By.CSS_SELECTOR, "div.Nv2PK")
            if len(listings) == prev_count:
                stalls += 1
                if stalls > 3:
                    break
            else:
                prev_count = len(listings)
                stalls = 0

        listings = driver.find_elements(By.CSS_SELECTOR, "div.Nv2PK")
        print(f"🔎 {origin_address}: {len(listings)} listings found for '{keyword}'")

        for r in listings:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", r)
                time.sleep(0.1)
            except Exception:
                pass

            name = rating = reviews_number = specialty = ""
            address_text = phone = website_link = gmb_link = ""
            city = state = zip_code_extracted = ""

            # Name
            try:
                name = r.find_element(By.CLASS_NAME, "qBF1Pd").text
            except Exception:
                pass

            # Rating
            try:
                rating_elem = r.find_element(By.CSS_SELECTOR, 'span[aria-label*="stars"]')
                rating_raw = rating_elem.get_attribute("aria-label").split()[0]
                rating = normalize(rating_raw)
            except Exception:
                pass

            # Reviews
            try:
                reviews_elem = r.find_element(By.CSS_SELECTOR, 'div.W4Efsd div.AJB7ye span.UY7F9')
                reviews_number = normalize(reviews_elem.text.strip().strip("()"))
            except Exception:
                pass

            # Specialty + address (preview)
            try:
                block = r.find_element(By.CSS_SELECTOR, "div.W4Efsd > div.W4Efsd")
                all_spans = block.find_elements(By.TAG_NAME, "span")
                for span in all_spans:
                    child_spans = span.find_elements(By.TAG_NAME, "span")
                    if child_spans:
                        specialty = normalize(child_spans[0].text)
                        break
                for span in reversed(all_spans):
                    inner_spans = span.find_elements(By.TAG_NAME, "span")
                    for s in inner_spans:
                        if not s.get_attribute("aria-hidden"):
                            address_text = normalize(s.text)
                            break
                    if address_text:
                        break
            except Exception:
                pass

            # Phone
            try:
                phone = normalize(r.find_element(By.CLASS_NAME, "UsdlK").text)
            except Exception:
                pass

            # Website (from card)
            try:
                website_elem = r.find_element(By.CSS_SELECTOR, 'a[data-value="Website"]')
                website_link = extract_domain(website_elem.get_attribute("href"))
            except Exception:
                pass

            # GMB link
            try:
                gmb_elem = r.find_element(By.CSS_SELECTOR, "a.hfpxzc")
                gmb_link = gmb_elem.get_attribute("href") or gmb_link
            except Exception:
                pass

            # Open details to capture full address/website more reliably
            try:
                if gmb_link:
                    driver.execute_script("arguments[0].click();", r.find_element(By.CSS_SELECTOR, "a.hfpxzc"))
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.bJzME.Hu9e2e.tTVLSc")))
                    time.sleep(0.8)

                    buttons = driver.find_elements(By.CSS_SELECTOR, "div.bJzME.Hu9e2e.tTVLSc button.CsEnBe")
                    for btn in buttons:
                        item_id = btn.get_attribute("data-item-id")
                        if item_id == "address":
                            full_addr = normalize(btn.get_attribute("aria-label"))
                            if full_addr:
                                address_text = full_addr
                                m_city_state_zip = re.search(r",\s*([^,]+),\s*([A-Z]{2})\s+(\d{5})", full_addr)
                                if m_city_state_zip:
                                    city = normalize(m_city_state_zip.group(1))
                                    state = normalize(m_city_state_zip.group(2))
                                    zip_code_extracted = normalize(m_city_state_zip.group(3))
                        elif item_id == "authority" and not website_link:
                            raw_url = btn.get_attribute("href")
                            website_link = extract_domain(raw_url)

                    driver.back()
                    wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="feed"]')))
                    time.sleep(0.3)
            except Exception:
                try:
                    driver.back()
                    wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="feed"]')))
                    time.sleep(0.3)
                except Exception:
                    pass

            # Website status
            status = "No Website"
            if website_link:
                status = check_website_status("http://" + website_link)

            # Driving distance via Directions (shortest route)
            dest_query = (address_text or name).replace(" ", "+")
            distance_miles = get_shortest_driving_distance_miles(driver, origin_query, dest_query)
            if distance_miles is not None and distance_miles > 500:
                distance_miles = None  # sanity check

            rows.append({
                "Origin_Address": normalize(origin_address),
                "Name": normalize(name),
                "Rating": rating,
                "Reviews Numbers": reviews_number,
                "Specialty": normalize(specialty),
                "Address": normalize(address_text),
                "City": normalize(city),
                "State": normalize(state),
                "ZipCode": normalize(zip_code_extracted),
                "Phone": normalize(phone),
                "Website": website_link,
                "GMB_Link": gmb_link,
                "Website_Status": status,
                "Distance_from_Target_Miles": distance_miles
            })

        return rows

    finally:
        driver.quit()

# ================== MAIN (sequential per-origin, single-sheet) ==================
if __name__ == "__main__":
    all_rows = []

    for i, origin in enumerate(TARGET_ADDRESSES, 1):
        print(f"\n===== [{i}/{len(TARGET_ADDRESSES)}] Starting: {origin} =====")
        try:
            rows = scrape_for_origin(origin, KEYWORD, RADIUS_MILES)
            all_rows.extend(rows)
            # short cool-down between origins to reduce blocks
            time.sleep(random.uniform(2.5, 4.5))
        except Exception as e:
            print(f"❌ Error for origin '{origin}': {e}")
        print(f"===== Finished: {origin} =====")

    master_df = pd.DataFrame(all_rows)

    # Normalize Bengali digits & trim; sort by Origin then Distance
    if not master_df.empty:
        for col in master_df.columns:
            if master_df[col].dtype == "object":
                master_df[col] = master_df[col].astype(str).apply(bn_to_en).str.strip()
        if "Distance_from_Target_Miles" in master_df.columns:
            master_df = master_df.sort_values(
                by=["Origin_Address", "Distance_from_Target_Miles"],
                na_position="last"
            )

    # Save ONE sheet only
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        (master_df if not master_df.empty else pd.DataFrame()).to_excel(
            writer, sheet_name="Master", index=False
        )

    mins, secs = divmod(time.time() - start_time, 60)
    total_rows = 0 if master_df is None else len(master_df)
    print(f"\n✅ Saved Excel with {total_rows} rows → {OUTPUT_XLSX}")
    print(f"⏱️ Total time: {int(mins)} minutes {int(secs)} seconds")
