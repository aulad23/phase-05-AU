# -*- coding: utf-8 -*-
# linkedin_company_people_all_FINAL.py
#
# INPUT  : linkedin_company_urls.xlsx  (must have column: LinkedIn URL)
# OUTPUT : linkedin_company_people_all.xlsx
#
# ✅ Keeps "LinkedIn URL" in output (company url)
# ✅ Scrapes About: Website, Phone, Industry, Address
# ✅ Scrapes People: Number of Employee (associated members)
# ✅ Scrapes ALL contacts (including private "LinkedIn Member" cards)
# ✅ One row per contact (company info repeated)


import time
import re
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ================= CONFIG =================
INPUT_FILE = "linkedin_urls.xlsx"      # Column: LinkedIn URL
OUTPUT_FILE = "linkedin_company_people_all.xlsx"

WAIT_SEC = 20
POLITE_DELAY = 1.2
PEOPLE_SCROLL_ROUNDS = 10  # increase to load more people cards


# ============== DRIVER ==============
def create_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    # options.add_argument("--headless=new")  # LinkedIn often blocks headless
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    driver.set_page_load_timeout(60)
    return driver


# ============== HELPERS ==============
def sleep_polite():
    time.sleep(POLITE_DELAY)

def text_clean(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def clean_profile_url(url: str) -> str:
    """Normalize /in/ urls by removing miniProfileUrn/tracking params."""
    if not url:
        return ""
    url = url.strip()
    parts = urlsplit(url)
    q = dict(parse_qsl(parts.query))
    for k in list(q.keys()):
        if k.lower() in ["miniprofileurn", "lipi", "trk", "trkinfo", "originalreferer"]:
            q.pop(k, None)
    new_query = urlencode(q, doseq=True)
    clean = urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, ""))
    return clean.rstrip("/")

def parse_employee_count(text: str):
    if not text:
        return ""
    t = text.lower()
    m = re.search(r"(\d[\d,]*)\s+associated\s+members", t)
    if m:
        return int(m.group(1).replace(",", ""))
    m2 = re.search(r"(\d[\d,]*)", text)
    if m2:
        return int(m2.group(1).replace(",", ""))
    return ""

def parse_address_parts(addr: str):
    out = {
        "Street Address": "",
        "Street Address 02": "",
        "City": "",
        "State": "",
        "Zip Code": "",
        "Country": ""
    }
    if not addr:
        return out

    s = text_clean(addr)

    # Example: "7100 TPC Dr, Suite 100, Orlando, Florida 32822, US"
    m = re.search(
        r"^(.*?),\s*(.*?),\s*([^,]+),\s*([A-Za-z\s]+)\s+(\d{5}(?:-\d{4})?)\s*,\s*(.*)$",
        s
    )
    if m:
        out["Street Address"] = m.group(1).strip()
        out["Street Address 02"] = m.group(2).strip()
        out["City"] = m.group(3).strip()
        out["State"] = m.group(4).strip()
        out["Zip Code"] = m.group(5).strip()
        out["Country"] = m.group(6).strip()
        return out

    # fallback (generic)
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if parts:
        out["Street Address"] = parts[0]
    if len(parts) >= 2:
        out["Street Address 02"] = parts[1]
    if len(parts) >= 3:
        out["City"] = parts[-3]
    if len(parts) >= 2:
        out["Country"] = parts[-1]

    mz = re.search(r"\b(\d{5}(?:-\d{4})?)\b", s)
    if mz:
        out["Zip Code"] = mz.group(1)

    # if state name present, keep as-is (US states often full name here)
    return out


# ========= ABOUT PAGE =========
def get_about_field(driver: webdriver.Chrome, label_text: str) -> str:
    """
    LinkedIn About page often:
    <dl><dt><h3>Website</h3></dt><dd>...</dd></dl>
    """
    xpaths = [
        f"//dl//dt[.//h3[normalize-space()='{label_text}']]/following-sibling::dd[1]",
        f"//dt[.//h3[normalize-space()='{label_text}']]/following-sibling::dd[1]",
        f"//h3[normalize-space()='{label_text}']/ancestor::dt/following-sibling::dd[1]",
    ]
    for xp in xpaths:
        try:
            el = driver.find_element(By.XPATH, xp)
            txt = text_clean(el.text)
            if txt:
                return txt
        except Exception:
            pass
    return ""

def get_company_name(driver: webdriver.Chrome) -> str:
    try:
        h1 = driver.find_element(By.CSS_SELECTOR, "h1.org-top-card-summary__title")
        return text_clean(h1.text)
    except Exception:
        try:
            h1 = driver.find_element(By.TAG_NAME, "h1")
            return text_clean(h1.text)
        except Exception:
            return ""

def get_company_address(driver: webdriver.Chrome) -> str:
    # usually p.break-words
    candidates = driver.find_elements(By.CSS_SELECTOR, "p.t-14.t-black--light.t-normal.break-words")
    for p in candidates:
        txt = text_clean(p.text)
        if txt and ("," in txt) and re.search(r"\d", txt):
            return txt
    try:
        p = driver.find_element(By.CSS_SELECTOR, "p.break-words")
        return text_clean(p.text)
    except Exception:
        return ""


# ========= PEOPLE PAGE =========
def get_associated_members_count(driver: webdriver.Chrome) -> str:
    try:
        h2 = driver.find_element(
            By.XPATH,
            "//h2[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'associated members')]"
        )
        return parse_employee_count(h2.text)
    except Exception:
        return ""

def scroll_people_page(driver: webdriver.Chrome):
    for _ in range(PEOPLE_SCROLL_ROUNDS):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.1)
        try:
            btn = driver.find_element(By.CSS_SELECTOR, "button.scaffold-finite-scroll__load-button")
            if btn.is_displayed() and btn.is_enabled():
                btn.click()
                time.sleep(1.2)
        except Exception:
            pass

def scrape_all_people_cards_including_private(driver: webdriver.Chrome):
    """
    Scrape ALL cards including private 'LinkedIn Member' cards.
    For private cards:
      - Name => 'LinkedIn Member (Private)' (or whatever visible)
      - URL  => miniProfileUrn link if exists, else blank
    """
    people = []
    cards = driver.find_elements(By.CSS_SELECTOR, "li.org-people-profile-card__profile-card-spacing")

    for c in cards:
        try:
            # 1) Try normal /in/ url
            profile_url = ""
            a_in = c.find_elements(By.CSS_SELECTOR, "a[href*='/in/']")
            if a_in:
                profile_url = clean_profile_url(a_in[0].get_attribute("href") or "")

            # 2) Fallback: miniProfileUrn
            if not profile_url:
                a_mini = c.find_elements(By.CSS_SELECTOR, "a[href*='miniProfileUrn=']")
                if a_mini:
                    profile_url = a_mini[0].get_attribute("href") or ""

            # Name (visible or LinkedIn Member)
            name = ""
            try:
                name_el = c.find_element(By.CSS_SELECTOR, ".artdeco-entity-lockup__title .lt-line-clamp--single-line")
                name = text_clean(name_el.text)
            except Exception:
                name = ""

            if not name:
                name = "LinkedIn Member (Private)"

            # Designation/headline
            designation = ""
            try:
                des_el = c.find_element(By.CSS_SELECTOR, ".artdeco-entity-lockup__subtitle .lt-line-clamp--multi-line")
                designation = text_clean(des_el.text)
            except Exception:
                try:
                    des_el = c.find_element(By.CSS_SELECTOR, "div.lt-line-clamp--multi-line")
                    designation = text_clean(des_el.text)
                except Exception:
                    designation = ""

            people.append({
                "Name": name,
                "Designation": designation,
                "Contact Linkedin URL": profile_url,
                "Parsonal Email": ""
            })

        except Exception:
            pass

    # de-dupe: if URL missing, use name+designation
    uniq = {}
    for p in people:
        key = p.get("Contact Linkedin URL") or (p.get("Name", "") + "|" + p.get("Designation", ""))
        if key not in uniq:
            uniq[key] = p
    return list(uniq.values())


# ============== MAIN ==============
def main():
    driver = create_driver()
    wait = WebDriverWait(driver, WAIT_SEC)

    # Manual login
    driver.get("https://www.linkedin.com/feed/")
    input("🔑 Login manually, then press ENTER here...")

    df = pd.read_excel(INPUT_FILE)
    if "LinkedIn URL" not in df.columns:
        raise ValueError("INPUT_FILE must have a column named: LinkedIn URL")

    rows = []

    for _, r in df.iterrows():
        company_base = str(r.get("LinkedIn URL", "")).strip().rstrip("/")
        if not company_base or company_base.lower() == "nan":
            continue

        about_url = company_base + "/about/"
        people_url = company_base + "/people/"

        # Defaults
        company_name = ""
        website = ""
        phone = ""
        industry = ""
        num_emp = ""

        address_full = ""
        address_parts = {
            "Street Address": "",
            "Street Address 02": "",
            "City": "",
            "State": "",
            "Zip Code": "",
            "Country": ""
        }

        # ------- ABOUT -------
        try:
            driver.get(about_url)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

            company_name = get_company_name(driver)
            address_full = get_company_address(driver)
            address_parts = parse_address_parts(address_full)

            website = get_about_field(driver, "Website")
            phone = get_about_field(driver, "Phone")
            industry = get_about_field(driver, "Industry")

        except Exception:
            pass

        sleep_polite()

        # ------- PEOPLE -------
        people = []
        try:
            driver.get(people_url)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

            num_emp = get_associated_members_count(driver)

            scroll_people_page(driver)
            people = scrape_all_people_cards_including_private(driver)

        except Exception:
            people = []

        # ------- ROWS (one row per person) -------
        if not people:
            rows.append({
                "LinkedIn URL": company_base,
                "Company Name": company_name,
                "Website": website,
                "Company Phone Number": phone,
                "Industry": industry,
                "Number of Employee": num_emp,
                "Street Address": address_parts["Street Address"],
                "Street Address 02": address_parts["Street Address 02"],
                "City": address_parts["City"],
                "State": address_parts["State"],
                "Zip Code": address_parts["Zip Code"],
                "Country": address_parts["Country"],
                "Name": "",
                "Director Level or Up": "",
                "Designation": "",
                "Contact Linkedin URL": "",
                "Parsonal Email": ""
            })
        else:
            for p in people:
                rows.append({
                    "LinkedIn URL": company_base,
                    "Company Name": company_name,
                    "Website": website,
                    "Company Phone Number": phone,
                    "Industry": industry,
                    "Number of Employee": num_emp,
                    "Street Address": address_parts["Street Address"],
                    "Street Address 02": address_parts["Street Address 02"],
                    "City": address_parts["City"],
                    "State": address_parts["State"],
                    "Zip Code": address_parts["Zip Code"],
                    "Country": address_parts["Country"],
                    "Name": p["Name"],
                    "Director Level or Up": "",   # you want ALL contacts
                    "Designation": p["Designation"],
                    "Contact Linkedin URL": p["Contact Linkedin URL"],
                    "Parsonal Email": ""
                })

        # Save after each company to avoid data loss
        pd.DataFrame(rows).to_excel(OUTPUT_FILE, index=False)
        print(f"✅ Saved: {company_name or company_base} | total rows: {len(rows)}")

        sleep_polite()

    driver.quit()
    print(f"\n✅ DONE. Output: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
