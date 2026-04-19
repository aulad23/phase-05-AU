# step3_FINAL_UNBREAKABLE.py
# ✅ Google search করে class="b8lM7" থেকে first linkedin.com/in URL নেবে
# ✅ CAPTCHA এ গেলে pause করবে → আপনি solve করবেন → তারপর automatically results detect করে same row complete করবে
# ✅ Title থেকে নাম extract করে "Contact Name" এ বসাবে (যদি Contact Name = LinkedIn Member)
# ✅ Chrome window close/crash হলে auto-restart করে same row retry করবে
# ✅ Resume + Auto-save

import pandas as pd
import time, random, os, re
from urllib.parse import quote_plus, urlparse, parse_qs, unquote

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# ================= CONFIG =================
INPUT_FILE = "linkedin_company_people_all.xlsx"
OUTPUT_FILE = "output_with_linkedin.xlsx"
SAVE_EVERY = 10

HEADLESS = False                 # ✅ CAPTCHA solve দরকার, তাই False
WAIT_TIMEOUT = 18

PAGE_WAIT_RANGE = (6, 10)
LOOP_DELAY_RANGE = (9, 16)

# CAPTCHA solve করার পরে result page আসা পর্যন্ত wait/check
MAX_CAPTCHA_WAIT_LOOPS = 90      # loops * ~2s = ~3 মিনিট
CAPTCHA_CHECK_SLEEP = (1.6, 2.4)

# ================= GLOBALS =================
driver = None
wait = None


# ================= UTIL =================
def human_sleep(a, b):
    time.sleep(random.uniform(a, b))

def norm(s):
    if not s:
        return ""
    return re.sub(r"\s+", " ", str(s).strip().lower())

def is_linkedin_member(name):
    return norm(name) in ["linkedin member", "member", "linkedin", ""]

def build_query(contact_name, designation, company):
    if is_linkedin_member(contact_name):
        return f'{company} {designation} LinkedIn'
    return f'{contact_name} {designation} {company} LinkedIn'

def extract_name_from_title(title):
    if not title:
        return ""
    title = re.sub(r"\.\.\.$", "", title).strip()
    if " - " in title:
        return title.split(" - ", 1)[0].strip()
    return ""

def extract_linkedin_from_google_redirect(href):
    try:
        qs = parse_qs(urlparse(href).query)
        real = qs.get("url", [""])[0]
        return unquote(real)
    except:
        return ""

def google_captcha_detected():
    global driver
    try:
        src = norm(driver.page_source)
        cur = (driver.current_url or "").lower()
        return (
            "google.com/sorry" in cur
            or "unusual traffic" in src
            or "automated queries" in src
            or "our systems have detected unusual traffic" in src
        )
    except:
        return True


# ================= DRIVER =================
def ensure_driver():
    global driver, wait

    # already ok?
    try:
        if driver is not None:
            _ = driver.current_url
            return
    except:
        pass

    chrome_options = Options()
    if HEADLESS:
        chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--remote-debugging-port=9222")

    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    wait = WebDriverWait(driver, WAIT_TIMEOUT)


# ================= GOOGLE PARSER (b8lM7) =================
def google_get_first_linkedin_from_b8lM7():
    """
    ✅ আপনার দেওয়া HTML অনুযায়ী:
    div.b8lM7 এর ভিতর থেকে first linkedin.com/in URL + h3 title
    """
    global driver, wait

    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.b8lM7")))
    except:
        return "", ""

    blocks = driver.find_elements(By.CSS_SELECTOR, "div.b8lM7")

    for block in blocks:
        try:
            # direct linkedin link inside block (best)
            a = None
            try:
                a = block.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]')
            except:
                a = block.find_element(By.CSS_SELECTOR, 'a[href]')

            href = a.get_attribute("href") or ""
            final = href

            # google redirect -> real url
            low = href.lower()
            if "google.com/url" in low and "url=" in low:
                real = extract_linkedin_from_google_redirect(href)
                if real:
                    final = real

            if "linkedin.com/in/" not in final.lower():
                # try any linkedin profile link inside
                a2 = block.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]')
                final = a2.get_attribute("href") or ""

            if "linkedin.com/in/" in final.lower():
                title = ""
                try:
                    h3 = block.find_element(By.CSS_SELECTOR, "h3.LC20lb")
                    title = h3.text.strip()
                except:
                    pass
                return final, title
        except:
            continue

    return "", ""


def wait_until_results_after_captcha():
    """
    ✅ CAPTCHA solve করার পরে:
    - sorry page চলে গেছে কিনা
    - b8lM7 results আসা পর্যন্ত wait
    """
    for _ in range(MAX_CAPTCHA_WAIT_LOOPS):
        if not google_captcha_detected():
            url, title = google_get_first_linkedin_from_b8lM7()
            if url:
                return url, title
        human_sleep(*CAPTCHA_CHECK_SLEEP)
    return "", ""


def google_search_with_manual_captcha(query):
    """
    ✅ CAPTCHA এ গেলে pause করে user solve করাবে
    ✅ তারপর result page detect করে same row complete করবে
    """
    global driver
    ensure_driver()

    search_url = "https://www.google.com/search?q=" + quote_plus(query)

    try:
        driver.get(search_url)
    except (NoSuchWindowException, WebDriverException):
        ensure_driver()
        driver.get(search_url)

    human_sleep(*PAGE_WAIT_RANGE)

    if google_captcha_detected():
        print("⚠️ Google CAPTCHA এসেছে.")
        print("👉 Chrome window বন্ধ করবেন না.")
        print("👉 CAPTCHA solve করুন + results page আসা পর্যন্ত অপেক্ষা করুন.")
        input("Press Enter after solving CAPTCHA...")

        # Solve করার পর real results আসা পর্যন্ত wait
        url, title = wait_until_results_after_captcha()
        return url, title

    return google_get_first_linkedin_from_b8lM7()


def find_linkedin(company, designation, contact_name):
    query = build_query(contact_name, designation, company)
    url, title = google_search_with_manual_captcha(query)
    return url, title, "Google"


# ================= LOAD / RESUME =================
if os.path.exists(OUTPUT_FILE):
    df = pd.read_excel(OUTPUT_FILE)
    print("🔁 Resuming from:", OUTPUT_FILE)
else:
    df = pd.read_excel(INPUT_FILE)
    if "Contact LinkedIn URL" not in df.columns:
        df["Contact LinkedIn URL"] = ""
    if "Matched Title" not in df.columns:
        df["Matched Title"] = ""
    if "Source" not in df.columns:
        df["Source"] = ""

# ================= RUN =================
ensure_driver()

try:
    for i, row in df.iterrows():
        existing = str(row.get("Contact LinkedIn URL", "")).strip()
        if existing and existing.lower() != "nan":
            continue

        company = str(row["Company Name"]).strip()
        designation = str(row["Designation"]).strip()
        contact_name = str(row["Contact Name"]).strip()

        print(f"\n🔍 [{i+1}] {contact_name} | {designation} | {company}")

        try:
            linkedin_url, title, source = find_linkedin(company, designation, contact_name)
        except (NoSuchWindowException, WebDriverException):
            print("⚠️ Chrome window closed/crashed. Restarting driver and retrying this row…")
            ensure_driver()
            linkedin_url, title, source = find_linkedin(company, designation, contact_name)

        if linkedin_url:
            df.at[i, "Contact LinkedIn URL"] = linkedin_url
            df.at[i, "Matched Title"] = title
            df.at[i, "Source"] = source

            extracted_name = extract_name_from_title(title)
            if extracted_name and is_linkedin_member(contact_name):
                df.at[i, "Contact Name"] = extracted_name
                print(f"   🧠 Name updated → {extracted_name}")

            print(f"   ✅ SAVED ({source}): {linkedin_url}")
            if title:
                print("   🧾 Title:", title)
        else:
            df.at[i, "Source"] = source
            print("   ❌ NOT FOUND (CAPTCHA solved but results not detected / no b8lM7 found)")

        if (i + 1) % SAVE_EVERY == 0:
            df.to_excel(OUTPUT_FILE, index=False)
            print(f"💾 Auto-saved at row {i+1}")

        human_sleep(*LOOP_DELAY_RANGE)

    df.to_excel(OUTPUT_FILE, index=False)
    print("\n✅ DONE! File saved:", OUTPUT_FILE)

finally:
    try:
        if driver:
            driver.quit()
    except:
        pass
