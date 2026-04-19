import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ========== CONFIG ==========

BASE_URL = "https://www.crlaine.com"

# 🔹 Script er folder (jeikhane .py file ache)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 🔹 Input/Output same folder e
INPUT_FILE  = os.path.join(SCRIPT_DIR, "Crlaine_Trims.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "Crlaine_Trims_Details.xlsx")

HEADLESS = False          # ✅ Window open korar jonno False
WAIT_TIME = 2.5           # seconds between page loads
SAVE_BATCH_SIZE = 5       # save every 5 rows to avoid data loss
# ============================


def setup_driver(headless=True):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def extract_specification(page_source):
    """
    prod-details pura block-er text Specification column-e rakhbo.
    <div class="prod-details"> ... </div> theke sob text niye flat string banabo.
    """
    soup = BeautifulSoup(page_source, "html.parser")
    block = soup.find("div", class_="prod-details")
    if not block:
        return ""
    # sob text ekshathe, extra gap remove kore
    spec_text = " ".join(block.stripped_strings)
    return spec_text


def extract_description(page_source):
    """Comments block theke Description make korbo (semi-colon diye join)."""
    soup = BeautifulSoup(page_source, "html.parser")
    desc_texts = []

    comments_block = soup.find("div", class_="commentsBlock")
    if comments_block:
        lis = comments_block.find_all("li")
        for li in lis:
            t = li.get_text(strip=True)
            if t:
                desc_texts.append(t)

    description = "; ".join(desc_texts)
    return description


def extract_contents_and_repeat(page_source):
    """
    dimtable block theke:
    Contents: 100%  SUNBRELLA SOLUTION DYED ACRYLIC
    Repeat:   V 1.27 x  H 1.25
    """
    soup = BeautifulSoup(page_source, "html.parser")
    contents = ""
    repeat = ""

    # dimtable div
    dimtable = soup.find("div", class_="dimtable")
    if not dimtable:
        # fallback jodi class combination thake
        dimtable = soup.find("div", class_="pure-u-1 dimtable")

    if not dimtable:
        return contents, repeat

    for div in dimtable.find_all("div"):
        label = div.find("span", class_="detailInfoLabel")
        if not label:
            continue

        label_text = label.get_text(strip=True).upper().rstrip(":")
        spans = div.find_all("span")
        if len(spans) < 2:
            continue

        val = spans[1].get_text(" ", strip=True)

        if "CONTENTS" in label_text and not contents:
            contents = val.strip()
        elif "REPEAT" in label_text and not repeat:
            repeat = val.strip()

    return contents, repeat


def scrape_fabric_details(driver, url):
    """Visit a single trim/fabric page and extract Specification + Description + Contents + Repeat."""
    try:
        full_url = urljoin(BASE_URL, url) if url.startswith("/") else url
        driver.get(full_url)
        time.sleep(WAIT_TIME)

        page_source = driver.page_source

        specification = extract_specification(page_source)
        desc = extract_description(page_source)
        contents, repeat = extract_contents_and_repeat(page_source)

        return {
            "Specification": specification.strip(),
            "Description": desc.strip(),
            "Contents": contents.strip(),
            "Repeat": repeat.strip(),
            "Weight": "",
            "Width": "",
            "Depth": "",
            "Diameter": "",
            "Height": ""
        }
    except Exception as e:
        print(f"⚠️ Error scraping {url}: {e}")
        return {
            "Specification": "",
            "Description": "",
            "Contents": "",
            "Repeat": "",
            "Weight": "",
            "Width": "",
            "Depth": "",
            "Diameter": "",
            "Height": ""
        }


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input file not found: {INPUT_FILE}")
        return

    df_input = pd.read_excel(INPUT_FILE)
    required_cols = ["Product URL", "Image URL", "Product Name", "SKU"]
    for col in required_cols:
        if col not in df_input.columns:
            df_input[col] = ""

    driver = setup_driver(headless=HEADLESS)
    results = []
    total = len(df_input)
    print(f"🚀 Starting detail scrape for {total} trims...")

    try:
        for idx, row in df_input.iterrows():
            url = str(row.get("Product URL", "")).strip()
            if not url:
                continue

            print(f"[{idx+1}/{total}] Scraping: {url}")
            details = scrape_fabric_details(driver, url)

            row_out = {
                "Product URL": row.get("Product URL", ""),
                "Image URL": row.get("Image URL", ""),
                "Product Name": row.get("Product Name", ""),
                # 🔹 Product Family Id = Product Name
                "Product Family Id": row.get("Product Name", ""),
                "SKU": row.get("SKU", ""),
                # 🔹 New big block
                "Specification": details.get("Specification", ""),
                # 🔹 From comments
                "Description": details.get("Description", ""),
                # 🔹 From dimtable
                "Contents": details.get("Contents", ""),
                "Repeat": details.get("Repeat", ""),
                "Weight": details.get("Weight", ""),
                "Width": details.get("Width", ""),
                "Depth": details.get("Depth", ""),
                "Diameter": details.get("Diameter", ""),
                "Height": details.get("Height", "")
            }
            results.append(row_out)

            # 🔸 Auto-save after every batch
            if (idx + 1) % SAVE_BATCH_SIZE == 0 or (idx + 1) == total:
                out_df = pd.DataFrame(results, columns=[
                    "Product URL", "Image URL", "Product Name", "Product Family Id", "SKU",
                    "Specification",
                    "Description", "Contents", "Repeat",
                    "Weight", "Width", "Depth", "Diameter", "Height"
                ])
                out_df.to_excel(OUTPUT_FILE, index=False)
                print(f"💾 Auto-saved {len(out_df)} rows to file.")

    finally:
        driver.quit()

    print(f"\n✅ Done! Total {len(results)} products saved to:\n{OUTPUT_FILE}")


if __name__ == "__main__":
    main()
