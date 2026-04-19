import pandas as pd
import re
from urllib.parse import urljoin, urldefrag
from collections import deque
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from tqdm import tqdm
import xlwings as xw
import os
import string

INPUT_XLSX = r"C:\Users\Acord Tech Solutions\Downloads\test.xlsx"
INPUT_COL = "organization_website_url"
OUTPUT_XLSX = r"C:\Users\Acord Tech Solutions\Downloads\test_results.xlsx"

SPECIFIC_KEYWORDS = [
    "Home Services", "HVAC", "PET", "NURSING", "DOG", "CAT",
    "Air Condition", "Heating", "Ventilation", "plumbers",
    "electric", "pest control", "roofing", "carpet cleaning"
]

BATCH_SIZE = 5
SLEEP_BETWEEN_REQUESTS = 1
MAX_PAGES = 30
PAGE_LOAD_TIMEOUT = 40

def clean_text(text):
    """Lowercase, remove punctuation and extra spaces."""
    text = text.lower()
    text = text.translate(str.maketrans("", "", string.punctuation))
    text = " ".join(text.split())
    return text

def create_driver():
    """Configures and creates a headless Selenium WebDriver instance."""
    options = Options()
    options.headless = True
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--log-level=3")  # Suppress browser logs
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver

def normalize_link(base_url, link):
    link, _ = urldefrag(urljoin(base_url, link))
    return link.rstrip("/")

def crawl_site(url):
    """Crawls a website to find keywords on its pages."""
    result = {"url": url}
    for kw in SPECIFIC_KEYWORDS:
        result[kw] = "No"
        result[f"{kw} Source"] = ""

    visited = set()
    queue = deque([url])
    driver = create_driver()

    try:
        while queue and len(visited) < MAX_PAGES:
            current_url = queue.popleft()
            if current_url in visited:
                continue
            visited.add(current_url)

            try:
                driver.get(current_url)
                WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                time.sleep(SLEEP_BETWEEN_REQUESTS)
            except TimeoutException:
                print(f"⚠ Timeout loading {current_url}")
                continue
            except WebDriverException as e:
                print(f"Error fetching {current_url}: {e}")
                continue

            html = driver.page_source
            soup = BeautifulSoup(html, "lxml")

            text_parts = []
            for tag in ["p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "span", "div", "a"]:
                for element in soup.find_all(tag):
                    text_parts.append(element.get_text(" ", strip=True))
            full_text = " ".join(text_parts)

            cleaned_full_text = clean_text(full_text)

            for kw in SPECIFIC_KEYWORDS:
                if kw == "Home Services":
                    # Home Service| Home Service
                    variants = ["home services", "home service" ,]
                    for variant in variants:
                        if variant in cleaned_full_text and result[kw] == "No":
                            result[kw] = "Yes"
                            result[f"{kw} Source"] = current_url
                            print(f"✅ Found '{variant}' at {current_url}")
                            break

                elif kw == "plumbers":
                    # Accept plumber | plumbers | plumber service | plumbers service
                    variants = ["plumber", "plumbers", "plumber service", "plumbers service"]
                    for variant in variants:
                        if variant in cleaned_full_text and result[kw] == "No":
                            result[kw] = "Yes"
                            result[f"{kw} Source"] = current_url
                            print(f"✅ Found '{variant}' at {current_url}")
                            break

                else:
                   
                    kw_clean = clean_text(kw)
                    if kw_clean in cleaned_full_text and result[kw] == "No":
                        result[kw] = "Yes"
                        result[f"{kw} Source"] = current_url
                        print(f"✅ Found '{kw}' at {current_url}")

            for a in soup.find_all("a", href=True):
                link = normalize_link(url, a["href"])
                if link.startswith(url) and link not in visited and len(visited) + len(queue) < MAX_PAGES:
                    queue.append(link)
    finally:
        driver.quit()
    return result

def main():
    try:
        df = pd.read_excel(INPUT_XLSX, engine="openpyxl")
        df.columns = df.columns.str.strip()
    except FileNotFoundError:
        print(f"Error: The input file '{INPUT_XLSX}' was not found.")
        return

    if INPUT_COL not in df.columns:
        print(f"Error: The required column '{INPUT_COL}' was not found in the input file.")
        return

    if os.path.exists(OUTPUT_XLSX):
        try:
            existing_df = pd.read_excel(OUTPUT_XLSX, engine="openpyxl")
            existing_df.columns = existing_df.columns.str.strip()

            if SPECIFIC_KEYWORDS:
                processed_urls = existing_df[existing_df[SPECIFIC_KEYWORDS[0]] == "Yes"][INPUT_COL].tolist()
            else:
                processed_urls = existing_df[INPUT_COL].tolist()

            df = df.merge(existing_df, how="left", on=INPUT_COL)
        except Exception as e:
            print(f"Error reading existing output file: {e}. Starting fresh.")
            processed_urls = []
    else:
        processed_urls = []

    urls_to_process = df[~df[INPUT_COL].isin(processed_urls)][INPUT_COL].dropna().astype(str).tolist()
    print(f"Total URLs to process: {len(urls_to_process)}")

    try:
        wb = xw.Book(OUTPUT_XLSX)
    except Exception:
        wb = xw.Book()
        wb.save(OUTPUT_XLSX)

    sheet = wb.sheets[0]
    results = []

    for i in tqdm(range(0, len(urls_to_process), BATCH_SIZE), desc="Processing batches"):
        batch_urls = urls_to_process[i:i + BATCH_SIZE]
        clean_batch_urls = [(url if url.startswith("http") else "http://" + url) for url in batch_urls]

        batch_results = []
        with ThreadPoolExecutor(max_workers=BATCH_SIZE) as executor:
            futures = {executor.submit(crawl_site, url): url for url in clean_batch_urls}
            for future in as_completed(futures):
                try:
                    res = future.result()
                    batch_results.append(res)
                    print(f"Processed: {res['url']}")
                except Exception as e:
                    print(f"Error in crawling {futures[future]}: {e}")

        results.extend(batch_results)

        out_df = pd.DataFrame(results)
        merged = df.merge(out_df, how="left", left_on=INPUT_COL, right_on="url")
        merged.drop(columns=["url"], inplace=True)

        sheet.clear()
        sheet.range("A1").value = merged
        wb.save()
        print(f"\n📊 Batch {i // BATCH_SIZE + 1} saved to {OUTPUT_XLSX}")

    print(f"\n✅ Finished. Results saved to {OUTPUT_XLSX}")
    wb.app.quit()

if __name__ == "__main__":
    main()
