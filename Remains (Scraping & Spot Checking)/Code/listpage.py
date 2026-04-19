import time
from urllib.parse import urljoin, urlparse, parse_qs, urlunparse

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# -----------------------------
# CONFIG
# -----------------------------
BASE_URL = "https://remains.com/collections/sconces"
TARGET_COUNT = float("inf")

# Be generous on timeouts; flaky networks need headroom
PAGE_LOAD_TIMEOUT = 180
WAIT_TIMEOUT = 40
SCROLL_PAUSE = 0.5
CARD_SCROLL_PAUSE = 0.18

# Use your standard path
CHROMEDRIVER_PATH = r"C:/chromedriver.exe"
OUTPUT_XLSX = "remains_sconces.xlsx"

# -----------------------------
# HELPERS
# -----------------------------
def build_page_url(base, page_num: int) -> str:
    parsed = urlparse(base)
    q = parse_qs(parsed.query)
    q["page"] = [str(page_num)]
    new_query = "&".join([f"{k}={v[0] if isinstance(v, list) else v}" for k, v in q.items()])
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path, parsed.params, new_query, parsed.fragment))

def robust_find(driver, by, selector, timeout=WAIT_TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, selector)))

def wait_for_results_grid(driver, timeout=WAIT_TIMEOUT):
    # Wait for document ready first, then the results container
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
    )
    return robust_find(driver, By.CSS_SELECTOR, "div.usf-results", timeout=timeout)

def scroll_to_bottom_incremental(driver, stop_after_idle=3, step_px=1200):
    idle_passes = 0
    last_height = driver.execute_script("return document.body.scrollHeight") or 0
    while True:
        driver.execute_script(f"window.scrollBy(0, {step_px});")
        time.sleep(SCROLL_PAUSE)
        new_height = driver.execute_script("return document.body.scrollHeight") or last_height
        if new_height <= last_height:
            idle_passes += 1
        else:
            idle_passes = 0
            last_height = new_height
        if idle_passes >= stop_after_idle:
            break

def ensure_image_src(driver, img_el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", img_el)
    time.sleep(CARD_SCROLL_PAUSE)
    deadline = time.time() + 4.5
    while time.time() < deadline:
        try:
            src = (img_el.get_attribute("src") or "").strip()
            if src and not src.startswith("data:"):
                return src
        except StaleElementReferenceException:
            pass
        time.sleep(0.12)
    # fall back to data-src / srcset if present
    try:
        srcset = (img_el.get_attribute("srcset") or "").strip()
        if srcset:
            return srcset.split()[-2] if len(srcset.split()) >= 2 else srcset.split()[0]
    except Exception:
        pass
    return ""

def parse_cards_on_page(driver):
    """
    Extract rows with keys:
      - Product URL
      - Image URL
      - Product Name
    """
    container = wait_for_results_grid(driver)
    cards = container.find_elements(By.CSS_SELECTOR, "div.grid__item.grid-product")
    rows = []

    for card in cards:
        # Link
        try:
            link_el = card.find_element(By.CSS_SELECTOR, "a.grid-product__link")
            href = urljoin(driver.current_url, (link_el.get_attribute("href") or "").strip())
        except Exception:
            href = ""

        # Name
        try:
            name_el = card.find_element(By.CSS_SELECTOR, "div.grid-product__title.grid-product__title--heading")
            name_text = (name_el.text or "").strip()
        except Exception:
            name_text = ""

        # Image
        try:
            img_el = card.find_element(By.CSS_SELECTOR, "img.grid-product__image")
            img_src = ensure_image_src(driver, img_el) or ""
        except Exception:
            img_src = ""

        if href:
            rows.append({
                "Product URL": href,
                "Image URL": img_src,
                "Product Name": name_text
            })
    return rows

def safe_get(driver, url, attempts=3):
    """
    Load a URL with retries and stop heavy subresources if needed.
    Prevents the session from dying early due to slow assets.
    """
    for i in range(1, attempts + 1):
        try:
            driver.get(url)
            # Wait for at least a partially-ready page before returning
            WebDriverWait(driver, WAIT_TIMEOUT).until(
                lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
            )
            return
        except TimeoutException:
            try:
                driver.execute_script("window.stop();")
            except Exception:
                pass
            if i == attempts:
                raise
        except Exception:
            if i == attempts:
                raise
        time.sleep(0.8 * i)  # simple backoff

# -----------------------------
# MAIN
# -----------------------------
def main():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")  # harmless on Windows
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # Keep the browser open even if Python exits (helps debug early-quit issues)
    options.add_experimental_option("detach", True)

    # Prefer default loading (more stable than 'eager' on JS-heavy pages)
    # options.page_load_strategy = "eager"  # if you want speed over stability, re-enable
    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    driver.implicitly_wait(2)  # small implicit to smooth over tiny gaps

    all_rows, seen_links = [], set()
    page = 1

    try:
        while len(all_rows) < TARGET_COUNT:
            page_url = build_page_url(BASE_URL, page)
            print(f"Navigating to {page_url}")
            safe_get(driver, page_url, attempts=3)

            # Ensure the grid exists (waits internally)
            wait_for_results_grid(driver, timeout=WAIT_TIMEOUT)

            # Trigger lazyload for all cards/images
            scroll_to_bottom_incremental(driver, stop_after_idle=3, step_px=1200)

            # Parse
            page_rows = parse_cards_on_page(driver)

            # Dedupe by Product URL
            added = 0
            for r in page_rows:
                link = r.get("Product URL", "")
                if link and link not in seen_links:
                    seen_links.add(link)
                    all_rows.append(r)
                    added += 1
                    if len(all_rows) >= TARGET_COUNT:
                        break

            print(f"[Page {page}] Found {len(page_rows)} cards, added {added}. Total: {len(all_rows)}")

            # If nothing new was found, try one more gentle refresh before giving up
            if added == 0:
                print("No new items; refreshing page once to defeat lazyload/race.")
                safe_get(driver, page_url, attempts=1)
                scroll_to_bottom_incremental(driver, stop_after_idle=3, step_px=1200)
                page_rows = parse_cards_on_page(driver)
                for r in page_rows:
                    link = r.get("Product URL", "")
                    if link and link not in seen_links:
                        seen_links.add(link)
                        all_rows.append(r)
                        added += 1
                        if len(all_rows) >= TARGET_COUNT:
                            break
                print(f"[Page {page} refresh] Added {added}. Total: {len(all_rows)}")
                if added == 0:
                    print("Still no new items. Stopping.")
                    break

            page += 1
            time.sleep(1.0)

    finally:
        # With detach=True, the browser will stay open even if we quit the driver.
        try:
            driver.quit()
        except Exception:
            pass

    # Save
    df = pd.DataFrame(all_rows, columns=["Product URL", "Image URL", "Product Name"])
    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"Saved {len(df)} rows to {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
