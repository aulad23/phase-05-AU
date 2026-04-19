import time
import pandas as pd
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

# ----------------- CONFIG -----------------
CHROMEDRIVER_PATH = r"C:\chromedriver.exe"  # change this path
LIST_URLS = [
    "https://www.palecek.com/itembrowser.aspx?action=attributes&itemtype=furniture&custom%20department=furniture&custom%20category=stools&viewall=true",
]
OUTPUT_XLSX = "palecek_bar_stools.xlsx"

WAIT_TIMEOUT = 25
ENABLE_DETAIL_BACKFILL = True
DETAIL_BACKFILL_TIMEOUT = 12
# ------------------------------------------


def connect_driver() -> webdriver.Chrome:
    opts = Options()
    # opts.add_argument("--headless=new")  # uncomment if headless
    opts.add_argument("--disable-gpu")
    opts.add_argument("--start-maximized")
    service = Service(CHROMEDRIVER_PATH)
    return webdriver.Chrome(service=service, options=opts)


def normalize_image_url(src: str) -> str:
    if not src:
        return ""
    src = src.strip()
    if src.startswith("//"):
        return "https:" + src
    if src.startswith("/"):
        return urljoin("https://www.palecek.com", src)
    return src


def safe_text(el) -> str:
    if not el:
        return ""
    return (el.get_attribute("textContent") or "").strip()


def extract_item_strong(card, driver, max_retries=3):
    """Extract product data robustly."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card)
        time.sleep(0.15)
    except Exception:
        pass

    product_url, img_url, name_text, sku_text = "", "", "", ""

    try:
        a = card.find_element(By.CSS_SELECTOR, "a[href*='iteminformation.aspx']")
        product_url = urljoin("https://www.palecek.com", a.get_attribute("href"))
    except Exception:
        pass

    try:
        img = card.find_element(By.CSS_SELECTOR, "img.ProductThumbnailImg")
        img_url = img.get_attribute("data-src") or img.get_attribute("src") or ""
        img_url = normalize_image_url(img_url)
    except Exception:
        pass

    for _ in range(max_retries):
        try:
            name_el = card.find_element(
                By.CSS_SELECTOR,
                "div.ProductThumbnailDetails p.ProductThumbnailParagraphDescription a"
            )
            sku_el = card.find_element(
                By.CSS_SELECTOR,
                "div.ProductThumbnailDetails p.ProductThumbnailParagraphSkuName a"
            )
            name_text = safe_text(name_el)
            sku_text = safe_text(sku_el)

            if not sku_text:
                try:
                    h3 = card.find_element(By.CSS_SELECTOR, "div.ProductThumbnailDetails h3")
                    sku_text = safe_text(h3)
                except Exception:
                    pass
            if name_text and sku_text:
                break
            time.sleep(0.25)
        except StaleElementReferenceException:
            time.sleep(0.2)
        except Exception:
            time.sleep(0.2)

    return {
        "Product URL": product_url,
        "Image URL": img_url,
        "Product Name": name_text,
        "SKU": sku_text,
    }


# ---------- ONE-TIME SCROLL VERSION ----------
def scroll_and_collect(driver, base_url):
    """
    Loads the given 'viewall=true' URL once, scrolls once to ensure all images load,
    and collects all product cards (no pagination or repeated scrolling).
    """
    all_data = []
    seen_keys = set()

    driver.get(base_url)
    time.sleep(4)
    print("⬇️ Loading all products (one scroll)...")

    # One big scroll to ensure all lazy elements load
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(4)

    cards = driver.find_elements(
        By.CSS_SELECTOR,
        "div.ItemBrowserThumbnailContainer section.ProductThumbnailSection div.ProductThumbnail"
    )

    print(f"🔍 Found {len(cards)} product cards — extracting...")

    for card in cards:
        data = extract_item_strong(card, driver)
        key = data["Product URL"] or (data["Product Name"], data["SKU"])
        if key and key not in seen_keys:
            seen_keys.add(key)
            all_data.append(data)

    print(f"✅ Finished collecting {len(all_data)} unique products.")
    return all_data
# ----------------------------------------------------------


def fill_from_detail(driver, url, timeout=DETAIL_BACKFILL_TIMEOUT):
    """Optional: open detail page to fill missing name/SKU."""
    if not url:
        return ("", "")

    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])
    name = ""
    sku = ""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )
        for sel in ["h1", "h2", ".ItemName", ".ProductTitle", "#lblItemTitle"]:
            try:
                name = (driver.find_element(By.CSS_SELECTOR, sel)
                        .get_attribute("textContent") or "").strip()
                if name:
                    break
            except Exception:
                pass

        for sel in [".sku", ".ItemNumber", "#lblItemNumber", "[data-sku]", ".item-number"]:
            try:
                sku = (driver.find_element(By.CSS_SELECTOR, sel)
                       .get_attribute("textContent") or "").strip()
                if sku:
                    break
            except Exception:
                pass
    finally:
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    return (name, sku)


def harvest_from_list_page(driver, url):
    collected = scroll_and_collect(driver, url)

    if ENABLE_DETAIL_BACKFILL:
        for r in collected:
            if (not r["Product Name"] or not r["SKU"]) and r["Product URL"]:
                name, sku = fill_from_detail(driver, r["Product URL"])
                if name and not r["Product Name"]:
                    r["Product Name"] = name
                if sku and not r["SKU"]:
                    r["SKU"] = sku

    return collected


def main():
    driver = connect_driver()
    all_rows = []
    global_seen = set()

    try:
        for idx, list_url in enumerate(LIST_URLS, start=1):
            print(f"[{idx}/{len(LIST_URLS)}] Harvesting: {list_url}")
            rows = harvest_from_list_page(driver, list_url)

            for r in rows:
                key = r.get("Product URL") or (r.get("Product Name"), r.get("SKU"))
                if key and key not in global_seen:
                    global_seen.add(key)
                    all_rows.append(r)

            print(f"  Collected so far: {len(all_rows)}")

        cols = ["Product URL", "Image URL", "Product Name", "SKU"]
        df = pd.DataFrame(all_rows, columns=cols)
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"✅ Saved {len(df)} items to {OUTPUT_XLSX}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
