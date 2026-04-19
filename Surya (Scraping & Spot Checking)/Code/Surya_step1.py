# Requirements:
#   pip install selenium pandas chromedriver-autoinstaller openpyxl

import os, sys, time, re, socket, subprocess, shlex, random
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    SessionNotCreatedException, WebDriverException
)

import chromedriver_autoinstaller as cda

# ===================== USER SETTINGS =====================

# >>> ekhane shudhu category URL change korba <<<
START_URL = "https://www.surya.com/Catalog/Rugs/All"

USE_DEBUGGER = True
DEBUG_HOST = "127.0.0.1"
DEBUG_PORT = 9222
USER_DATA_DIR = r"C:\ChromeProfile\Surya"

# File path = script-er location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = SCRIPT_DIR
OUTPUT_PATH = os.path.join(DOWNLOAD_DIR, "surya_Chandeliers.xlsx")

# Scroll settings – boro category jonno tuned
INITIAL_WAIT = 25                  # first page render allowance
PAGE_WAIT = 30
SCROLL_STEP_PX = 2000              # Increased scroll step for quicker page load
BOTTOM_PAUSE = (0.5, 1.0)          # Reduced pause time to make scrolling faster
STABLE_ROUNDS_LIMIT = 45           # koto bar same count hole stop
GLOBAL_MAX_TIME = 7200             # max 2 ghonta scroll (4k–10k jonno safe)
LOAD_MORE_CLICK_PAUSE = (2.0, 3.5)
POST_LOAD_SETTLE = (0.5, 1.0)      # Slightly reduced settling time after each scroll

MAX_SCROLL_ERRORS = 6              # jodi bar bar error ase, scroll loop theke ber hoy

# joto new row add holo, tar por koto por por console e summary dekhabo
LOG_EVERY_NEW = 200

# ===================== SELECTORS =====================

GRID_HINTS = [
    "[data-test-selector='productGrid']",
    "div[data-testid='product-grid']",
    "div[data-test-selector='productGridItem']",
    "ul[role='list']",
    "div[class*='grid']",
]
SEL_PRODUCT_CARD = "div[data-test-selector='productGridItem']"
SEL_PRODUCT_LINKS = "a[data-test-selector='productImage'], a[data-test-selector='productDescriptionLink']"
SEL_SKU_CANDIDATES = [
    "span.TypographyStyle--11lquxl.gugPDz",
    "[data-test-selector='productStyleId']",
    "span[class*='sku']",
    "div[class*='sku']",
    "p[class*='sku']",
    "span[class*='style']",
    "div[class*='style']",
]
SEL_LOAD_MORE = "button[data-test-selector='loadMore'], button[data-testid='load-more'], button[aria-label='Load more']"

CONSENT_XPATHS = [
    "//button[contains(., 'Accept')]", "//button[contains(., 'I Accept')]",
    "//button[contains(., 'I agree')]", "//button[contains(., 'Agree')]",
    "//button[contains(., 'Got it')]", "//button[contains(., 'Allow')]",
    "//a[contains(., 'Accept')]", "//span[contains(., 'Accept')]",
    "//button[@id='onetrust-accept-btn-handler']",
]

BLOCK_TEXTS = ["verify you are human", "are you human", "access denied", "unusual traffic"]


# ===================== UTILS =====================

def port_open(host: str, port: int, timeout: float = 1.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False

def pause(a, b=None):
    time.sleep(random.uniform(a, b) if b is not None else a)

def parse_first_url_from_srcset(srcset: str) -> str:
    if not srcset:
        return ""
    try:
        return srcset.split(",")[0].strip().split(" ")[0].strip()
    except Exception:
        return ""

def extract_bg_image(style_val: str) -> str:
    if not style_val:
        return ""
    m = re.search(r'background-image\s*:\s*url\(["\']?(.*?)["\']?\)', style_val, re.I)
    return m.group(1) if m else ""

def save_debug(driver, prefix="surya_debug"):
    try:
        ss = os.path.join(DOWNLOAD_DIR, f"{prefix}.png")
        driver.save_screenshot(ss)
        print(f"[debug] Screenshot: {ss}")
    except Exception:
        pass
    try:
        htmlp = os.path.join(DOWNLOAD_DIR, f"{prefix}.html")
        with open(htmlp, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"[debug] HTML: {htmlp}")
    except Exception:
        pass

def still_blocked(driver) -> bool:
    try:
        page = driver.page_source.lower()
        return any(t in page for t in BLOCK_TEXTS)
    except Exception:
        return False

# --- SKU normalizer ---
def clean_sku(raw: str) -> str:
    """Return a clean SKU like 'ARLD-005' without trailing stray 'p'/'-p'/'_p'."""
    if not raw:
        return ""
    t = raw.strip()
    t = re.sub(r'\s+', '', t)
    m = re.search(r'([A-Za-z0-9]+(?:[-_][A-Za-z0-9]+)+)', t)
    if m:
        t = m.group(1)
    else:
        t = re.sub(r'[^A-Za-z0-9\-_]', '', t)
    t = re.sub(r'[-_]?p$', '', t, flags=re.I)
    return t

# ===================== CHROME HELPERS =====================

def ensure_debug_chrome_running():
    if port_open(DEBUG_HOST, DEBUG_PORT, timeout=0.5):
        return
    os.makedirs(USER_DATA_DIR, exist_ok=True)
    exe_path = None
    for p in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]:
        if os.path.exists(p):
            exe_path = p
            break
    if exe_path is None:
        exe_path = "chrome.exe"
    cmd = f'"{exe_path}" --remote-debugging-port={DEBUG_PORT} --user-data-dir="{USER_DATA_DIR}" --no-first-run --no-default-browser-check'
    try:
        if sys.platform.startswith("win"):
            DETACHED_PROCESS = 0x00000008
            subprocess.Popen(cmd, creationflags=DETACHED_PROCESS)
        else:
            subprocess.Popen(shlex.split(cmd))
    except Exception as e:
        raise RuntimeError(f"Failed to launch Chrome with debug: {e}\nTried: {cmd}")
    for _ in range(20):
        if port_open(DEBUG_HOST, DEBUG_PORT, timeout=0.5):
            return
        time.sleep(0.5)
    raise RuntimeError("Chrome debug port not reachable")

def setup_driver():
    cda.install()
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    if USE_DEBUGGER:
        ensure_debug_chrome_running()
        opts.add_experimental_option("debuggerAddress", f"{DEBUG_HOST}:{DEBUG_PORT}")
    try:
        return webdriver.Chrome(options=opts)
    except SessionNotCreatedException as e:
        raise RuntimeError(f"❌ Session not created: {e}")
    except WebDriverException as e:
        raise RuntimeError(f"WebDriver init failed: {e}")

# ===================== PAGE WAIT/DISMISS =====================

def dismiss_overlays(driver):
    for xp in CONSENT_XPATHS:
        try:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    pause(0.2, 0.5)
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    pause(0.4, 0.7)
                    return True
        except Exception:
            continue
    return False

def wait_until_products_render(driver, total_timeout=INITIAL_WAIT):
    end = time.time() + total_timeout
    while time.time() < end:
        dismiss_overlays(driver)

        try:
            for css in GRID_HINTS:
                if driver.find_elements(By.CSS_SELECTOR, css):
                    cards = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_CARD)
                    anchors = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_LINKS)
                    count = max(len(cards), len(anchors))
                    if count > 0:
                        return True
        except Exception:
            pass

        try:
            driver.execute_script("window.scrollBy(0, 800);")
            pause(0.5, 1.0)
            driver.execute_script("window.scrollBy(0, -500);")
            pause(0.3, 0.6)
        except Exception:
            pass

        try:
            cards = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_CARD)
            anchors = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_LINKS)
            count = max(len(cards), len(anchors))
            if count > 0:
                return True
        except Exception:
            pass

        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            pause(0.6, 1.2)
            driver.execute_script("window.scrollTo(0, 0);")
        except Exception:
            pass

        if still_blocked(driver):
            try:
                driver.execute_script("window.stop();")
            except Exception:
                pass
            pause(0.5, 1.0)
            driver.get(START_URL)
            pause(1.0, 1.6)

    return False

# ===================== IMAGE & SKU HELPERS =====================

def get_image_url_from_card(card):
    try:
        img = card.find_element(By.CSS_SELECTOR, "img")
        url = img.get_attribute("src") or img.get_attribute("data-src") or ""
        if not url:
            url = parse_first_url_from_srcset(img.get_attribute("srcset") or "")
        if url:
            return url
    except Exception:
        pass
    try:
        src = card.find_element(By.CSS_SELECTOR, "picture source")
        url = parse_first_url_from_srcset(src.get_attribute("srcset") or "")
        if url:
            return url
    except Exception:
        pass
    try:
        styled = card.find_elements(By.CSS_SELECTOR, "[style*='background-image']")
        for s in styled:
            url = extract_bg_image(s.get_attribute("style") or "")
            if url:
                return url
    except Exception:
        pass
    return ""

def get_sku_from_card(card, product_url):
    for css in SEL_SKU_CANDIDATES:
        try:
            el = card.find_element(By.CSS_SELECTOR, css)
            raw = (el.text or "").strip()
            if raw:
                sku = clean_sku(raw)
                if sku:
                    return sku
        except Exception:
            pass

    try:
        imgs = card.find_elements(By.CSS_SELECTOR, "img")
        if imgs:
            raw = (imgs[0].get_attribute("alt") or "").strip()
            sku = clean_sku(raw)
            if sku:
                return sku
    except Exception:
        pass

    if product_url:
        slug = product_url.rstrip("/").split("/")[-1]
        slug = clean_sku(slug)
        if slug:
            return slug
    return "MISSING"

# ===================== INCREMENTAL COLLECTOR =====================

def collect_new_products_from_dom(driver, seen_keys, collected_rows):
    try:
        cards = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_CARD)
        anchors = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_LINKS)
    except Exception:
        cards, anchors = [], []

    using_cards = len(cards) >= len(anchors)
    base = "https://www.surya.com"
    nodes = cards if using_cards else anchors

    new_count = 0

    for node in nodes:
        url = ""
        if using_cards:
            try:
                a = node.find_element(By.CSS_SELECTOR, SEL_PRODUCT_LINKS)
                url = a.get_attribute("href") or ""
            except Exception:
                try:
                    a2 = node.find_element(By.CSS_SELECTOR, "a")
                    url = a2.get_attribute("href") or ""
                except Exception:
                    url = ""
        else:
            try:
                url = node.get_attribute("href") or ""
            except Exception:
                url = ""

        if url.startswith("/"):
            url = base.rstrip("/") + url
        url = url.strip()

        if using_cards:
            img = get_image_url_from_card(node)
            sku = get_sku_from_card(node, url)
        else:
            img = ""
            try:
                img_el = node.find_element(By.XPATH, ".//img")
                img = img_el.get_attribute("src") or img_el.get_attribute("data-src") or ""
                if not img:
                    img = parse_first_url_from_srcset(img_el.get_attribute("srcset") or "")
            except Exception:
                pass

            raw = ""
            try:
                raw = (node.get_attribute("aria-label") or "").strip()
            except Exception:
                raw = ""
            if not raw:
                try:
                    raw = node.text.strip()
                except Exception:
                    raw = ""

            sku = clean_sku(raw) or clean_sku(url.rstrip("/").split("/")[-1]) or "MISSING"

        sku = sku.strip()

        key = (sku, url)
        if not url:
            continue
        if key in seen_keys:
            continue

        seen_keys.add(key)
        collected_rows.append({
            "Product URL": url,
            "Image URL": img,
            "SKU": sku
        })
        new_count += 1

    return new_count

# ===================== SCROLL & LOAD MORE (WITH COLLECT) =====================

def try_click_load_more(driver) -> bool:
    return False  # Disable Load More button clicks for faster data loading

def smart_slow_scroll_to_bottom(driver, seen_keys, collected_rows):
    start_time = time.time()
    stable_rounds = 0
    last_count = 0
    last_height = 0
    error_rounds = 0
    last_logged_total = 0

    base_new = collect_new_products_from_dom(driver, seen_keys, collected_rows)
    if base_new:
        print(f"📝 Collected {base_new} new products (total saved: {len(collected_rows)})")

    while True:
        if (time.time() - start_time) > GLOBAL_MAX_TIME:
            print("⏳ Reached global max time cap; stopping scroll.")
            break

        try:
            cards = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_CARD)
            anchors = driver.find_elements(By.CSS_SELECTOR, SEL_PRODUCT_LINKS)
        except Exception:
            cards, anchors = [], []

        cur_dom_count = max(len(cards), len(anchors))

        try:
            height = driver.execute_script("return document.body.scrollHeight || 0;")
        except Exception:
            height = 0

        if cur_dom_count > last_count:
            print(f"✅ Loaded {cur_dom_count} products so far (DOM nodes)…")
            last_count = cur_dom_count
            stable_rounds = 0
        else:
            if abs(height - last_height) < 50:
                stable_rounds += 1
            else:
                stable_rounds = 0

        last_height = height

        try:
            driver.execute_script(f"window.scrollBy(0, {SCROLL_STEP_PX});")
        except Exception as e:
            error_rounds += 1
            print(f"⚠️ Scroll error (round {error_rounds}): {e}")
            if error_rounds >= MAX_SCROLL_ERRORS:
                print("🚧 Too many scroll errors, stopping scroll loop.")
                break
        pause(*BOTTOM_PAUSE)

        if stable_rounds % 3 == 0:
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            except Exception:
                error_rounds += 1
            pause(0.8, 1.4)

        if try_click_load_more(driver):
            pause(*POST_LOAD_SETTLE)

        new_count = collect_new_products_from_dom(driver, seen_keys, collected_rows)
        if new_count:
            total_saved = len(collected_rows)
            if total_saved - last_logged_total >= LOG_EVERY_NEW:
                print(f"📝 Collected +{new_count} new (total saved: {total_saved})")
                last_logged_total = total_saved

        if stable_rounds >= STABLE_ROUNDS_LIMIT:
            print("ℹ️ No new products and page height stable. Finishing scroll.")
            break

        if error_rounds >= MAX_SCROLL_ERRORS:
            print("🚧 Too many errors in scroll loop, breaking out.")
            break

# ===================== MAIN =====================

def main():
    driver = setup_driver()
    seen_keys = set()
    collected_rows = []

    try:
        driver.get(START_URL)

        print("⏳ Waiting for products to render…")
        if not wait_until_products_render(driver, total_timeout=INITIAL_WAIT):
            print("❌ Product grid not detected after robust wait.")
            save_debug(driver, prefix="surya_grid_fail")
            return

        print("✅ Grid/anchors detected. Scrolling slowly to load ALL products…")
        try:
            smart_slow_scroll_to_bottom(driver, seen_keys, collected_rows)
        except Exception as e:
            print(f"⚠️ Scroll loop crashed unexpectedly: {e}")
            print("➡️ Using whatever products are currently loaded (already collected).")

        extra_new = collect_new_products_from_dom(driver, seen_keys, collected_rows)
        if extra_new:
            print(f"📝 Final extra collect: +{extra_new} (total saved: {len(collected_rows)})")

        if not collected_rows:
            print("⚠️ No rows collected. Nothing to save.")
            return

        df = pd.DataFrame(collected_rows, columns=["Product URL", "Image URL", "SKU"])
        df.to_excel(OUTPUT_PATH, index=False)
        print(f"✅ Saved {len(df)} unique products -> {OUTPUT_PATH}")

        if sys.platform.startswith("win"):
            try:
                os.startfile(OUTPUT_PATH)
            except Exception:
                pass
        elif sys.platform.startswith("darwin"):
            os.system(f'open "{OUTPUT_PATH}"')
        else:
            try:
                os.system(f'xdg-open "{OUTPUT_PATH}"')
            except Exception:
                pass

    finally:
        print("Done. Closing in 4s…")
        time.sleep(4)
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
