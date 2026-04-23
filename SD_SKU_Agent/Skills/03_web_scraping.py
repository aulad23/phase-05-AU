"""
SKILL: Web Scraping
Selenium + Requests patterns. Agent এটা ব্যবহার করে scraper তৈরি করে।
"""

import time
import random
import requests
from pathlib import Path

# ── REQUESTS HEADERS ──────────────────────────────────────────────────────────
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
}


def get_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(DEFAULT_HEADERS)
    return s


def safe_get(session: requests.Session, url: str, retries: int = 3) -> requests.Response | None:
    for i in range(retries):
        try:
            r = session.get(url, timeout=20)
            r.raise_for_status()
            return r
        except Exception as e:
            print(f"  GET error ({i+1}/{retries}): {e}")
            time.sleep(2 * (i + 1))
    return None


# ── SELENIUM SETUP ────────────────────────────────────────────────────────────
def get_driver(headless: bool = True):
    """Chrome WebDriver with stealth settings."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(f"user-agent={DEFAULT_HEADERS['User-Agent']}")
    opts.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(options=opts)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"},
    )
    return driver


# ── SCROLL TO BOTTOM ──────────────────────────────────────────────────────────
def scroll_to_bottom(driver, pause: float = 1.5, max_rounds: int = 20):
    """Infinite scroll — stop when page height stops growing."""
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(max_rounds):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


# ── WAIT FOR ELEMENT ──────────────────────────────────────────────────────────
def wait_for(driver, selector: str, by="css", timeout: int = 10):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By

    by_map = {"css": By.CSS_SELECTOR, "xpath": By.XPATH, "id": By.ID}
    by_const = by_map.get(by, By.CSS_SELECTOR)
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by_const, selector))
    )


# ── PAGINATION: QUERY PARAM ───────────────────────────────────────────────────
def paginate_query(base_url: str, param: str = "page", start: int = 1):
    """Generator: yields page URLs until empty."""
    from bs4 import BeautifulSoup
    session = get_session()
    page = start
    while True:
        url = f"{base_url}?{param}={page}"
        r = safe_get(session, url)
        if not r:
            break
        soup = BeautifulSoup(r.text, "html.parser")
        yield url, soup, page
        page += 1
        time.sleep(random.uniform(1, 2))


# ── IMAGE FALLBACK ────────────────────────────────────────────────────────────
def get_image_url(img_tag) -> str:
    """Try src → data-src → data-lazy-src in order."""
    if img_tag is None:
        return ""
    for attr in ("src", "data-src", "data-lazy-src", "data-original"):
        val = img_tag.get(attr, "")
        if val and val.startswith("http"):
            return val
    return ""
