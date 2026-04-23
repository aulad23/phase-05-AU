"""
SKILL: VPN / Proxy / Anti-Bot Handling
কিছু vendor সরাসরি block করে — এই skill দিয়ে handle করা হয়।

DIFFICULTY LEVELS:
  EASY   → Requests+BS4, no JS needed
  MEDIUM → Selenium + stealth headers
  HIGH   → undetected_chromedriver OR Chrome Debug Port + fingerprint spoof
  VPN    → VPN লাগবে (manually connect করুন), তারপর PROXY= set করুন
"""

import os, time, socket, subprocess, sys, random, shlex

# ── KNOWN DIFFICULT VENDORS ───────────────────────────────────────────────────
DIFFICULT_VENDORS = {
    "Surya":          {"difficulty": "HIGH", "method": "chrome_debug_port + uc"},
    "Loloi":          {"difficulty": "HIGH", "method": "selenium_stealth"},
    "Visual Comfort": {"difficulty": "MEDIUM", "method": "selenium_stealth"},
    "Holly Hunt":     {"difficulty": "MEDIUM", "method": "selenium_stealth"},
    "Janus et Cie":   {"difficulty": "MEDIUM", "method": "selenium_stealth"},
    "Liaigre":        {"difficulty": "HIGH",   "method": "uc"},
}

BLOCK_SIGNALS = [
    "verify you are human", "are you human", "access denied",
    "unusual traffic", "403 forbidden", "captcha", "cloudflare",
    "bot detection", "please enable javascript", "checking your browser",
    "ddos protection", "ray id",
]


# ── BLOCK DETECTION ───────────────────────────────────────────────────────────
def is_blocked(page_source: str) -> bool:
    lower = page_source.lower()
    return any(signal in lower for signal in BLOCK_SIGNALS)


def check_and_warn(vendor_name: str) -> dict:
    """Print warning if vendor is known difficult. Return info dict."""
    info = DIFFICULT_VENDORS.get(vendor_name, {})
    if info:
        print(f"\n  ⚠️  [{vendor_name}] Difficulty: {info['difficulty']}")
        print(f"      Method: {info['method']}")
        if "vpn" in info.get("method", "").lower():
            print(f"      ACTION: VPN manually connect করুন তারপর Enter চাপুন")
            input("      VPN ready? [Enter] ")
    return info


# ── CHROME DEBUG PORT (Surya-style) ───────────────────────────────────────────
DEBUG_HOST = "127.0.0.1"
DEBUG_PORT = 9222
USER_DATA_DIR = r"C:\ChromeProfile\SuryaAgent"

def port_open(host: str, port: int, timeout: float = 1.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False


def launch_debug_chrome(user_data_dir: str = USER_DATA_DIR,
                         port: int = DEBUG_PORT) -> None:
    """Launch Chrome with remote-debugging-port (connects to existing session)."""
    if port_open(DEBUG_HOST, port):
        print(f"  Chrome debug port {port} already open.")
        return

    os.makedirs(user_data_dir, exist_ok=True)
    exe = None
    for p in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]:
        if os.path.exists(p):
            exe = p
            break
    if not exe:
        exe = "chrome.exe"

    cmd = (f'"{exe}" --remote-debugging-port={port} '
           f'--user-data-dir="{user_data_dir}" '
           f'--no-first-run --no-default-browser-check')

    if sys.platform.startswith("win"):
        subprocess.Popen(cmd, creationflags=0x00000008)
    else:
        subprocess.Popen(shlex.split(cmd))

    for _ in range(20):
        if port_open(DEBUG_HOST, port):
            return
        time.sleep(0.5)
    raise RuntimeError(f"Chrome debug port {port} not reachable after 10s")


def get_debug_driver(port: int = DEBUG_PORT):
    """Connect Selenium to an already-running Chrome debug session."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    import chromedriver_autoinstaller as cda
    cda.install()
    opts = Options()
    opts.add_experimental_option("debuggerAddress", f"{DEBUG_HOST}:{port}")
    opts.add_argument("--start-maximized")
    return webdriver.Chrome(options=opts)


# ── UNDETECTED CHROMEDRIVER (UC) ──────────────────────────────────────────────
SPOOF_UA       = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
SPOOF_PLATFORM = "Win32"
SPOOF_VENDOR   = "Google Inc."
SPOOF_WEBGL    = "Intel Iris OpenGL Engine"

def get_uc_driver(headless: bool = False, proxy: str = None,
                  user_data_dir: str = None):
    """
    Undetected ChromeDriver — bypasses Cloudflare + bot detection.
    proxy format: "http://user:pass@host:port" or "http://host:port"
    """
    import undetected_chromedriver as uc

    opts = uc.ChromeOptions()
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    if headless:
        opts.add_argument("--headless=new")
    if proxy:
        opts.add_argument(f"--proxy-server={proxy}")
    if user_data_dir and os.path.isdir(user_data_dir):
        opts.add_argument(f"--user-data-dir={user_data_dir}")

    driver = uc.Chrome(options=opts)

    # Browser fingerprint spoof
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": f"""
        Object.defineProperty(navigator, 'webdriver', {{get: () => undefined}});
        Object.defineProperty(navigator, 'platform', {{get: () => '{SPOOF_PLATFORM}'}});
        Object.defineProperty(navigator, 'vendor', {{get: () => '{SPOOF_VENDOR}'}});
        Object.defineProperty(navigator, 'languages', {{get: () => ['en-US','en']}});
    """})
    return driver


# ── STEALTH REQUESTS SESSION ──────────────────────────────────────────────────
def get_stealth_session(proxy: str = None):
    """requests.Session with stealth headers + optional proxy."""
    import requests
    s = requests.Session()
    s.headers.update({
        "User-Agent": SPOOF_UA,
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    })
    if proxy:
        s.proxies = {"http": proxy, "https": proxy}
    return s


# ── AUTO-SELECT DRIVER ────────────────────────────────────────────────────────
def auto_driver(vendor_name: str, proxy: str = None):
    """
    Based on vendor difficulty, auto-select the right driver.
    Returns (driver, driver_type) where driver_type in ['uc','debug','selenium','requests']
    """
    info = DIFFICULT_VENDORS.get(vendor_name, {})
    method = info.get("method", "selenium")

    if "chrome_debug_port" in method:
        launch_debug_chrome()
        return get_debug_driver(), "debug"
    elif "uc" in method:
        return get_uc_driver(proxy=proxy), "uc"
    elif "selenium" in method:
        from Agent.Skills._03_web_scraping import get_driver
        return get_driver(), "selenium"
    else:
        return get_stealth_session(proxy=proxy), "requests"


# ── RETRY ON BLOCK ────────────────────────────────────────────────────────────
def retry_on_block(driver, url: str, max_retries: int = 3,
                   wait: float = 5.0) -> bool:
    """Navigate to URL, retry if blocked."""
    for attempt in range(1, max_retries + 1):
        driver.get(url)
        time.sleep(wait + random.uniform(0, 2))
        if not is_blocked(driver.page_source):
            return True
        print(f"  Blocked (attempt {attempt}/{max_retries}). Waiting {wait*2:.0f}s...")
        time.sleep(wait * 2)
    print("  Still blocked after retries. Try VPN or Chrome Debug mode.")
    return False
