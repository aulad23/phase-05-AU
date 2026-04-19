import os, sys, re, time, random, socket, subprocess, shlex, json, collections, glob
from urllib.parse import urljoin

import pandas as pd
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service  # Import Service
from webdriver_manager.chrome import ChromeDriverManager  # WebDriverManager for automatic ChromeDriver installation

# ---------- guard against local name shadowing ----------
def _check_local_shadow():
    here = os.path.abspath(os.getcwd())
    for name in ["selenium.py", "webdriver.py", "selenium", "webdriver"]:
        p = os.path.join(here, name)
        if os.path.isfile(p) or os.path.isdir(p):
            raise RuntimeError(f"Local '{name}' found at {p}. Rename/delete; it shadows Selenium.")
_check_local_shadow()

# ================= USER SETTINGS (FIXED FOR UC) =================
HEADLESS = False
USE_DEBUGGER = False
USE_STEALTH = True
USE_UC_FALLBACK = True

DEBUG_HOST = "127.0.0.1"
DEBUG_PORT = 9222
USER_DATA_DIR = r"C:\ChromeProfile\Surya"

PROXY = None

# WebDriverManager will automatically manage ChromeDriver version
CHROMEDRIVER_PATH = ChromeDriverManager().install()  # Automatically installs the correct version of ChromeDriver

SPOOFED_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
SPOOFED_PLATFORM = "Win32"
SPOOFED_LANGUAGES = ["en-US", "en"]
SPOOFED_VENDOR = "Google Inc."
SPOOFED_WEBGL_VENDOR = "Intel Inc."
SPOOFED_WEBGL_RENDERER = "Intel Iris OpenGL Engine"

# ================= FUNCTION: Setup Driver =================
def setup_driver():
    # 🔹 First try Undetected Chrome (UC)
    if USE_UC_FALLBACK:
        options = uc.ChromeOptions()
        args = [
            "--disable-gpu", "--disable-dev-shm-usage", "--no-sandbox",
            "--start-maximized", "--disable-blink-features=AutomationControlled",
        ]
        if HEADLESS:
            args.append("--headless=new")
        if PROXY:
            args.append(f"--proxy-server={PROXY}")
        if os.path.isdir(USER_DATA_DIR):
            options.add_argument(f'--user-data-dir={USER_DATA_DIR}')
        for a in args:
            options.add_argument(a)
        try:
            driver = uc.Chrome(options=options)
            if USE_STEALTH:
                harden_driver(driver)
            return driver
        except Exception as e:
            print(f"❌ UC Driver init failed: {e}. Falling back to standard Selenium...", "red")

    # 🔹 Standard Selenium fallback with WebDriver Manager
    options = Options()
    for a in ["--start-maximized", "--disable-gpu", "--disable-dev-shm-usage"]:
        options.add_argument(a)
    try:
        options.add_argument("--disable-blink-features=AutomationControlled")
    except Exception:
        pass
    if HEADLESS:
        options.add_argument("--headless=new")
    if PROXY:
        options.add_argument(f"--proxy-server={PROXY}")

    # Use WebDriver Manager to handle ChromeDriver installation
    # Initialize ChromeDriver using Service object
    service = Service(CHROMEDRIVER_PATH)  # Provide ChromeDriver path through Service
    driver = webdriver.Chrome(service=service, options=options)  # Correct initialization
    return driver

# ================= START THE SCRIPT =================
if __name__ == "__main__":
    driver = setup_driver()
    # Continue with your scraping operations as before


# ================= START THE SCRIPT =================
if __name__ == "__main__":
    driver = setup_driver()
    # Continue with your scraping operations as before

    # You can now continue with your scraping operations as before


# ---- PATHS: same folder as script ----
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_PATH  = os.path.join(BASE_DIR, "surya_Table-Lamps.xlsx")
OUTPUT_PATH = os.path.join(BASE_DIR, "surya_Table-Lamps_details.xlsx")
TODO_PATH   = os.path.join(BASE_DIR, "surya_Table-Lamps_todo.json")
MASTER_PATH = os.path.join(BASE_DIR, "surya_Table-Lamps_details_MASTER.xlsx")

# ---- batching ----
BATCH_SIZE = 50
BATCH_BASENAME = "surya_products_details_step"

# ---- Checkpointing / resume ----
CHECKPOINT_EVERY = 10
TAIL_FLUSH_THRESHOLD = 10
RESUME_FROM_PREVIOUS = True
KEEP_BROWSER_ALIVE = True

# ---- “never blank” controls ----
STRICT_NO_BLANKS = True
MAX_ROUNDS = 4
RETRIES_PER_URL = 2

# ---- Slow-site tuning ----
PAGE_RENDER_BUDGET = 45
RENDER_SETTLE_WINDOW = 4
SCROLL_STEP_PX = 900
SCROLL_PAUSE_RANGE = (0.45, 1.0)
BACKOFF_BASE_SECONDS = 7.0
FIELD_MIN_REQUIREMENTS = ["product Name"]

# ---- Throttling ----
BETWEEN_URL_PAUSE = (1.8, 3.6)

# ================= SELECTORS =================
SEL_NAME_PREFERRED = "div.sc-bczRLJ.ffXKVL"
SEL_DESC_CONTAINER = "div.sc-bczRLJ.cfSZvL"
SEL_DESC_LIST = "ul[class^='UnorderedListStyle--'], ul.UnorderedListStyle--1koui9a"
SEL_DESC_ITEMS = "li[class^='ListItemStyle--'], li"
SEL_SIZE_ACTIVE = "button.activeClass span.TypographyStyle--11lquxl.jRSWon"
SEL_SIZE_ANY = "span.TypographyStyle--11lquxl.jRSWon"

CONSENT_XPATHS = [
    "//button[contains(., 'Accept')]", "//button[contains(., 'I Accept')]",
    "//button[contains(., 'I agree')]", "//button[contains(., 'Agree')]",
    "//button[contains(., 'Got it')]", "//button[contains(., 'Allow')]",
    "//button[@id='onetrust-accept-btn-handler']",
]

BLOCK_TEXTS = ["verify you are human", "are you human", "access denied", "unusual traffic", "captcha"]

# ================= GLOBALS =================
BLOCKED_URLS = []          # captcha-blocked URLs
NEW_TASKS = []             # variation tasks
SEEN_URLS_NORM = set()     # all seen urls (main + variation)

# ================= UTILS =================
def cprint(msg, color="default", end="\n"):
    if os.name == "nt":
        print(msg, end=end)
    else:
        COLORS = {
            "green": "\033[92m", "yellow": "\033[93m", "red": "\033[91m",
            "cyan": "\033[96m", "magenta": "\033[95m", "default": ""
        }
        RESET = "\033[0m"
        print(COLORS.get(color, "") + msg + RESET, end=end)

def pause(a, b=None):
    import time as _t, random as _r
    _t.sleep(_r.uniform(a, b) if b is not None else a)

def ensure_debug_chrome_running():
    import time as _t

    def port_open(host, port, timeout=0.5):
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except Exception:
            return False

    if port_open(DEBUG_HOST, DEBUG_PORT):
        return
    os.makedirs(USER_DATA_DIR, exist_ok=True)
    exe = None
    for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]:
        if os.path.exists(p):
            exe = p
            break
    if exe is None:
        exe = "chrome.exe"
    cmd = f'"{exe}" --remote-debugging-port={DEBUG_PORT} --user-data-dir="{USER_DATA_DIR}" --no-first-run --no-default-browser-check'
    try:
        if sys.platform.startswith("win"):
            DETACHED_PROCESS = 0x00000008
            subprocess.Popen(cmd, creationflags=DETACHED_PROCESS)
        else:
            subprocess.Popen(shlex.split(cmd))
    except Exception as e:
        raise RuntimeError(f"Failed to launch Chrome with debug: {e}\nTried: {cmd}")
    for _ in range(40):
        if port_open(DEBUG_HOST, DEBUG_PORT):
            return
        _t.sleep(0.25)
    raise RuntimeError("Chrome debug port not reachable")

def harden_driver(driver):
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": f"""
                Object.defineProperty(navigator, 'webdriver', {{get: () => undefined}});
                window.chrome = {{ runtime: {{}} }};
                Object.defineProperty(navigator, 'languages', {{get: () => {json.dumps(SPOOFED_LANGUAGES)}}});
                Object.defineProperty(navigator, 'plugins', {{get: () => [1,2,3]}});
                const originalQuery = window.navigator.permissions.query;
                window.navigator.permissions.query = (parameters) => (
                    parameters.name === 'notifications' ?
                    Promise.resolve({{ state: 'granted' }}) :
                    originalQuery(parameters)
                );
                Object.defineProperty(window, 'outerWidth',  {{get: () => window.innerWidth + 16}});
                Object.defineProperty(window, 'outerHeight', {{get: () => window.innerHeight + 96}});
            """}
        )
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setUserAgentOverride", {
            "userAgent": SPOOFED_UA,
            "platform": SPOOFED_PLATFORM,
            "acceptLanguage": ",".join(SPOOFED_LANGUAGES),
        })
    except Exception:
        pass

def setup_driver():
    # 🔹 First try Undetected Chrome (UC)
    if USE_UC_FALLBACK:
        options = uc.ChromeOptions()
        args = [
            "--disable-gpu", "--disable-dev-shm-usage", "--no-sandbox",
            "--start-maximized", "--disable-blink-features=AutomationControlled",
        ]
        if HEADLESS:
            args.append("--headless=new")
        if PROXY:
            args.append(f"--proxy-server={PROXY}")
        if os.path.isdir(USER_DATA_DIR):
            options.add_argument(f'--user-data-dir={USER_DATA_DIR}')
        for a in args:
            options.add_argument(a)
        try:
            driver = uc.Chrome(options=options)
            if USE_STEALTH:
                harden_driver(driver)
            return driver
        except Exception as e:
            cprint(f"❌ UC Driver init failed: {e}. Falling back to standard Selenium...", "red")

    # 🔹 Standard Selenium fallback with manual ChromeDriver path
    opts = Options()
    for a in ["--start-maximized", "--disable-gpu", "--disable-dev-shm-usage"]:
        opts.add_argument(a)
    try:
        opts.add_argument("--disable-blink-features=AutomationControlled")
    except Exception:
        pass
    if HEADLESS:
        opts.add_argument("--headless=new")
    if PROXY:
        opts.add_argument(f"--proxy-server={PROXY}")

    if not os.path.exists(CHROMEDRIVER_PATH):
        raise RuntimeError(f"ChromeDriver not found at CHROMEDRIVER_PATH: {CHROMEDRIVER_PATH}")

    try:
        service = Service(CHROMEDRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=opts)
        if USE_STEALTH:
            harden_driver(driver)
        return driver
    except SessionNotCreatedException as e:
        raise RuntimeError(f"❌ Session not created: {e}")
    except WebDriverException as e:
        raise RuntimeError(f"WebDriver init failed: {e}")

def dismiss_overlays(driver):
    for xp in CONSENT_XPATHS:
        try:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    pause(0.15, 0.35)
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    pause(0.2, 0.4)
                    return True
        except Exception:
            pass
    return False

# ============ BLOCK DETECTION ============
def is_blocked(driver):
    try:
        url = driver.current_url.lower()
        title = (driver.title or "").lower()
        patterns = [
            "captcha", "challenge", "verify", "unusual-traffic",
            "access-denied", "just a moment", "attention required", "robot"
        ]
        if any(p in url for p in patterns) or any(p in title for p in patterns):
            return True

        texts = [
            "verify you are human", "are you human", "access denied",
            "unusual traffic", "complete the captcha", "press and hold",
            "just a moment", "checking your browser"
        ]
        for t in texts:
            els = driver.find_elements(
                By.XPATH,
                f"//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{t}')]"
            )
            for el in els:
                try:
                    if el.is_displayed() and el.size.get("height", 0) > 1 and el.size.get("width", 0) > 1:
                        return True
                except Exception:
                    continue
        return False
    except Exception:
        return False

def _product_ready(drv, timeout=25):
    try:
        WebDriverWait(drv, timeout).until(
            EC.any_of(
                EC.visibility_of_element_located((By.CSS_SELECTOR, SEL_NAME_PREFERRED)),
                EC.visibility_of_element_located((By.CSS_SELECTOR, "h1")),
                EC.visibility_of_element_located((By.CSS_SELECTOR, "div[role='heading']"))
            )
        )
        txtlen = len((drv.find_element(By.TAG_NAME, "body").text or "").strip())
        return txtlen > 200 and not is_blocked(drv)
    except Exception:
        return False

def wait_for_human_clear_blocking(driver, url,
                                  prompt_text="🛑 Verification detected. Solve in Chrome, then press ENTER…"):
    import time as _t

    cprint("   ↻ Trying quick refresh to clear UC block…", "yellow")
    try:
        driver.get(url)
    except Exception:
        pass
    if not is_blocked(driver) and _product_ready(driver, timeout=12):
        return

    for i in range(5):
        if not is_blocked(driver) and _product_ready(driver, timeout=15):
            cprint("   ✓ Block cleared automatically or quickly.", "green")
            return

        pause_time = BACKOFF_BASE_SECONDS * (2 ** i) * random.uniform(0.8, 1.4)
        if HEADLESS:
            cprint(f"🛑 Bot check (headless). Sleeping {pause_time:.1f}s then retry…", "magenta")
            _t.sleep(pause_time)
        else:
            cprint("\n" + "=" * 50, "red")
            cprint(f"🛑 BLOCK DETECTED. URL: {url}", "red")
            cprint("➡️ অনুগ্রহ করে **Chrome Window-টি খুলুন** এবং Captcha/ভেরিফিকেশনটি সম্পূর্ণ করুন।", "red")
            cprint("➡️ ভেরিফিকেশন সলভ হওয়ার পর, **এই টার্মিনালে ফিরে এসে ENTER চাপুন**।", "red")
            cprint("=" * 50 + "\n", "red")
            try:
                input(">>> প্রেস ENTER যখন আপনি ক্যাপচা সলভ করবেন এবং পেইজ লোড হবে: ")
            except EOFError:
                _t.sleep(10)

        try:
            driver.get(url)
            WebDriverWait(driver, 25).until(
                lambda d: d.execute_script("return document.readyState") == "complete")
        except Exception:
            pass

        page_scroll_settle(driver, budget_seconds=8, settle_window=1.0)
        if not is_blocked(driver) and _product_ready(driver, timeout=15):
            return

    if is_blocked(driver):
        raise RuntimeError("🛑 Final attempt failed. Still blocked after all retries/manual solve.")

def page_scroll_settle(driver, budget_seconds=PAGE_RENDER_BUDGET, settle_window=RENDER_SETTLE_WINDOW):
    import time as _t, random as _r
    start = _t.time()
    settled_for = 0.0
    last_text_len = 0
    last_height = 0
    while _t.time() - start < budget_seconds:
        dismiss_overlays(driver)
        driver.execute_script("window.scrollBy(0, arguments[0]);", SCROLL_STEP_PX)
        pause(*SCROLL_PAUSE_RANGE)
        driver.execute_script("window.scrollBy(0, arguments[0]);", -SCROLL_STEP_PX // 3)
        pause(0.15, 0.35)
        if _r.random() < 0.25:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            pause(*SCROLL_PAUSE_RANGE)
        try:
            text_len = len(driver.find_element(By.TAG_NAME, "body").text)
        except Exception:
            text_len = 0
        try:
            height = driver.execute_script("return document.body.scrollHeight || 0;") or 0
        except Exception:
            height = 0
        if abs(text_len - last_text_len) < 30 and abs(height - last_height) < 50:
            settled_for += 0.5
        else:
            settled_for = 0.0
        last_text_len, last_height = text_len, height
        if settled_for >= settle_window:
            break

def get_text_safe(el):
    try:
        return el.text.strip()
    except Exception:
        return ""

# ================= DIMENSION PARSER =================
def parse_dimension_fields(dim_text: str):
    """
    Dimension theke 5 column:
    Width, Depth, Diameter, Length, Height
    - H -> Height
    - W -> Width
    - D -> Depth
    - Dia/Diameter/Ø -> Diameter
    - L/Length -> Length
    - Jodi '2' x 3'' type format hoy ar label na thake:
      1st = Length, 2nd = Width
    """
    if not dim_text:
        return "", "", "", "", ""

    t = (dim_text or "").replace("”", '"').replace("“", '"').replace("′", "'").replace("″", '"')
    width = depth = diameter = length = height = ""

    # Height (H)
    m = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*[Hh]\b', t)
    if m:
        height = m.group(1)

    # Width (W)
    m = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*[Ww]\b', t)
    if m:
        width = m.group(1)

    # Depth (D)
    m = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*[Dd]\b', t)
    if m:
        depth = m.group(1)

    # Diameter
    m = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*(?:Dia|Diameter|Ø)\b', t, flags=re.I)
    if m:
        diameter = m.group(1)

    # Length (L / Length)
    m = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*(?:[Ll]\b|Length\b)', t, flags=re.I)
    if m:
        length = m.group(1)

    # Fallback: 2' x 3' (no labels)
    if not (length or width):
        m = re.search(r"([\d.'\"]+)\s*[xX]\s*([\d.'\"]+)", t)
        if m:
            length = m.group(1).strip()
            width = m.group(2).strip()

    return width, depth, diameter, length, height

# ================= DETAILS PARSER =================
def parse_details_fields(details: str):
    """
    Details column theke:
      Cushion, Seat Depth, Seat Width, Seat Height, Socket, Wattage
    """
    out = {
        "Cushion": "",
        "Seat Depth": "",
        "Seat Width": "",
        "Seat Height": "",
        "Socket": "",
        "Wattage": "",
    }
    if not details:
        return out

    text = details

    # --- Cushion / Seats ---
    m = re.search(r'Cushion[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Cushion"] = m.group(1).strip()

    m = re.search(r'Seat\s*Depth[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Seat Depth"] = m.group(1).strip()

    m = re.search(r'Seat\s*Width[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Seat Width"] = m.group(1).strip()

    m = re.search(r'Seat\s*Height[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Seat Height"] = m.group(1).strip()

    # --- Socket (E-12 Socket / E26 Socket / G9 Socket / etc.) ---
    # ধরবে: "E-12 Socket", "E26 Socket", "G9 Socket" ইত্যাদি
    m = re.search(r'((?:[A-Za-z]+\s*-?\s*\d+)\s*Socket)', text, flags=re.I)
    if m:
        out["Socket"] = m.group(1).strip()
    else:
        # fallback: "Socket: E-12" টাইপ কিছু থাকলে
        m = re.search(r'Socket[^:]*:\s*([^;/]+)', text, flags=re.I)
        if m:
            out["Socket"] = m.group(1).strip()

    # --- Wattage (40 Max Wattage / 60 Wattage / 40W / 60 Watts / 40 Watt) ---
    m = re.search(r'(\d+)\s*(?:Max\s*)?(?:Wattage|Watt|Watts|W)\b', text, flags=re.I)
    if m:
        num = m.group(1)
        # সব time normalize করে রাখছি: "40 Max Wattage"
        out["Wattage"] = f"{num} Max Wattage"

    return out

# ================= DESCRIPTION PARSER =================
def parse_description_fields(description: str):
    """
    Description theke:
      Shade Details, Finish, Base
    """
    out = {
        "Shade Details": "",
        "Finish": "",
        "Base": ""
    }

    if not description:
        return out

    text = description

    # Shade: White Linen
    m = re.search(r'Shade[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Shade Details"] = m.group(1).strip()

    # Finish / Body Finish: Electroplated Metallic - Gold
    m = re.search(r'Finish[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Finish"] = m.group(1).strip()

    # Base: Brown Natural Oak
    m = re.search(r'Base[^:]*:\s*([^;]+)', text, flags=re.I)
    if m:
        out["Base"] = m.group(1).strip()

    return out

# ================= FIELD EXTRACTORS =================
def get_product_name(driver):
    try:
        el = driver.find_element(By.CSS_SELECTOR, SEL_NAME_PREFERRED)
        nm = get_text_safe(el)
        if nm:
            return nm
    except Exception:
        pass
    for css in ["h1", "div[role='heading']"]:
        try:
            nm = get_text_safe(driver.find_element(By.CSS_SELECTOR, css))
            if nm:
                return nm
        except Exception:
            pass
    return ""

def get_description(driver):
    items = []
    try:
        containers = driver.find_elements(By.CSS_SELECTOR, SEL_DESC_CONTAINER)
        for c in containers:
            uls = c.find_elements(By.CSS_SELECTOR, SEL_DESC_LIST)
            for ul in uls:
                lis = ul.find_elements(By.CSS_SELECTOR, SEL_DESC_ITEMS)
                for li in lis:
                    t = get_text_safe(li)
                    if t:
                        items.append(t)
    except Exception:
        pass
    if not items:
        try:
            lis = driver.find_elements(By.CSS_SELECTOR, "ul li")
            for li in lis[:20]:
                t = get_text_safe(li)
                if t and len(t) < 200:
                    items.append(t)
        except Exception:
            pass
    return "; ".join(dict.fromkeys(items))

def find_size_line_candidates(text):
    lines = []
    for line in (text or "").splitlines():
        if re.search(r'\d', line) and re.search(r'(?:\b[HWD]\b|Dia|Diameter|Ø)', line, flags=re.I):
            lines.append(line.strip())
    return lines[:3]

def get_size_text(driver):
    try:
        t = get_text_safe(driver.find_element(By.CSS_SELECTOR, SEL_SIZE_ACTIVE))
        if t:
            return t
    except Exception:
        pass
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, SEL_SIZE_ANY):
            t = get_text_safe(el)
            if re.search(r'\d', t) and re.search(r'(?:\b[HWD]\b|Dia|Diameter|Ø)', t, flags=re.I):
                return t
    except Exception:
        pass
    try:
        txt = driver.find_element(By.TAG_NAME, "body").text
        cands = find_size_line_candidates(txt)
        if cands:
            return cands[0]
    except Exception:
        pass
    return ""

def get_weight(driver):
    try:
        txt = driver.find_element(By.TAG_NAME, "body").text
        m = re.search(r'Weight[^0-9]*?(\d+(?:\.\d+)?)', txt, flags=re.I)
        return m.group(1) if m else ""
    except Exception:
        return ""

def get_dimension(driver):
    try:
        return get_size_text(driver)
    except Exception:
        return ""

def get_list_price(driver):
    try:
        spans = driver.find_elements(
            By.CSS_SELECTOR,
            "div.sc-bczRLJ.iUstDD span.TypographyStyle-sc-11lquxl.kmkpzB"
        )
        for el in reversed(spans):
            t = get_text_safe(el)
            if "$" in t:
                return t
        if spans:
            return get_text_safe(spans[-1])
    except Exception:
        pass
    return ""

def get_details(driver):
    # (same as earlier – unchanged, keeps Details text)
    try:
        accordions = driver.find_elements(By.CSS_SELECTOR, "dl.AccordionStyle-sc-1ijmlek.indofj")
        for acc in accordions:
            try:
                header = acc.find_element(By.CSS_SELECTOR, "dt[data-test-selector='sectionHeader']")
                key = (header.get_attribute("data-test-key") or "").strip().lower()
                header_text = get_text_safe(header).lower()
                if key != "details" and "details" not in header_text:
                    continue

                try:
                    btn = header.find_element(By.TAG_NAME, "button")
                    expanded = (btn.get_attribute("aria-expanded") or "").lower()
                    if expanded != "true":
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                        pause(0.1, 0.3)
                        try:
                            btn.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", btn)
                        time.sleep(0.5)
                except Exception:
                    pass

                try:
                    panel = acc.find_element(
                        By.CSS_SELECTOR,
                        "dd[data-test-selector='sectionPanel']"
                    )
                except Exception:
                    panel = acc

                lis = panel.find_elements(By.CSS_SELECTOR, "li")
                items = [get_text_safe(li) for li in lis if get_text_safe(li)]

                if not items:
                    raw = get_text_safe(panel)
                    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
                    items = lines

                if items:
                    return "; ".join(items)
            except Exception:
                continue
    except Exception:
        pass

    try:
        title_els = driver.find_elements(
            By.XPATH,
            "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'details')]"
        )
        for h in title_els:
            try:
                acc = h.find_element(By.XPATH, "ancestor::dl[contains(@class,'AccordionStyle-sc-1ijmlek')]")
            except Exception:
                continue
            try:
                btn = acc.find_element(By.TAG_NAME, "button")
                expanded = (btn.get_attribute("aria-expanded") or "").lower()
                if expanded != "true":
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                    pause(0.1, 0.3)
                    try:
                        btn.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
            except Exception:
                pass

            lis = acc.find_elements(By.CSS_SELECTOR, "li")
            items = [get_text_safe(li) for li in lis if get_text_safe(li)]
            if not items:
                try:
                    panel = acc.find_element(By.CSS_SELECTOR, "dd")
                    raw = get_text_safe(panel)
                    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
                    items = lines
                except Exception:
                    pass
            if items:
                return "; ".join(items)
    except Exception:
        pass

    try:
        blocks = driver.find_elements(By.CSS_SELECTOR, "div.sc-bczRLJ.YNZZW")
        for blk in blocks:
            try:
                title_spans = blk.find_elements(
                    By.CSS_SELECTOR,
                    "span.TypographyStyle-sc-11lquxl.ehGNvU"
                )
                txt = " ".join(get_text_safe(s) for s in title_spans).lower()
                if "details" not in txt:
                    continue
                lis = blk.find_elements(By.CSS_SELECTOR, "li")
                items = [get_text_safe(li) for li in lis if get_text_safe(li)]
                if not items:
                    raw = get_text_safe(blk)
                    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
                    items = lines
                if items:
                    return "; ".join(items)
            except Exception:
                continue
    except Exception:
        pass

    return ""

# ============ VARIATION EXTRACTOR ============
def get_variation_tasks(driver):
    variations = []
    try:
        blocks = driver.find_elements(By.CSS_SELECTOR, "div.sc-bczRLJ.dXFjGj")
        for blk in blocks:
            try:
                title_span = blk.find_element(By.CSS_SELECTOR, "span.TypographyStyle-sc-11lquxl.fDPByD")
                if "colors" not in get_text_safe(title_span).lower():
                    continue
            except Exception:
                continue

            try:
                anchors = blk.find_elements(By.CSS_SELECTOR, "a.StyledButton-sc-1y32st")
            except Exception:
                anchors = []
            for a in anchors:
                try:
                    href = a.get_attribute("href") or ""
                    if not href:
                        href = a.get_attribute("data-href") or ""
                    if not href:
                        continue
                    variation_url = urljoin(driver.current_url, href)

                    try:
                        img_el = a.find_element(By.TAG_NAME, "img")
                        variation_img = img_el.get_attribute("src") or ""
                    except Exception:
                        variation_img = ""

                    variation_sku = ""
                    try:
                        sku_div = a.find_element(By.CSS_SELECTOR, "div.sc-bczRLJ.kLzVtl")
                        variation_sku = get_text_safe(sku_div)
                    except Exception:
                        pass
                    if not variation_sku:
                        variation_sku = a.get_attribute("aria-label") or ""

                    variations.append({
                        "url": variation_url,
                        "img": variation_img,
                        "sku": variation_sku,
                    })
                except Exception:
                    continue
    except Exception:
        pass
    return variations

# ================= SAVE HELPERS =================
def _atomic_save_df(df: pd.DataFrame, path: str):
    import time as _t
    dirn = os.path.dirname(path) or "."
    base = os.path.basename(path)
    tmp = os.path.join(dirn, f"~${base}.tmp.xlsx")
    for _ in range(3):
        try:
            df.to_excel(tmp, index=False)
            os.replace(tmp, path)
            return
        except Exception:
            _t.sleep(0.8)
    raise

# ================= SCRAPE ONE =================
def extract_details_for_url(driver, url, image_from_list="", sku_from_list=""):
    global BLOCKED_URLS, NEW_TASKS, SEEN_URLS_NORM

    out = {
        "Product Url": url,
        "Image Url": image_from_list,
        "product Name": "",
        "Sku": sku_from_list,
        "Product Family Name": "",
        "Description": "",
        "List Price": "",
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Length": "",
        "Height": "",
        "Cushion":"",
        "Seat Depth": "",
        "Seat Width": "",
        "Seat Height": "",
        "Socket": "",
        "Wattage": "",
        "Shade Details": "",
        "Finish": "",
        "Base": "",
    }

    cprint(f"   → GET {url}", "cyan")
    try:
        driver.get(url)
    except Exception:
        pass

    if is_blocked(driver):
        cprint("🛑 Block detected on first load.", "red")
        try:
            wait_for_human_clear_blocking(driver, url)
        except RuntimeError as e:
            cprint(f"   ⚠️ {e}", "red")
            if url not in BLOCKED_URLS:
                BLOCKED_URLS.append(url)
            if not HEADLESS:
                cprint("➡️ ব্রাউজার উইন্ডোতে গিয়ে CAPTCHA/ভেরিফিকেশন সল্ভ করুন।", "yellow")
                cprint("  এরপর এই টার্মিনালে ফিরে এসে Enter চাপুন, অথবা 's' লিখে Enter দিলে অপেক্ষা না করে এ URL স্কিপ হবে.", "yellow")
                try:
                    val = input("সমাধান করে Enter চাপুন / বা 's' + Enter (skip): ").strip().lower()
                except EOFError:
                    val = ""
                if val == "s":
                    cprint("↷ Skipping this URL for now.", "yellow")
                    return out
                try:
                    driver.get(url)
                    WebDriverWait(driver, 12).until(
                        lambda d: d.execute_script("return document.readyState") == "complete")
                except Exception:
                    pass
                if is_blocked(driver):
                    cprint("✖ এখনও ব্লক আছে বা পেইজ ঠিকমত লোড হয়নি — স্কিপ করা হচ্ছে (URL তালিকায় যোগ করা হলো)।", "red")
                    if url not in BLOCKED_URLS:
                        BLOCKED_URLS.append(url)
                    return out
            else:
                return out

    try:
        WebDriverWait(driver, 25).until(
            lambda d: d.execute_script("return document.readyState") == "complete")
    except Exception:
        pass

    _product_ready(driver, timeout=25)
    page_scroll_settle(driver)

    if is_blocked(driver):
        cprint("   ✖ Blocked again at parse phase.", "yellow")
        if url not in BLOCKED_URLS:
            BLOCKED_URLS.append(url)
        return out

    name = get_product_name(driver)
    desc = get_description(driver)
    dimension_str = get_dimension(driver)
    wt = get_weight(driver)
    list_price = get_list_price(driver)
    details_text = get_details(driver)

    product_family_name = name

    # Dimension: 5 value
    width, depth, diameter, length, height = parse_dimension_fields(dimension_str)

    # Details fields
    details_fields = parse_details_fields(details_text)

    # Description-based fields
    desc_fields = parse_description_fields(desc)

    out.update({
        "product Name": name,
        "Product Family Name": product_family_name,
        "Description": desc,
        "List Price": list_price,
        "Weight": wt,
        "Width": width,
        "Depth": depth,
        "Diameter": diameter,
        "Length": length,
        "Height": height,
        "Seat Depth": details_fields.get("Seat Depth", ""),
        "Seat Width": details_fields.get("Seat Width", ""),
        "Seat Height": details_fields.get("Seat Height", ""),
        "Socket": details_fields.get("Socket", ""),
        "Cushion": details_fields.get("Cushion", ""),
        "Wattage": details_fields.get("Wattage", ""),
        "Shade Details": desc_fields.get("Shade Details", ""),
        "Finish": desc_fields.get("Finish", ""),
        "Base": desc_fields.get("Base", ""),
    })

    cprint(f"      Name       : {name or '—'}", "green" if name else "yellow")
    cprint(f"      Dimension  : {dimension_str or '—'}", "default")
    cprint(f"      W/D/Dia/L/H: {width or '—'} / {depth or '—'} / {diameter or '—'} / {length or '—'} / {height or '—'}", "default")
    cprint(f"      Weight     : {wt or '—'}", "default")
    cprint(f"      List Price : {list_price or '—'}", "default")
    cprint(f"      Seat D/W/H : {out['Seat Depth'] or '—'} / {out['Seat Width'] or '—'} / {out['Seat Height'] or '—'}", "default")
    cprint(f"      Socket/W   : {out['Socket'] or '—'} / {out['Wattage'] or '—'}", "default")
    cprint(f"      Details.len: {len(details_text)}", "default")
    cprint(f"      Desc.len   : {len(desc)}", "default")
    cprint(f"      Shade      : {out['Shade Details'] or '—'}", "default")
    cprint(f"      Finish     : {out['Finish'] or '—'}", "default")
    cprint(f"      Base       : {out['Base'] or '—'}", "default")

    variations = get_variation_tasks(driver)
    if variations:
        cprint(f"      Variations : {len(variations)} found", "magenta")
    for v in variations:
        v_url = (v.get("url") or "").strip()
        if not v_url:
            continue
        v_url_norm = v_url.lower()
        if v_url_norm == url.lower():
            continue
        if v_url_norm in SEEN_URLS_NORM:
            continue
        SEEN_URLS_NORM.add(v_url_norm)
        NEW_TASKS.append({
            "url": v_url,
            "img": v.get("img", ""),
            "sku": v.get("sku", ""),
        })

    return out

def needs_retry(record):
    for k in FIELD_MIN_REQUIREMENTS:
        if not (record.get(k) or "").strip():
            return True
    return False

def backoff_sleep(base, attempt):
    import random as _r, time as _t
    pause_time = base * (2 ** (attempt - 1)) * _r.uniform(0.75, 1.35)
    cprint(f"      …backoff {pause_time:.1f}s", "magenta")
    _t.sleep(pause_time)

# ============ BATCHING HELPERS ============
def _batch_filename(idx: int) -> str:
    return os.path.join(BASE_DIR, f"{BATCH_BASENAME}{idx:02d}.xlsx")

def _count_existing_batches() -> int:
    pattern = os.path.join(BASE_DIR, f"{BATCH_BASENAME}*.xlsx")
    files = sorted(glob.glob(pattern))
    return len(files)

def _save_batch_if_needed(all_results: list, out_cols: list, last_batch_saved_count: int) -> int:
    total_done = len(all_results)
    to_save = total_done - last_batch_saved_count
    if to_save >= BATCH_SIZE:
        n_new_batches = to_save // BATCH_SIZE
        start_batch_idx = _count_existing_batches() + 1
        for j in range(n_new_batches):
            batch_idx = start_batch_idx + j
            start = last_batch_saved_count + j * BATCH_SIZE
            end = start + BATCH_SIZE
            df_batch = pd.DataFrame(all_results[start:end], columns=out_cols)
            batch_path = _batch_filename(batch_idx)
            _atomic_save_df(df_batch, batch_path)
            cprint(f"📦 Saved batch ({batch_idx:02d}) → {batch_path}", "magenta")
        last_batch_saved_count += n_new_batches * BATCH_SIZE
    return last_batch_saved_count

def _build_master_from_input_and_results(df_input: pd.DataFrame, results: list, out_cols: list) -> pd.DataFrame:
    cols = {c.lower().strip(): c for c in df_input.columns}

    def colget(*names):
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    col_url = colget("Product URL", "Product Url", "url")
    col_img = colget("Image URL", "Image Url", "image")
    col_sku = colget("SKU", "Sku")

    base = df_input[[col_url, col_img, col_sku]].copy()
    base.columns = ["Product Url", "Image Url", "Sku"]

    df_res = pd.DataFrame(results, columns=out_cols)
    if df_res.empty:
        df_res = pd.DataFrame(columns=out_cols)

    master = base.merge(
        df_res.drop(columns=["Image Url"], errors="ignore"),
        on="Product Url",
        how="left",
        suffixes=("", "_res")
    )

    for col in out_cols:
        if col not in master.columns:
            master[col] = ""

    master = master[out_cols]
    return master

# ================= MAIN =================
def main():
    global NEW_TASKS, SEEN_URLS_NORM

    if not os.path.exists(INPUT_PATH):
        cprint(f"❌ Input file not found: {INPUT_PATH}", "red")
        return

    NEW_TASKS = []
    SEEN_URLS_NORM = set()

    cprint(f"📄 Reading: {INPUT_PATH}", "cyan")
    df_raw = pd.read_excel(INPUT_PATH)

    cols = {c.lower().strip(): c for c in df_raw.columns}

    def colget(*names):
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    col_url = colget("Product URL", "Product Url", "url")
    col_img = colget("Image URL", "Image Url", "image")
    col_sku = colget("SKU", "Sku")
    if not col_url or not col_img or not col_sku:
        cprint("❌ Need columns: Product URL, Image URL, SKU in input file.", "red")
        return

    tasks = []
    for _, row in df_raw.iterrows():
        raw_url = row.get(col_url, "")
        if pd.isna(raw_url):
            continue
        url = str(raw_url).strip()
        if not url or url.lower() in ("nan", "none", "null"):
            continue
        img = "" if pd.isna(row.get(col_img, "")) else str(row.get(col_img, "")).strip()
        sku = "" if pd.isna(row.get(col_sku, "")) else str(row.get(col_sku, "")).strip()
        tasks.append({"url": url, "img": img, "sku": sku})
        SEEN_URLS_NORM.add(url.lower())

    if not tasks:
        cprint("Nothing to do.", "yellow")
        return

    # Final output order (with Description, Shade Details, Finish, Base)
    out_cols = [
        "Product Url",
        "Image Url",
        "product Name",
        "Sku",
        "Product Family Name",
        "Description",
        "List Price",
        "Weight",
        "Width",
        "Depth",
        "Diameter",
        "Length",
        "Height",
        "Cushion",
        "Seat Depth",
        "Seat Width",
        "Seat Height",
        "Socket",
        "Wattage",
        "Shade Details",
        "Finish",
        "Base",
    ]

    results = []
    done_urls_norm = set()
    if RESUME_FROM_PREVIOUS and os.path.exists(OUTPUT_PATH):
        try:
            prev = pd.read_excel(OUTPUT_PATH)
            if set(out_cols).issubset(prev.columns):
                results = prev[out_cols].to_dict(orient="records")
                for u in prev["Product Url"].astype(str).fillna(""):
                    un = u.strip().lower()
                    done_urls_norm.add(un)
                    SEEN_URLS_NORM.add(un)
                cprint(f"↺ Resume: {len(done_urls_norm)} URLs already saved.", "yellow")
        except Exception:
            pass

    dq = collections.deque([t for t in tasks if t["url"].lower() not in done_urls_norm])
    attempts = collections.Counter()

    driver = setup_driver()

    last_batch_saved_count = _count_existing_batches() * BATCH_SIZE

    def flush_results():
        nonlocal last_batch_saved_count
        try:
            _atomic_save_df(pd.DataFrame(results, columns=out_cols), OUTPUT_PATH)
            cprint(f"💾 Saved (rolling) → {OUTPUT_PATH}", "magenta")
        except Exception as e:
            cprint(f"   ⚠️ Save error: {e}", "red")

        last_batch_saved_count = _save_batch_if_needed(results, out_cols, last_batch_saved_count)

    total = len(dq) + len(done_urls_norm)
    processed_since_flush = 0

    try:
        round_no = 1
        while dq and round_no <= MAX_ROUNDS:
            cprint(f"\n===== ROUND {round_no} (queue size: {len(dq)}) =====", "cyan")
            q_len_this_round = len(dq)
            for _ in range(q_len_this_round):
                task = dq.popleft()
                url, img, sku = task["url"], task["img"], task["sku"]
                url_norm = url.lower()
                if url_norm in done_urls_norm:
                    continue

                cprint(f"\n[{len(done_urls_norm) + 1}/{total}] Processing", "cyan")
                cprint(f"URL : {url}", "default")
                cprint(f"SKU : {sku or '—'}", "default")

                ok = False
                for attempt in range(1, RETRIES_PER_URL + 1):
                    try:
                        rec = extract_details_for_url(driver, url, img, sku)
                    except Exception as e:
                        cprint(f"   ⚠️ Error: {e}", "red")
                        rec = {
                            "Product Url": url,
                            "Image Url": img,
                            "product Name": "",
                            "Sku": sku,
                            "Product Family Name": "",
                            "Description": "",
                            "List Price": "",
                            "Weight": "",
                            "Width": "",
                            "Depth": "",
                            "Diameter": "",
                            "Length": "",
                            "Height": "",
                            "Cushion" : "",
                            "Seat Depth": "",
                            "Seat Width": "",
                            "Seat Height": "",
                            "Socket": "",
                            "Wattage": "",
                            "Shade Details": "",
                            "Finish": "",
                            "Base": "",
                        }

                    if not needs_retry(rec):
                        results.append(rec)
                        done_urls_norm.add(url_norm)
                        ok = True
                        break
                    else:
                        cprint("   ✖ Missing key fields", "yellow")
                        if is_blocked(driver):
                            cprint("   ⚠️ Verification hit mid-run. Will requeue and mark blocked.", "yellow")
                            if url not in BLOCKED_URLS:
                                BLOCKED_URLS.append(url)
                        if attempt < RETRIES_PER_URL:
                            try:
                                driver.refresh()
                            except Exception:
                                pass
                            page_scroll_settle(
                                driver,
                                budget_seconds=PAGE_RENDER_BUDGET + 15,
                                settle_window=RENDER_SETTLE_WINDOW + 2
                            )
                            backoff_sleep(BACKOFF_BASE_SECONDS, attempt + 1)

                if not ok:
                    attempts[url_norm] += 1
                    dq.append(task)

                if NEW_TASKS:
                    cprint(f"➕ Adding {len(NEW_TASKS)} variation tasks to queue", "magenta")
                    for vt in NEW_TASKS:
                        dq.append(vt)
                    NEW_TASKS = []

                processed_since_flush += 1
                remaining_est = len(dq)
                if (processed_since_flush % CHECKPOINT_EVERY == 0) or (remaining_est <= TAIL_FLUSH_THRESHOLD):
                    flush_results()
                    processed_since_flush = 0
                pause(*BETWEEN_URL_PAUSE)

            round_no += 1

        if dq and STRICT_NO_BLANKS:
            cprint(f"\n🧹 Interactive cleanup for last {len(dq)} URLs (STRICT_NO_BLANKS)", "magenta")
            while dq:
                task = dq.popleft()
                url, img, sku = task["url"], task["img"], task["sku"]
                url_norm = url.lower()
                if url_norm in done_urls_norm:
                    continue

                cprint(f"\n[CLEANUP] {url}", "cyan")
                try:
                    driver.get(url)
                except Exception:
                    pass
                wait_for_human_clear_blocking(
                    driver,
                    url,
                    "🛑 Solve verification for this product, then press ENTER…"
                )
                try:
                    WebDriverWait(driver, 20).until(
                        lambda d: d.execute_script("return document.readyState") == "complete")
                except Exception:
                    pass
                page_scroll_settle(
                    driver,
                    budget_seconds=PAGE_RENDER_BUDGET + 15,
                    settle_window=RENDER_SETTLE_WINDOW + 1
                )

                for _ in range(6):
                    rec = extract_details_for_url(driver, url, img, sku)
                    if not needs_retry(rec):
                        results.append(rec)
                        done_urls_norm.add(url_norm)
                        cprint("   ✓ Captured after verification.", "green")
                        break
                    else:
                        cprint("   …still missing, refreshing once more.", "yellow")
                        try:
                            driver.refresh()
                        except Exception:
                            pass
                        page_scroll_settle(
                            driver,
                            budget_seconds=PAGE_RENDER_BUDGET + 10,
                            settle_window=RENDER_SETTLE_WINDOW + 1
                        )

                if url_norm not in done_urls_norm:
                    wait_for_human_clear_blocking(
                        driver,
                        url,
                        "🔁 Still missing fields. Solve any prompts and press ENTER…"
                    )
                    rec = extract_details_for_url(driver, url, img, sku)
                    results.append(rec)
                    done_urls_norm.add(url_norm)
                    cprint("   ✓ Captured via manual confirm.", "green")

                flush_results()

        flush_results()
        cprint(f"\n✅ Completed scraping. Total rows in rolling output: {len(results)}", "green")

        try:
            df_master = _build_master_from_input_and_results(df_raw, results, out_cols)
            _atomic_save_df(df_master, MASTER_PATH)
            cprint(f"📚 Master file saved → {MASTER_PATH}", "magenta")
        except Exception as e:
            cprint(f"⚠️ Master merge failed: {e}", "red")

    finally:
        try:
            todo = [t for t in dq]
            with open(TODO_PATH, "w", encoding="utf-8") as f:
                json.dump(todo, f, ensure_ascii=False, indent=2)
            if todo:
                cprint(f"📝 Remaining queue saved → {TODO_PATH}", "yellow")
        except Exception as e:
            cprint(f"   ⚠️ Could not write TODO snapshot: {e}", "red")

        try:
            if BLOCKED_URLS:
                blocked_path = os.path.join(BASE_DIR, "surya_blocked_urls.txt")
                with open(blocked_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(BLOCKED_URLS))
                blocked_json = os.path.join(BASE_DIR, "surya_blocked_urls.json")
                with open(blocked_json, "w", encoding="utf-8") as f:
                    json.dump([{"url": u} for u in BLOCKED_URLS], f, ensure_ascii=False, indent=2)
                cprint(f"🚧 Blocked URLs saved → {blocked_path}", "yellow")
                cprint(f"🚧 Blocked URLs (json) → {blocked_json}", "yellow")
        except Exception as e:
            cprint(f"⚠️ Could not save blocked URLs: {e}", "red")

        try:
            if not KEEP_BROWSER_ALIVE:
                driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
