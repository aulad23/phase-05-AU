# AGENT BRAIN — Decision Framework

নতুন vendor পেলে এই Brain.md পড়ো এবং প্রতিটা decision নিচের rules মেনে নাও।
User-এর মতো করে ভাবো — তারপর instruction.json বানাও।

---

## DECISION 1 — Site Type Detection

Website-এ গিয়ে চেক করো:

| Signal | Site Type |
|--------|-----------|
| HTML-এ `Shopify.theme` বা `cdn.shopify.com` | **Shopify** |
| `/wp-content/` বা `wp-json` URL | **WordPress / WooCommerce** |
| Page source-এ `window.__NUXT__` বা `__NEXT_DATA__` | **Next.js / Nuxt** |
| কিছু না মিললে | **Custom** |

**Shortcut:** URL-এ `.myshopify.com` redirect হলে → Shopify নিশ্চিত।

---

## DECISION 2 — Scraping Method Selection

Site type অনুযায়ী method বেছে নাও:

```
Shopify
  └── /collections/{handle}/products.json API test করো
        ├── Response আসলে → METHOD = "Shopify API" (Selenium লাগবে না)
        └── 403/empty → METHOD = "Selenium + BS4"

WordPress / WooCommerce
  └── /wp-json/wc/v3/products API test করো
        ├── আসলে → METHOD = "WooCommerce API"
        └── না আসলে → METHOD = "Requests + BS4"

Next.js / Nuxt
  └── Network tab দেখো → XHR/fetch calls আছে কিনা
        ├── JSON API পেলে → METHOD = "Requests + JSON API"
        └── না পেলে → METHOD = "Selenium + BS4"

Custom
  └── Pagination কীভাবে কাজ করে?
        ├── ?page=N → METHOD = "Requests + BS4 + Query Param"
        ├── Button click → METHOD = "Selenium + BS4"
        └── Infinite scroll → METHOD = "Selenium + scroll"
```

**Rule:** Selenium শেষ option — যতটা সম্ভব API বা Requests দিয়ে করো।

---

## DECISION 3 — Column Availability Check

Sample product fetch করে নিচের fields খোঁজো:

| Column | Shopify API Field | Product Page | Notes |
|--------|------------------|--------------|-------|
| SKU | `variants[0].sku` | spec table | API-তেই আছে |
| Price | `variants[0].price` | — | API-তেই আছে |
| Weight | `variants[0].grams` → lbs | spec table | grams ÷ 453.592 |
| Image URL | `images[0].src` | — | CDN URL নাও |
| Product Name | `title` | — | API-তেই আছে |
| Description | `body_html` → strip HTML | — | HTML clean করো |
| Width | — | spec table | page scrape |
| Depth | — | spec table | page scrape |
| Height | — | spec table | page scrape |
| Diameter | — | spec table | শুধু round items |
| Seat Width/Depth/Height | — | spec table | furniture only |
| Arm Height/Width | — | spec table | chair/sofa only |
| Finish | `options` (Color/Finish) বা spec table | both | check করো |
| Materials | `tags` বা spec table | both | check করো |
| Tags | `tags` list | — | comma join |

**Decision Rule:**
- API-তে পেলে → API থেকে নাও (faster, no page load)
- না পেলে → product page scrape করো
- কোনোভাবেই না পেলে → সেই column blank রাখো, note করো

---

## DECISION 4 — Pagination Strategy

Website-এ category page খুলে দেখো:

```
URL-এ ?page=2 কাজ করে?
  ├── হ্যাঁ → PAGINATION = "query_param"
  └── না →
        "Load More" button আছে?
          ├── হ্যাঁ → PAGINATION = "button_click"
          └── না →
                Scroll করলে নতুন product আসে?
                  ├── হ্যাঁ → PAGINATION = "infinite_scroll"
                  └── না → PAGINATION = "single_page"

Shopify API?
  → সবসময় PAGINATION = "api_page_param" (?limit=250&page=N)
```

---

## DECISION 5 — Duplicate Handling

```
একই product কি multiple categories-এ থাকতে পারে?
  ├── হ্যাঁ (যেমন "Sofa" → Living Room + Sofas) →
  │     DEMO_MODE: allow duplicates (সব sheet-এ data দেখাও)
  │     FULL_MODE: seen_handles set দিয়ে skip
  └── না → dedup logic দরকার নেই
```

---

## DECISION 6 — Rate Limiting

```
Website কি bot block করে?
  ├── Cloudflare/Captcha দেখা যাচ্ছে → time.sleep(2-3) + headers
  ├── 429 Too Many Requests → time.sleep(5) + exponential backoff
  ├── Normal site → time.sleep(0.5) enough
  └── API (Shopify) → time.sleep(0.3) between pages
```

---

## DECISION 7 — Category Name Resolution

```
Vendor Excel-এ category কীভাবে আছে?
  → সেই নামই instruction.json-এ রাখো (EXACTLY)
  → Website-এর category name আলাদা হলে → "notes" field-এ লেখো
  → Link = website-এর actual URL (Excel-এ link থাকলে সেটা নাও)
```

---

## BRAIN OUTPUT FORMAT

সব decision নেওয়ার পরে এই format-এ instruction.json বানাও:

```json
{
  "vendor": "Brand Name",
  "url": "https://www.brandsite.com",
  "site_type": "shopify",
  "scraping_method": "Shopify API + Requests+BS4 (product page specs)",
  "pagination": "api_page_param",
  "rate_limit_seconds": 0.3,
  "demo_per_category": 3,
  "dedup": true,
  "categories": [
    {
      "name": "Exact Name from Vendor Excel",
      "handle": "shopify-collection-handle",
      "link": "https://www.brandsite.com/collections/handle"
    }
  ],
  "columns_available": {
    "from_api": ["SKU", "Price", "Weight", "Image URL", "Product Name", "Description", "Tags"],
    "from_product_page": ["Width", "Depth", "Height", "Seat Height", "Finish", "Materials"],
    "not_available": ["Diameter", "Arm Height"]
  },
  "spec_selector": ".specs-list li",
  "notes": "কোনো বিশেষ বিষয় এখানে লেখো"
}
```

---

## BRAIN THINKING TEMPLATE

নতুন vendor পেলে মনে মনে এই প্রশ্নগুলো করো:

1. **"এটা কোন ধরনের site?"** → Site type detect
2. **"API আছে?"** → Test করো আগে
3. **"কোন data কোথায় আছে?"** → API vs page
4. **"Categories কটা, কীভাবে paginate করে?"** → Strategy
5. **"Vendor Excel-এর সাথে মিলছে?"** → Category name match
6. **"আগে কি এই ধরনের site করেছি?"** → Memory check
7. **"কোথায় সমস্যা হতে পারে?"** → Notes-এ লেখো

এই ৭টা প্রশ্নের উত্তর = instruction.json ready।

---

## STANDARD COLUMN FORMAT — CLIENT UPDATE (MANDATORY)

**সব নতুন code এই format follow করবে। পুরানো code update করতে বললে এই format দিয়ে করো।**

```python
STANDARD_COLUMNS = [
    "Manufacturer",      # Brand name (hardcoded, e.g. "Surya", "Gabby")
    "Source",            # Product page URL  ← আগে "Product URL" ছিল
    "Image URL",
    "Product Name",
    "SKU",
    "Product Family Id",
    "Description",
    "Weight",
    "Width",
    "Depth",
    "Diameter",
    "Length",
    "Height",
    "Seat Width",
    "Seat Depth",
    "Seat Height",
    "Arm Height",
    "Arm Width",
    "List Price",
]
```

**Column History:**
- আগে: `Product URL` → এখন: `Source`
- আগে: নেই → এখন: `Manufacturer` (brand name hardcoded)
- ZUO = সর্বশেষ updated vendor যেটায় নতুন format আছে
- Surya সহ পুরানো vendors এখনো `Product URL` ব্যবহার করছে → update দরকার

---

## SKILLS REFERENCE — কোন Skill কখন ব্যবহার করবো

### Skill 01 — Excel Read & Write
```python
# নতুন Excel তৈরি (Row 1=Brand, Row 2=URL, Row 3=Date, Row 4=Headers, Row 5+=Data)
wb, ws = create_output_excel(output_path, brand_name="Surya", source_url="https://surya.com")

# Step1 output পড়া
df = read_step1_excel("surya_Chandeliers.xlsx")

# Resume: আগে যেসব URL scrape হয়েছে সেগুলো skip করো
done = get_done_urls(output_path, url_col="Source")

# Auto-save প্রতি 10 rows
auto_save(wb, path, count, every=10)
```

### Skill 02 — Python Utilities
```python
# Fraction convert: "26 3/4" → "26.75"
val = frac_to_decimal("26 3/4 inches")

# Dimension text থেকে W/D/H/L/Dia/Weight extract
dims = extract_dims("26W x 14D x 32H")  # → {"Width":"26","Depth":"14","Height":"32",...}

# Retry decorator
@retry(times=3, delay=2)
def scrape_page(url): ...

# Random sleep (polite scraping)
polite_sleep(1.0, 3.0)
```

### Skill 03 — Web Scraping
```python
# Requests session with headers
session = get_session()
r = safe_get(session, url, retries=3)

# Selenium driver (stealth mode)
driver = get_driver(headless=True)

# Infinite scroll
scroll_to_bottom(driver, pause=1.5, max_rounds=20)

# Wait for element
el = wait_for(driver, "div.product-card", by="css", timeout=10)

# Query param pagination (generator)
for url, soup, page in paginate_query(base_url, param="page"):
    ...

# Image fallback (src → data-src → data-lazy-src → data-original)
img_url = get_image_url(img_tag)
```

### Skill 04 — Data Optimization
```python
# Dedup by URL
df = dedup_by_url(df, url_col="Source")

# Manufacturer column fill
df = fill_manufacturer(df, brand="Surya")

# Source column (maps from Product URL)
df = fill_source(df, url="https://surya.com")

# Column reorder (standard order)
df = reorder_columns(df, extra_cols=["Finish", "Base"])

# Summary report
print_summary(df, vendor="Surya")
```

### Skill 05 — Data Extraction
```python
# SKU extract from text
sku = extract_sku("Item #: ABC-123")

# Spec table parse (BS4 or plain text)
dims = parse_spec_table(soup_element)
# LABEL_MAP covers: w/width, d/depth, h/height, l/length,
#                   dia/diameter, wt/weight,
#                   sw/seat width, sd/seat depth, sh/seat height,
#                   ah/arm height, aw/arm width

# Shopify JSON extract from page HTML
data = extract_shopify_json(page_html)

# Description clean (stops at "dimensions", "specifications" etc.)
desc = extract_description(element, max_chars=800)

# Price parse: "$1,234.00" → "1234.00"
price = extract_price("$1,234.00")
```

### Skill 06 — Libraries
```python
# Standard imports for Step1/Step2 — copy করো
# Step1: STEP1_IMPORTS + STEP1_SELENIUM_IMPORTS বা STEP1_BS4_IMPORTS
# Step2: STEP2_IMPORTS
```

### Skill 07 — VPN / Anti-Bot
```python
# Vendor difficulty check
info = check_and_warn("Surya")   # → HIGH, chrome_debug_port + uc

# Chrome Debug Port launch (Surya Step1 style)
launch_debug_chrome(user_data_dir=r"C:\ChromeProfile\Surya")
driver = get_debug_driver(port=9222)

# Undetected Chrome (Surya Step2 style)
driver = get_uc_driver(headless=False, proxy=None)

# Block detection
if is_blocked(driver.page_source): ...

# Auto-select based on vendor
driver, dtype = auto_driver("Surya")

# DIFFICULT VENDORS:
# Surya → HIGH: chrome_debug_port + uc
# Loloi → HIGH: selenium_stealth
# Liaigre → HIGH: uc
# Visual Comfort → MEDIUM: selenium_stealth
# Holly Hunt → MEDIUM: selenium_stealth
# Janus et Cie → MEDIUM: selenium_stealth
```

---

## VENDOR DETAIL PAGE PATTERNS — Per-Vendor Analysis

প্রতিটা vendor-এর detail page কীভাবে কাজ করে — technique, selectors, special columns।

### TECHNIQUE GROUPS

#### GROUP A — Requests + BS4 (No JS needed)
```
Allan Knight    → div.product-desc-holder, strong.category (SKU)
Art & Forge     → window.wpdExtensionIsProductPage JSON (Shopify variant)
Artesia         → Wix tabs: data-hook="tab-item" (DESCRIPTION/SIZE/PRODUCT CARE)
Bennett         → div.product__description.rte, span.price-item--regular
Cameron         → div.product-description (simple)
Caracole        → accordion/details tag, "Dimensions in Inches: W X D X H"
Gabby           → div.specs-attributes-section (Shopify gabriellawhite.com)
Honning         → dimension line detection + regex W/D/H/L/DIA
Niermann Weeks  → div.col.grid_6_of_12
Pierre Frey     → div.specTechInfos (| separated blocks)
Stahl & Band    → requests+BS4
Johnathan Brown → requests+BS4
```

#### GROUP B — Selenium (JS-heavy, no bot protection)
```
Alliad Maker  → //li[h4/a[text()='DIMENSIONS']] XPath, .sku-spec h4
Bernhardt     → //panel[@panel-title='Dimensions'] + Fabrics panel (COM)
Bright Chair  → Two tabs: Finishes (.text-nav li[info-view="product-finishes-shell"])
                          then Specs (.text-nav li[info-view="product-details"])
Century       → //summary[.//h2[contains(text(),'Dimensions')]] accordion click
Chaddock      → tr#ctl00_ctl00_...trDimensionsOverall XPath table
Loloi         → div#details-drawer, dl/dt/dd definition lists
Sunpan        → //h2[contains(text(),'Dimensions')]/ancestor::summary accordion
                + table.tab-dimensions__table
ZUO           → Tab panels → Product Dimensions + Seat Dims (DIM_MAP + SEAT_MAP)
```

#### GROUP C — Selenium + UC/Stealth (Bot protection)
```
Surya         → UC + Chrome Debug Port (Surya_step2.py)
Liaigre       → UC
Holly Hunt    → Selenium stealth
Visual Comfort → Selenium stealth
Janus et Cie  → Selenium stealth
```

---

### DIMENSION EXTRACTION TECHNIQUES

#### 1. Regex on Description/Body Text (most common)
```python
# Pattern: "26W x 14D x 32H" or "26 3/4W x 14D"
re.search(r'(\d+\.?\d*)\s*"?\s*W\b', text)
re.search(r'(\d+\.?\d*)\s*"?\s*D\b', text)
re.search(r'(\d+\.?\d*)\s*"?\s*H\b', text)
re.search(r'(\d+\.?\d*)\s*"?\s*(DIA|DIAM|Diameter)', text, re.I)

# Pattern: "Dimensions in Inches: 46.0W X 12.0D X 72.0H"  (Caracole)
re.search(r'Dimensions\s*in\s*Inches\s*:\s*(.+)', text, re.I)

# Pattern: "W: 26" / "H: 32"  (Bernhardt style)
PAT_W = re.compile(r'W:\s*([\d.,]+)', flags=re.I)
PAT_H = re.compile(r'H:\s*([\d.,]+)', flags=re.I)
```

#### 2. Accordion Click → innerHTML parse (Century, Sunpan)
```python
# Click accordion button
summary = driver.find_element(By.XPATH, "//summary[.//h2[contains(text(),'Dimensions')]]")
driver.execute_script("arguments[0].click();", summary)
# Parse innerHTML <p> tags with regex
```

#### 3. XPath Table Row (Chaddock, Alliad Maker)
```python
li = driver.find_element(By.XPATH, "//li[h4/a[text()='DIMENSIONS']]")
dim_text = li.find_element(By.TAG_NAME, "p").text
```

#### 4. DL/DT/DD Definition Lists (Loloi, Chelsea Textiles)
```python
dls = container.find_elements(By.TAG_NAME, "dl")
for dl in dls:
    dt = dl.find_element(By.TAG_NAME, "dt").text  # label
    dd = dl.find_element(By.TAG_NAME, "dd").text  # value
```

#### 5. Shopify JSON Extract (Art & Forge)
```python
# Page source থেকে JSON parse
data = extract_balanced_json(html, "window.wpdExtensionIsProductPage")
desc_html = data.get("description", "")
# Variants থেকে color/size/price
for v in data.get("variants", []):
    color = v.get("option1")
    price = v.get("price")  # cents → divide by 100
```

#### 6. Panel/Tab System (Bernhardt, Bright Chair)
```python
# Bernhardt: panel[@panel-title='Dimensions']
panel = driver.find_element(By.XPATH, "//panel[@panel-title='Dimensions']")
# Bright Chair: tab click then spec extract
safe_click(driver, '.text-nav li[info-view="product-finishes-shell"]')
safe_click(driver, '.text-nav li[info-view="product-details"]')
```

#### 7. Wix Tab System (Artesia)
```python
tabs = {}
for tab in soup.select('div[data-hook="tab-item"]'):
    label = tab.select_one('span[data-hook="tab-item-label"]').text.upper()
    panel_id = tab.get("aria-controls")
    tabs[label] = panel_id
# tabs["SIZE"], tabs["DESCRIPTION"], tabs["PRODUCT CARE"]
```

---

### SPECIAL COLUMNS BY VENDOR TYPE

#### Lighting Vendors (Chandeliers, Sconces, Pendants, Lamps)
```
Wattage      → regex: "40 Max Wattage" / "60W max"
Socket       → "E-12 Socket", "E26 Socket", "G9 Socket"
Shade Details → "shade 20 x 21 x 11" / "silk shade 15t, 16b"
Lamping      → Alliad Maker specific
Base         → finish type
```

#### Seating Vendors (Chairs, Sofas, Stools)
```
Seat Height  → "sh" / "to seat" / "Seat Height:"
Seat Depth   → "sd" / "Seat Depth:"
Seat Width   → "sw" / "Seat Width:"
Arm Height   → "ah" / "Arm Height:"
Arm Width    → "aw" / "Arm Width:"
COM          → "COM-4.5 yards" / "Body Fabric: X yards"
COL          → leather yardage
```

#### Outdoor/Furniture General
```
Finish       → color/finish options
Base         → base material/finish
Materials    → material list
```

#### Rugs/Textiles (Loloi, Chelsea Textiles, Pierre Frey)
```
Construction     → "Hand Tufted", "Machine Made"
Material/Content → "Wool 100%", "Polyester"
Pile Height      → "0.125 inches"
Backing          → "Cotton"
Country of Origin→ "India"
Pattern Repeat   → fabric patterns
Size             → "2x3", "5x8" etc.
```

---

### ZUO FORMAT — NEW STANDARD REFERENCE
```python
# ZUO = সর্বশেষ updated vendor, নতুন column format এর reference
FIXED_COLUMNS = [
    "Manufacturer", "Source", "Image URL", "Product Name",
    "SKU", "Product Family Id", "Description", "Weight",
    "Width", "Depth", "Diameter", "Length", "Height",
]
# Dynamic seat dims (যখন product-এ আছে):
SEAT_MAP = {"SW": "Seat Width", "SD": "Seat Depth", "SH": "Seat Height",
            "AH": "Arm Height", "AW": "Arm Width"}
```

---

### PRODUCT FAMILY ID EXTRACTION RULES
```python
# Rule 1: Name split on " - " (Gabby, Century)
"Blair Pull - Dark"  → "Blair Pull"
"Soho Bed - King"    → "Soho Bed"

# Rule 2: Name split on "," (Allan Knight)
"Belden Chair, Natural"  → "Belden Chair"

# Rule 3: SKU prefix before "-" (Loloi, Surya)
"ABI-01-SF"  → "ABI"   (SKU prefix)

# Rule 4: Full name if no separator (Cameron, simple)
"Hampton Chair"  → "Hampton Chair"

# Rule 5: Name split on "-" first occurrence (Bright Chair)
"Belden-01"  → "Belden"
```

---

### SKU — পাওয়া না গেলে কীভাবে Generate করবো

#### STEP 1 — আগে খোঁজার চেষ্টা করো (Priority Order)

```python
# Priority 1: CSS selector দিয়ে page থেকে
sku = el.find_element(By.CSS_SELECTOR, "[data-test-selector='productStyleId']").text
sku = soup.select_one("strong.category").get_text(strip=True)
sku = driver.find_element(By.CSS_SELECTOR, ".sku-spec h4").text

# Priority 2: img alt থেকে (Surya pattern)
sku = img.get_attribute("alt").strip()

# Priority 3: URL slug থেকে (last part)
sku = url.rstrip("/").split("/")[-1]  # "https://surya.com/Product/BLND-001" → "BLND-001"

# Priority 4: JSON/API থেকে (Shopify)
sku = variant.get("sku", "")
```

#### STEP 2 — কোনোভাবেই না পেলে → Generate করো

**Standard Format:**
```
VENDOR_CODE (3 letters) + CATEGORY_CODE (2 letters) + INDEX (3 digits, zero-padded)

Example:
  Vendor = "Honning"  → HON
  Category = "Mirrors" → MI
  Index = 1           → 001
  SKU = "HONMI001"
```

**Generation Function (সব vendor-এ এটাই ব্যবহার করো):**
```python
def make_sku(vendor_name: str, category: str, index: int) -> str:
    vendor_code   = re.sub(r"[^A-Za-z]", "", vendor_name).upper()[:3].ljust(3, "X")
    category_code = re.sub(r"[^A-Za-z]", "", category).upper()[:2].ljust(2, "X")
    return f"{vendor_code}{category_code}{str(index).zfill(3)}"

# Examples:
make_sku("Honning",  "Mirrors",        1)  → "HONMI001"
make_sku("McLean",   "Lighting",       5)  → "MCLLI005"
make_sku("Cameron",  "Ottomans",       2)  → "CAMOT002"
make_sku("Liaigre",  "Objects",        3)  → "LIAOB003"
make_sku("Stahl & Band", "Pulls",      7)  → "STAPU007"
make_sku("Siemon & Salazar", "Chairs", 10) → "SIECH010"
```

#### Vendor-Specific SKU Patterns (আগে যেভাবে করেছি)

| Vendor | Pattern | Example |
|--------|---------|---------|
| Honning | `HON + category[:2] + idx` | HONMI001 |
| McLean | `MCL + category_url_slug[:2] + idx` | MCLLI005 |
| Cameron | `CAM + url_path_code + idx.zfill(2)` | CAMOT01 |
| Liaigre | `LIA + CategoryCode + idx.zfill(3)` | LIAOb001 |
| Stahl & Band | `STA-PU-idx.zfill(3)` | STA-PU-007 |
| Siemon & Salazar | `name[0][:3] + category_code + idx.zfill(3)` | SIECH010 |
| Art & Forge | `vendor[:3] + category[:2] + counter` | ARTPU5 |
| Surya | URL slug → img alt → MISSING | BLND-001 |

#### Clean SKU (website থেকে পেলে normalize করো)
```python
def clean_sku(raw: str) -> str:
    t = raw.strip()
    t = re.sub(r"\s+", "", t)  # spaces remove
    m = re.search(r"([A-Za-z0-9]+(?:[-_][A-Za-z0-9]+)+)", t)  # "ABC-001" pattern
    if m:
        t = m.group(1)
    t = re.sub(r"[-_]?p$", "", t, flags=re.I)  # trailing 'p' remove
    return t
# "BLND-001-P" → "BLND-001"
# "blnd001p"   → "BLND-001" (if pattern found)
```

#### Decision Rule
```
Website-এ SKU দেখা যাচ্ছে?
  হ্যাঁ → scrape করো + clean_sku() দিয়ে normalize করো
  না →
    URL slug কি meaningful? (like "abc-123") → URL slug ব্যবহার করো
    না →
      generate করো → make_sku(vendor, category, index)
      NEVER "MISSING" রাখো final output-এ
```

---

### FRACTION → DECIMAL (Always include)
```python
fraction_map = {"¼": 0.25, "½": 0.5, "¾": 0.75}
# Also handle: "26 3/4" → 26.75
# Regex: r"(\d+)?\s*(\d+)/(\d+)"
```

---

## DECISION 8 — Website Access Methodology (Step-by-Step)

নতুন website আসলে এই order-এ চেষ্টা করো:

### STEP A — Simple Request দিয়ে চেষ্টা

```
WebFetch(url) চেষ্টা করো:

200 OK     → HTML পাওয়া গেছে → BS4 দিয়ে parse করো
403/429    → Bot protection → STEP B যাও
Redirect   → নতুন URL ধরে আবার চেষ্টা
Timeout    → Site slow বা JS-heavy → Selenium লাগবে
```

### STEP B — Bot Protection Level বোঝো

```
Error type দেখে সিদ্ধান্ত নাও:

403 Forbidden     → Bot detection (Cloudflare, custom)
429 Too Many      → Rate limiting → sleep বাড়াও
CAPTCHA page      → Undetected Chrome (UC) লাগবে
JS-only render    → Selenium লাগবে (API না থাকলে)
```

### STEP C — API আছে কিনা খোঁজো (সবার আগে)

```
Shopify:
  /products.json?limit=250                    → test করো
  /collections/{handle}/products.json         → test করো

WordPress/WooCommerce:
  /wp-json/wc/v3/products                     → test করো

Next.js / Nuxt:
  Page source এ __NEXT_DATA__ বা window.__NUXT__ দেখো
  Browser Network tab এ XHR/fetch calls দেখো
  /_next/data/ path এ JSON আছে কিনা দেখো

Custom:
  /api/products, /api/catalog, /graphql        → test করো
  Network tab এ background API calls দেখো
```

### STEP D — Scraping Method Final Decision

```
API পাওয়া গেলে:
  → requests + json.loads()     [সবচেয়ে fast ও stable]

Bot protection নেই + no JS render:
  → requests + BeautifulSoup    [lightweight, fast]

JS heavy + bot protection নেই:
  → Selenium (regular Chrome)   [slow কিন্তু কাজ করে]

Cloudflare / surya.com type:
  → Undetected Chrome (UC)      [stealth mode]

Captcha আসে:
  → UC + manual solve prompt    [user intervention]

Rule: Selenium শেষ option — API বা requests আগে চেষ্টা করো।
```

### STEP E — Pagination Type Detect

```
URL-এ ?page=2 কাজ করে?     → query_param
/page/2/ URL pattern?        → url_path
"Load More" button আছে?     → button_click
Scroll করলে নতুন আসে?      → infinite_scroll (surya এটা)
Shopify API?                 → api_page_param (?limit=250&page=N)
```

### STEP F — Data Location Decision

```
প্রতিটা column কোথায় পাবো সেটা decide করো:

Category Listing page থেকে:
  → SKU, Image URL, Product URL, Product Name (basic)

Product Detail page থেকে:
  → Dimensions (W/D/H/Dia/L), Weight, Price
  → Finish, Materials, Description, Specs

API JSON থেকে:
  → Fields map করো → column-এ বসাও

কোনোভাবেই পাওয়া না গেলে:
  → instruction.json এ "not_available" list-এ রাখো
```

### Quick Reference Table

| Site Type | API আছে? | Method | Pagination |
|-----------|----------|--------|------------|
| Shopify | হ্যাঁ (products.json) | requests+json | api_page_param |
| WooCommerce | হ্যাঁ (wp-json) | requests+json | query_param |
| Next.js | অনেক সময় | requests+XHR | varies |
| Custom (no bot) | না | requests+BS4 | check manually |
| Custom (bot) | না | Selenium/UC | check manually |
| surya.com type | না | UC + stealth | infinite_scroll |

---

## VENDOR NOTES — Gabby

**Site:** `gabby.com` → redirects to `gabriellawhite.com` (Shopify)
**Method:** Shopify API `/collections/{handle}/products.json` on gabriellawhite.com
**SKU / Price / Description:** Shopify API থেকে (Price = integer, no $)
**Dimensions / Features / Specs:** Product page `div.specs-attributes-section` (h3: DIMENSIONS / FEATURES / ASSEMBLY)
**Warranty / Care / More Info:** `div.accordion-item` → `div.pb-5`
**Specification Sheet:** `div[sub-section-id*="spec_sheet_link"] a`
**Pillow Size / Fill:** `input[name="properties[pillow_size]"]` / `input[name="properties[fill_material]"]`

### Dynamic Columns Rule (Gabby — apply to similar sites)
- **DIMENSIONS section:** শুধু 6টা fixed column → Weight, Width, Depth, Diameter, Length, Height
  বাকি সব key (Seat Height, Clearance, Drawer depth ইত্যাদি) → dynamic column
- **FEATURES section:** সব key-value pair → dynamic column (কোনো FEAT_MAP নেই)
- Dynamic columns per-sheet: প্রতিটা sheet-এ শুধু সেই columns আসবে যেখানে ওই sheet-এ কমপক্ষে একটা value আছে
- Duplicate prevention: `_FIXED_COLS = set(COLUMNS)` — কোনো dynamic key যদি fixed column-এর নামের সাথে মিলে যায়, সেটা fixed column-এ যাবে, extra-তে না

### Shared Collection Handle → product_type Filter Pattern
একই collection handle একাধিক category-তে ব্যবহার হলে → `product_type` substring filter দাও:
```python
# COLLECTIONS format: (name, [handles], dedup_key, product_type_filter)
("Chandeliers",   ["hanging-lighting"],      "local", "chandelier"),
("Pendants",      ["ceiling-lights"],        "local", "pendant"),
("Sconces",       ["wall-lights"],           "local", "sconce"),
("Flush Mount",   ["ceiling-lights"],        "local", "flush"),
("Table Lamps",   ["lamps"],                 "local", "table lamp"),
("Floor Lamps",   ["lamps"],                 "local", "floor lamp"),
("Pillows & Throws", ["decorative-accessories"], "local", "pillow"),
("Rugs",          ["decorative-accessories"], "local", "rug"),
```
Filter logic: `type_filter in p.get("product_type", "").lower()` (partial match, case-insensitive)

### Excel Header Format (per sheet)
```
Row 1: Brand: | Gabby
Row 2: Link:  | <category link>
Row 3: (empty)
Row 4: Column headers
Row 5+: Data
```
- Category link auto-generate: `f"{GABBY_BASE}/collections/{handles[0]}"`
- Client-specific override: `CAT_LINKS_OVERRIDE` dict দিয়ে override করো
  - Cabinets: `https://gabby.com/products/indoor-dining/cabinets, https://gabby.com/products/storage/sideboards`
  - Dressers & Chests: `https://gabby.com/products/bedroom/dressers, https://gabby.com/products/bedroom/chests`

### Price Format
`price_to_int()` → `$1,549.00` → `1549` (integer string, no $ sign)

**Code Update Rule:**
যখন user পুরানো code দিয়ে update বলে → শুধু requested change করো, বাকি সব same রাখো।

---

## DEMO_DATA VENDOR ANALYSIS (D:\Demo_Data — Client Updated/Refresh)

`D:\Demo_Data` = 16 vendors, client-refreshed versions. এই folder থেকে নতুন pattern শেখো।

### COLUMN FORMAT STATUS

| Vendor | URL Column | Manufacturer | Format Status |
|--------|-----------|--------------|---------------|
| Alfonso Marina | Product URL | No | OLD |
| Alliad Maker | Product URL | No | OLD |
| **Arteriors** | **Source** | **Yes** | **NEW ✅ REFERENCE** |
| Bernhardt | Product URL | Yes | MIXED (has Mfr, no Source) |
| Century | Product URL | No | OLD |
| Chaddock | Product URL | No | OLD |
| Curry & Company | Product URL | No | OLD |
| Fairfield | Product Url (lowercase) | No | VERY OLD (code uses Source, Excel doesn't) |
| Palecek | Product URL, `Field No.` instead of `Index` | No | OLD |
| Rejuvenation | Product URL | No | OLD |
| Theodore Alexander | Product URL, `Categori` typo, `Product Family Name` | No | OLD |
| The Future Perfect | Product URL | No | OLD |
| Visual Comfort | Product URL | No | OLD (no Index/Category in Step1) |

**Rule:** নতুন vendor কোড লেখার সময় Arteriors format follow করো (Source + Manufacturer).
পুরানো vendor update করতে বললে শুধু column rename করো, logic same রাখো।

---

### VENDOR-BY-VENDOR SCRAPING PATTERNS

#### 1. Alfonso Marina — WordPress/Elementor Custom CMS
- **Site:** alfonsomarina.com
- **Step1:** Selenium + infinite scroll (scroll until same count for 3 rounds)
- **Card selector:** `div.registroProducto`
- **URL:** `a.registroImagen[href]` | **Name:** `a.registroTitulo` | **Image:** `pick_real_image()` (checks data-src, data-lazy-src, srcset, src — skips 1px.png)
- **Pagination:** infinite scroll only (no page param)
- **Step2:** Selenium, dim prefix text (`W:`, `D:`, `H:`, `Dia:`, `Weight:`) inside `div.elementor-widget-container`
- **SKU:** `re.search("PRODUCT CODE", string)` in page text
- **Description:** `h2.elementor-heading-title` containing "DETAILS" → next section text
- **Finish:** `div.elementor-widget-text-editor` containing "As Shown:"
- **Product Family Id:** = full Product Name
- **Fraction:** `convert_fraction_to_decimal()` → handles `26 3/4` → `26.75`

#### 2. Alliad Maker — NetSuite Commerce
- **Site:** alliedmaker.com
- **Step1:** Selenium + `?page=N` (break when empty)
- **Card selector:** `div.facets-item-cell-list`
- **Name:** `.facets-item-cell-list-name span[itemprop='name']`
- **URL:** `a.facets-item-cell-list-name[href]` or `a.facets-item-cell-list-anchor[href]`
- **Image:** `img.facets-item-cell-list-image[src]` (split "?" to remove size params)
- **Skip banners:** `if "Commerce-category-banners" in image_url: continue`
- **Step2:** Selenium, XPath accordions:
  - Dims: `//li[h4/a[text()='DIMENSIONS']]` → `p` tag → regex parse
  - Weight: `//li[h4/a[text()='WEIGHT']]`
  - Lamping: `//li[h4/a[text()='LAMPING']]`
- **dim regex:** `r'(\d+\.?\d*)\s*"?\s*(DIA|W|D|H|L)\b'`
- **Extra column:** `Lamping` (lighting-specific)
- **Product Family Id:** = Product Name

#### 3. Arteriors — React/Next.js + Alpine.js (NEW FORMAT REFERENCE)
- **Site:** arteriorshome.com
- **Step1:** Selenium + Load More click loop
  - Selector strategies (tries all): `div[id^='slide-']`, `li.product-item`, `div.product-item-info`, `[data-product-id]`
  - SKU from URL: `re.search(r'([A-Za-z]{2,5}\d{1,5}-\d{2,5})$', slug).upper()`
  - Image: tries `source[type='image/webp']` srcset first, then `img[src]`
  - Signal click: `save_and_exit()` on CTRL+C (partial save)
- **Step2:** Alpine.js accordion sections:
  1. `click_accordion("Appearance & Dimensions")` → `parse_label_value_rows()`
  2. `click_accordion("Wiring & Bulbs")` → Voltage, Socket Type, etc.
  3. `click_accordion("Compliance & Certification")` → dynamic columns
  4. `click_accordion("Technical Documents")` → Tearsheet Link, Specsheet Link, etc.
  5. `click_accordion("Returns Policy")`, `click_accordion("Warranty")`
- **Accordion find:** `//span[contains(@class,'text-primary-500') and normalize-space(text())='Label']/ancestor::div[contains(@class,'cursor-pointer')]`
- **Content parse:** `div.flex.justify-between` rows → label `span.text-secondary-400`, value `span.text-primary-500`
- **Dynamic cols:** `FIXED_COL_ORDER` first, then extra columns appended alphabetically
- **Product Family ID:** `re.split(r"[-_.]", name, maxsplit=1)[0]`
- **Weight:** strip "lbs" → numeric only
- **Dim parse:** `parse_dimension_string()` → `r"\b(Dia|W|D|L|H):\s*([\d.]+)\s*in"` pattern
- **OVERRIDE_COLS:** always overwrite Description, Finish, Width, Depth, Height, etc.
- **Column order:** Manufacturer, Source, Image URL, Product Name, SKU, Product Family ID, Description, Weight, Width, Depth, Diameter, Length, Height, [extra...]

#### 4. Bernhardt — Angular SPA with Hash Routing
- **Site:** bernhardt.com
- **Step1:** Selenium + URL format string `?page={page}`
- **Card:** `div.grid-item`, **Name:** `div.product-header`, **Image:** `img.grid-image[src|data-src]`
- **SKU:** `span.product-id` + `div.meta-component.ng-binding` joined with ` | `
- **Pagination:** insert `{page}` in URL template, break when `div.grid-item` empty
- **Wait:** `time.sleep(8)` — Angular takes long to render

#### 5. Century — Shopify (Requests only, NO Selenium for Step1)
- **Site:** shop.centuryfurniture.com (Shopify)
- **Step1:** `requests.Session()` + BS4, NO Selenium needed
- **Card selector:** `div.card-wrapper` → `a[id*='CardLink']`
- **Image:** `img[srcset]` → split ",", take [0], strip `&width=...`
- **Pagination:** `?page=N` query param, break when `div.card-wrapper` empty
- **Step2:** Selenium accordion: `//summary[.//h2[contains(text(),'Dimensions')]]`
  - Click → get `div.accordion__content` → parse `<p>` tags
  - Regex: `WEIGHT:`, `HEIGHT:`, `WIDTH:`, `DEPTH:`, `DIAMETER:`, `Seat Height:`, `Arm Height:`
- **SKU:** `div.hideAll span.sku`
- **Description:** `div.product__description.rte` (filter out dimension lines)
- **Finish:** from same description div, look for `<br>Finish:` pattern
- **Price:** `div.price.price--large`, strip "MSRP" and "USD"
- **Extra columns:** `Com`, `Finish`, `Seat Height`, `Arm Height`

#### 6. Chaddock — ASP.NET WebForms CMS
- **Site:** chaddock.com/styles
- **Step1:** Selenium + View All button + lazy scroll
  - Dropdown: `ctl00_ctl00_..._ddlPageSize` set to 2000000000, dispatch change event
  - Button: `ctl00_ctl00_..._btnViewAll` — JS click (handles GA banner)
  - Lazy scroll: loop until same count for `NO_GROWTH_LIMIT=4` rounds
- **Card:** `div.grid_3.SearchResults_Container`
- **Name + SKU:** `span.CHAD_SearchResult_Title` → `.text.split("\n")` → [0]=name, [1]=SKU
- **Step2:** Specific hardcoded CSS IDs:
  - `tr#...trDimensionsOverall td:nth-of-type(2)` → parse W/D/H from text
  - `tr#...trDimensionsSeat td:nth-of-type(2)` → Seat H/D/W
  - `tr#...trDimensionsArmHeight td:nth-of-type(2)` → Arm Height
  - `tr#...trDiameter td:nth-of-type(2)` → Diameter
  - `tr#...trWeight td:nth-of-type(2)` → Weight
  - `tr#...trComFabric td:nth-of-type(2)` → COM
- **Product Family Id:** `product_name.split(" - ")[0]`

#### 7. Curry & Company — React + Material UI (MUI) Accordions
- **Site:** curreyandcompany.com
- **Step1:** Selenium + `?page=N` (if URL has `?` → `&page=N` else `page=N`)
- **Card:** `div.relative.group`
- **SKU:** `div.paragraph-3b-sm` (from listing page — rare to get SKU on listing)
- **Step2:** MUI accordion pattern:
  - Find: `//*[contains(@class,'MuiAccordionSummary-root')][.//*[contains(translate(...),'keyword')]]`
  - Open: click if not `Mui-expanded`
  - Content: `div.MuiAccordionDetails-root`
  - Keywords tried: `"dimensions"`, `"dimensions & weight"`, `"dimension"`, `"size"`
- **Dim extraction:** `extract_value_block(full_text, "Overall")` → parse `r'([\d\.]+)\s*"\s*([a-z\.]+)'`
- **Extra:** Shade Details (Shade Top/Bottom/Height → 3 numbers), Canopy, Seat H/W/D, Arm H/W/L
- **Specs:** separate accordion → Finish, Color Temperature, Socket Type, Wattage

#### 8. Fairfield — Custom React CMS
- **Site:** fairfieldchair.com
- **Step1:** Selenium + `?page=N` format string in URL
- **Card:** `div.product-widget, div.product-card`
- **Name:** `h4.product-title` or `h3.product-title`
- **SKU:** `h5.product-sku` or `span.sku`
- **Image trick:** `currentSrc || src` via JS, naturalWidth check (`> 10`) for lazy load
- **⚠️ Code uses `Source` column but Excel has `Product Url` (lowercase)** — mismatch needs fix
- **Category URL template:** single category per script run (not multi-category dict)

#### 9. Palecek — Custom .aspx CMS
- **Site:** palecek.com/itembrowser.aspx
- **Step1:** Selenium + `viewall=true` in URL + single scroll to bottom
- **Card:** `div.ProductThumbnailSection div.ProductThumbnail`
- **URL:** `a[href*='iteminformation.aspx']`
- **Image:** `img.ProductThumbnailImg[data-src]` → normalize to https
- **Name:** `p.ProductThumbnailParagraphDescription a`
- **SKU:** `p.ProductThumbnailParagraphSkuName a` → fallback `h3`
- **Backfill:** if name/SKU missing → open detail page in new tab to get it
- **⚠️ Uses `Field No.` instead of `Index`** — unique column name

#### 10. Rejuvenation — Williams-Sonoma Custom
- **Site:** rejuvenation.com
- **Step1:** MANUAL MODE — `opts.add_experimental_option("detach", True)` Chrome stays open
  - User scrolls manually, then `input("press Enter")` to extract
- **Card:** `[data-component='Shop-GridItem'], .grid-item`
- **Step2:** Complex variation system:
  - Radio inputs: `ul[data-test-id="product-attributes"] input[type="radio"]`
  - Click each variation → different SKU/image/finish per row
  - SKU: `[data-test-id="sku-display"]` → regex `SKU\s*:\s*([A-Za-z0-9\-_]+)`
  - `sku_fallback()` if not found: `VEN + CAT + INDEX`
- **Dimension extraction:** accordion accordion-panel + fallback XPath + body text scan
  - KV pattern: `r'\b{key}\b\s*:\s*({num})'`
  - Fraction: `_to_float_str()` → `"34-1/4"` → `"34.25"`, `"17 3/4"` → `"17.75"`
- **Extra:** Shade Details, Wattage (max + efficiency), Socket Type, Canopy, Length, Arm Depth/Width

#### 11. Theodore Alexander — Custom Angular-like CMS
- **Site:** theodorealexander.com
- **Step1:** Selenium + manual `total_pages` variable (not auto-detect)
- **Card:** `div.info`
- **URL:** `a.productImage[href]` (prepend domain)
- **Image:** `img[src]` inside `a.productImage`
- **Name:** `div.name a[title]` attribute
- **SKU:** `div.sku` text
- **⚠️ Uses `Categori` (typo), `Product Family Name` (not `Product Family Id`)** 

#### 12. The Future Perfect — WooCommerce
- **Site:** thefutureperfect.com
- **Step1:** Selenium + infinite scroll (height comparison)
- **Card:** `li.product`
- **Price extraction:** `re.search(r'[€£$]\s*[\d,]+(?:\.\d+)?', text)` → clean `$7,000`
- **Output:** single Excel per category, not multi-sheet workbook

#### 13. Visual Comfort — Magento-style
- **Site:** visualcomfort.com
- **Step1:** Selenium + manual `START_PAGE`/`END_PAGE` config
- **Card:** `li.product-card` (multiple selector fallbacks)
- **Pagination:** `?p=N` param
- **⚠️ No `Index` or `Category` column in Step1 output** — raw flat list
- **Three scripts:** listpage.py (Step1), detailspage.py (Step2), step3.py (merge/format)
- **Special columns:** Socket, Wattage, Chain Length, Rating, Spec Sheet, Install Guide, CAD Block, 3D Rendering

---

### PAGINATION METHODS SUMMARY (all vendors)

| Method | Vendors | Code Pattern |
|--------|---------|-------------|
| `?page=N` | Alliad Maker, Bernhardt, Curry, Century, Fairfield, Theodore Alexander | `url + f"&page={page}"` or format string |
| `?p=N` | Visual Comfort | Magento style |
| Infinite scroll | Alfonso Marina, The Future Perfect | scroll until height/count stable |
| Load More button | Arteriors | `get_load_more_button()` + JS click loop |
| View All button | Chaddock | dropdown set + button click |
| `viewall=true` in URL | Palecek | single URL param |
| Manual scroll | Rejuvenation | `input()` prompt, user scrolls |

---

### DIMENSION EXTRACTION METHODS SUMMARY

| Method | Vendors | Pattern |
|--------|---------|---------|
| Elementor div prefix text | Alfonso Marina | `text.startswith("W:")` |
| XPath accordion `//li[h4/a[text()='LABEL']]` | Alliad Maker | accordion tab by exact label |
| Alpine.js accordion click + label-value rows | Arteriors | `click_accordion()` → `parse_label_value_rows()` |
| HTML `<details><summary>` accordion | Century | `//summary[.//h2[contains(text(),'Dimensions')]]` |
| ASP.NET hardcoded CSS IDs | Chaddock | `tr#ctl00_...trDimensionsOverall td:nth-of-type(2)` |
| MUI accordion (multiple keyword fallback) | Curry | `MuiAccordionSummary-root` + open/close toggle |
| KV text regex | Rejuvenation | `r'\b{key}\b\s*:\s*([\d.]+)'` on full text |
| Manual page config | Visual Comfort | detailspage.py separate script |

---

### ACCORDION CLICK PATTERNS (Quick Reference)

```python
# MUI (Curry & Company style)
root = driver.find_element(By.XPATH,
    "//*[contains(@class,'MuiAccordionSummary-root')]"
    "[.//*[contains(translate(normalize-space(.),'ABCDE...','abcde...'),'dimensions')]]"
)
if "Mui-expanded" not in root.find_element(By.XPATH,"./..").get_attribute("class"):
    driver.execute_script("arguments[0].click();", root)

# HTML details/summary (Century style)
summary = driver.find_element(By.XPATH, "//summary[.//h2[contains(text(),'Dimensions')]]")
driver.execute_script("arguments[0].click();", summary)
content = summary.find_element(By.XPATH, "./following-sibling::div[contains(@class,'accordion__content')]")

# Alpine.js (Arteriors style)
xpath = f"//span[contains(@class,'text-primary-500') and normalize-space(text())='{label}']/ancestor::div[contains(@class,'cursor-pointer')]"
el = driver.find_element(By.XPATH, xpath)
driver.execute_script("arguments[0].click();", el)

# ASP.NET hardcoded ID (Chaddock style)
elem = driver.find_element(By.CSS_SELECTOR,
    "tr#ctl00_ctl00_ChildBodyContent_PageContent_trDimensionsOverall td:nth-of-type(2)")
```

---

### IMAGE RETRIEVAL PATTERNS (Quick Reference)

```python
# Lazy image with currentSrc check (Fairfield)
src = driver.execute_script("return arguments[0].currentSrc || arguments[0].src || ''", img)
natw = driver.execute_script("return arguments[0].naturalWidth || 0;", img)
if src and natw > 10: img_url = src

# data-src priority (Palecek, Alfonso Marina)
img_url = img.get("data-src") or img.get("data-lazy-src") or img.get("src")

# srcset → take last/best (Alfonso Marina)
if "," in val: return val.split(",")[-1].split(" ")[0]

# og:image meta (Rejuvenation)
og = driver.find_elements(By.CSS_SELECTOR, 'meta[property="og:image"]')
if og: img_url = og[0].get_attribute("content")
```


---

## DEMO_DATA VENDOR ANALYSIS - PART 2 (Remaining Vendors)
*Added: 2026-04-22 | Vendors: Gabby, Kravet, Bernhardt, The Future Perfect, Visual Comfort (full), Theodore Alexander*

---

### GABBY (gabby.com)

**Step1:**
- Method: Requests + BS4 -- NO Selenium needed
- Card: li.group/product-card
- Name: h3 > a (href = relative URL, prefix BASE_URL)
- Image fallback: data-src -> src -> srcset.split(" ")[0]; prefix https: if starts with //
- Pagination: page 1 = raw URL; page 2+ = ?page=N or &page=N depending on ? in url
- Stop: if not items: break
- Output columns: Source, Image URL, Product Name (NEW FORMAT, per-category files)
- File: Code/Gabby_step1.py

**Step2:**
- Method: requests.get() + BS4 (still no Selenium)
- SKU: p[id^=Sku-template] -> strip "SKU:"
- Description: div.inline-richtext
- Dimensions: div.specs-attributes-section -> h3 "DIMENSIONS" -> span pairs (label, value)
- Features (Finish/Color): div.specs-attributes-section -> h3 "FEATURES"
- Price: 5 fallback patterns (body-l strong, Shopify JSON cents/100, data-product-price, price classes, Our Price text)
- Product Family Id: re.sub(r" - .*", "", product_name) -- strip after " - "
- Resume: skip rows where SKU already filled; Batch save: every 5 rows
- Extra columns: Price, Dimensions, Features, More Information, Warranty, Care Instructions, Assembly Required, Specification Sheet, Pillow Size, Material
- File: Code/Gabby_step2.py

---

### KRAVET (kravet.com) - 4-STEP PIPELINE

**Architecture:**
- step1 -> step2 -> step3 (OR step4 = combined step2+step3)
- Per-category files: kravet_[Category]_step1.xlsx -> kravet_[Category]_final.xlsx

**Step1:**
- Method: Selenium + infinite scroll (single URL per category, no page param)
- Card selectors: ol.ais-Hits-list li.ais-Hits-item, ol.ais-InfiniteHits-list li.ais-InfiniteHits-item, li.product-item, div.product-item
- Name: strong.product.name.product-item-name a (multiple fallbacks)
- SKU: span.product-item-sku, .product-item-sku, span.sku, [data-sku]
- Image: img -> src -> data-src -> data-original -> data-lazy -> data-ll-src
- Cookie banner: button#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll
- Output: Product URL, Image URL, Product Name, SKU, Page URL (OLD FORMAT)

**Step2 (detail scraping):**
- Method: Selenium, time.sleep(4) for page render
- Details: #details table#product-attribute-specs-table tbody tr -> th (key) + td (val) -> each row = its OWN column (dynamic!)
- Skip hidden rows: if not tr.is_displayed(): continue
- Resources tab: click [aria-labelledby=tab-label-resources] -> PDF links -> Product Info column
- Description: div.product-description-container

**Step3 (normalization, pure pandas -- no Selenium):**
- Parses Details text (comma-separated Key: Value strings) via extract_key()
- keep_inches_only(): finds first inch value (12", 12 1/2", 3/4"), converts fractions
- keep_pounds_only(): finds lb/lbs value, converts kg->lb if needed
- normalize_base(): Candelabra (E12) or Medium (E26/E27)
- Product Family Id: split at first comma, hyphen, underscore, or & in name

**Step4 (combined step2+step3 in one script):**
- Same scraping as step2 (each detail row -> own column)
- Same normalization as step3 via dcol() helper (checks multiple column aliases)
- Final output: Product URL, Image URL, Product Name, Product Family Id, SKU, Description,
  Width, Depth, Diameter, Height, Length, Weight, Finish, Color, Socket, Wattage, Lightsource,
  Color Temperature, Extension, Rating, Shade Details, Base, Canopy, Chain Length,
  Seat Number, Base/Foot Type, COM Available, COM, COL, COT, Arm Height, Seat Height, Seat Depth, Cushion, Product Info
- Files: Code/step1.py, step2.py, step3.py, step4.py

---

### BERNHARDT (bernhardtfurniture.com) - ANGULAR SPA

**Step2:**
- Method: Selenium (headless=False)
- Input column: Product URL (OLD FORMAT)
- Dimensions panel XPath: //panel[@panel-title='Dimensions']//button -> click -> ul.no-bullets
  - First li = label (uppercased), remaining li with " in" = value
  - Maps: SEAT WIDTH/DEPTH/HEIGHT, ARM WIDTH/DEPTH/HEIGHT, WIDTH, DEPTH, HEIGHT, DIAMETER
- COM: //panel[@panel-title='Fabrics']//ul -> "Body Fabric:" text
- Description: //panel[@panel-title='Description']//div[contains(@class,'panel-body')]
- Image: 10 CSS selectors tried + meta og:image + link[rel=image_src] + JSON-LD
- Finish: div.items.with-images div.text-center.ng-binding -> list joined by ", "
- Weight: //div[contains(text(),'Weight')]/following-sibling::div
- Product Family ID: full product name from h1.product-name, div.product-title, div.one-up-title
- Batch save: every 5 rows
- Output: Product URL, Image URL, Product Name, SKU, Product Family ID, Description, Weight,
  Width, Depth, Diameter, Height, Finish, Seat Width/Depth/Height, Arm Width/Depth/Height, COM
- File: Code/Berhardt_step2.py

---

### THE FUTURE PERFECT (thefutureperfect.com) - WOOCOMMERCE

**Step2:**
- Method: Selenium
- Input: Product URL, Image URL, Product Name, List Price
- SKU extraction 3 strategies:
  1. #nitro-telemetry-meta innerHTML -> JSON "sku":"..."
  2. JSON-LD application/ld+json -> obj["sku"] or obj["offers"]["sku"]
  3. Any element text matching r'SKU[:\s#-]*([A-Za-z0-9._-]+)'
- Description: div.description p -> div.product-description p
- Specs: section.border.technical[data-script='ProductTechSpecs'] div.spec -> h6+p joined by " | "
- Dimension value: from spec where label contains "dimension"
- Weight: regex on specs text r'Weight[^|]*' -> first number
- Dimension parser: L/W/D/Dia/H patterns in dimension text
- Seat/Arm: regex r'Seat Height[:\s]*([\d.]+)' etc. on combined text
- COM/COL: regex on specs text
- Canopy: r'Canopy[^|]*' -> first number before quote char
- Shade Details: r'Shade\s*:?\s*([^|]+)' -> strip "Shade:" prefix
- Materials: r'(Materials[^|]+)' -> strip "Materials:" prefix
- Product Family Id = Product Name (full name, not split)
- Output via openpyxl (not pandas):
  Product URL, Image URL, Product Name, Product Family Id, List Price, SKU, Description,
  Weight, Specifications, Materials, Dimensions, Length, Width, Depth, Diameter, Height,
  Seat Height/Depth/Length, Arm Height/Width/Depth, COM, COL, Base, Canopy, Shade Details
- File: Code/The_Future-Perfect2.py

---

### VISUAL COMFORT (visualcomfort.com) - MAGENTO, 3-SCRIPT

**detailspage.py (OLDER step2):**
- Input: Product URL (OLD FORMAT)
- Variation carousel: .product-item-variation-carousel-wrapper a.configurable-thumbnail
  - data-product-sku or data-clp-sku attributes
  - Carousel scroll: button.owl-next (up to 20 attempts)
- Specs OLD layout: #product-specifications-bottom table.options tbody tr -> .label + .pure-value
- Specs NEW layout: #spec-inch-tab table.product-attribute-specs-table tbody tr td
- Finish + Color Temperature: select.super-attribute-select options
- Image: .fotorama__stage__frame.fotorama__active img + 4 fallbacks
- Creates base row + 1 row per variation (multiple rows per product)
- Resume: by Input Index

**step3.py (NEWER FIXED version) - IMPORTANT QUIRKS:**
- Column is "Manufacture" NOT "Manufacturer" -- HAS TYPO (known issue)
- Uses Source column (NEW FORMAT)
- Manufacture = "Visual Comfort" hardcoded, re-enforced at end of build_row()
- Image 5 strategies:
  1. img.dropin-image, img.dropin-image--loaded
  2. Scene7 CDN: images.visualcomfort.com/is/image/ (KEEP query string)
  3. SKU-like alt text pattern
  4. Product keyword in alt (chandelier/sconce/pendant etc.)
  5. Any VC image that passes is_vc_product_image()
- Specs A: .additional-product-information-item (label + value divs)
- Specs B: .specifications-list .specifications-item (colon-separated text)
- Dimension string: r'(Dia|L|W|H|D)\s*:\s*([\d.]+)' -> Width/Height/Length/Depth/Diameter
- Variation: swatch [data-sku] -> select options -> JSON scripts
- Tech resources: Spec Sheet, Install Guide, CAD Block, 3D Rendering, SketchUp, Revit
  (from /media/docs/ links or keyword matching)
- Atomic saves: _checkpoint.xlsx -> os.replace()
- Resume: by Source URL
- Output: Manufacture, Source, Image URL, Product Name, SKU, Product Family Id, Price,
  Description, Weight, Width, Depth, Diameter, Length, Height + fan-specific + tech resource cols
- File: Code/step3.py

---

### THEODORE ALEXANDER (theodorealexander.com)

**Step2:**
- Method: Requests + BS4 (NO Selenium for step2, despite step1 using Selenium!)
- Input: Product URL (OLD FORMAT)
- QUIRK: Column names are Width (in), Depth (in), Diameter (in), Height (in) -- has "(in)" suffix!
- Dimensions: table.tableDimension -> header th = column names -> row where th="in" -> cell values
- Description: div.product_detail_info_description > p -> div.product_desc -> #nav-detail .col-xl-8 p
- Finish: div.col-xl-8.col-md-12.w-100 li -> label with "finish" -> span.col-xl-8
- Weight (Gross only): .product_tab_content_detail-title with "gross weight" -> first numeric span
- Price: .price, [data-price] -> regex for currency symbol + digit
- Details tab: div#nav-detail div.row.p-0.m-0.w-100 -> label/value pairs for:
  Collection, Room/Type, Main Materials, Shapes Materials, Net Weight,
  Seat Height, Arm Height, Inside Seat Depth, Inside Seat Width
- QUIRK: Uses "Product Family Name" (NOT "Product Family Id"!)
- Progress: tqdm library, autosave every 25 rows
- Output: Product URL, Image URL, Product Name, SKU, Product Family Name, Description,
  List Price, Weight, Net Weight, Width (in), Depth (in), Diameter (in), Height (in),
  Finish, Collection, Room/Type, Main Materials, Shapes Materials,
  Seat Height, Arm Height, Inside Seat Depth, Inside Seat Width
- File: Code/Theodone_alexander_step_2.py

---

### EICHHOLTZ
- Only .idea folder exists -- no code, no data found.

---

### COMPLETE DEMO_DATA COLUMN FORMAT SUMMARY

| Vendor | URL Column | Has Manufacturer | Format |
|--------|-----------|-----------------|--------|
| Alfonso Marina | Product URL | No | OLD |
| Alliad Maker | Product URL | No | OLD |
| Arteriors | Source | Yes (Manufacturer) | NEW |
| Bernhardt | Product URL | No | OLD |
| Century | Product URL | No | OLD |
| Chaddock | Product URL | No | OLD |
| Curry & Company | Product URL | No | OLD |
| Eichholtz | No code | - | - |
| Fairfield | Product Url (lowercase!) | No | VERY OLD |
| Gabby | Source | No | MIXED |
| Kravet | Product URL | No | OLD |
| Palecek | Product URL | No | OLD |
| Rejuvenation | Product URL | No | OLD |
| The Future Perfect | Product URL | No | OLD |
| Theodore Alexander | Product URL | No | OLD |
| Visual Comfort | Source + Manufacture (typo!) | Yes (typo) | MIXED |

Target NEW format: Manufacturer, Source, Image URL, Product Name, SKU, Product Family Id, ...


---

## PHASE-05 (AU) VENDOR SCRAPING PATTERNS — PART 2
*Added: 2026-04-22 | Vendors: Loloi, McLean, Mr. Brown London, New Classics, Niermann Weeks,
Palmer Hargrave, Pierre Frey, Powell & Bonnell, Quatrine, Remains, Rowe, Siemon & Salazar,
Stahl & Band, Studio Twenty Seven, Sudio Bel Vetro, Sunpan, Surya, Sutherland, Troscan,
Utter Most, Verellen, Vill & House, Wells Abbott, Worlds Away, Zuo,
+ Invisible Collection, Bennett, CR Laine, Highland House, Johnathan Brown Inc*

---

### LOLOI (loloirugs.com) — Selenium Load More

- **Method:** Selenium, single list URL, load more button
- **Card:** `div.product-card.relative.card-block__item`
- **URL+Name:** `a.js-product-name.product-card__name` → `title` attr (full name), `span` inside → SKU
- **Image:** `img.product-card__image` currentSrc, fallback `picture source` srcset → `parse_srcset_for_best()` (largest width descriptor)
- **Pagination:** `button#load-more` click loop, stop when count stable or >= EXPECTED_TOTAL (3234)
- **Dedup:** by Product URL
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Loloi (Scraping & Spot Checking)/Code/step1.py

---

### McLEAN LIGHTING (mcleanlighting.com) — Requests+BS4

- **Method:** Requests+BS4, no Selenium needed
- **Card:** `li.product`
- **URL:** `a[href]`, **Image:** `img[src]`, **Name:** `h3`
- **SKU:** Generated — `MCL + category_slug[:2].upper() + index`
- **Pagination:** None (single page per URL)
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** McLean Lighting Work/Code/Step1.py

---

### MR. BROWN LONDON (mrbrownhome.com) — Selenium WooCommerce

- **Method:** Selenium + WebdriverManager
- **Card:** `ul.products.columns-4 > li.product.type-product`
- **URL:** `a.woocommerce-LoopProduct-link[href]`
- **Name:** `h2.woocommerce-loop-product__title`
- **Image:** `img[src]`, fallback `srcset.split(",")[0].split(" ")[0]`
- **SKU:** Generated — `VEN[:3] + CAT[:2] + index` (using VENDOR_NAME = "Mr Brown Home")
- **Pagination:** `base_url + "page/{N}/"`, stop when `ul.products.columns-4` absent
- **Categories:** 21 categories in CATEGORIES dict (furniture + lighting)
- **Output:** `Category, Product URL, Image URL, Product Name, SKU`
- **File:** Mr. Brown London (Scraping & Spot Checking)/Code/Step1.py

---

### NEW CLASSICS (newclassicfurniture.com) — Selenium WooCommerce

- **Method:** Selenium (no WDM, uses default Chrome)
- **Card:** `.woolentor-grid-view-content .woolentor-product-image`
- **URL+Name:** `a[href]` + `a[title]` attribute
- **Image:** `img[src]` inside the anchor
- **Pagination:** `base_url + "page/{N}/"`, detect end via `a.next.page-numbers` absence
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** New Classics/Code/Step1.py

---

### NIERMANN WEEKS (niermannweeks.com) — Requests+BS4

- **Method:** Requests+BS4
- **Card:** `li.nw-product`
- **URL:** `a[href]`, **Image:** `img[src]`, **Name:** `h3`
- **Pagination:** `base_url + "page/{N}/"`, stop on status != 200
- **Output:** `Category, Product URL, Image URL, Product Name` (no SKU)
- **File:** Niermann Weeks (Scraping & Spot Checking)/Code/step1.py

---

### PALMER HARGRAVE (palmerhargrave.com) — Requests+BS4

- **Method:** Requests+BS4, single URL only
- **Card:** `article.e-add-post`
- **URL:** `a.e-add-post-image[href]`
- **Image:** `img[src]`
- **Name:** `h3.e-add-post-title a`
- **SKU:** `div.e-add-item_custommeta span`
- **Pagination:** None (single page)
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Palmer Hargrave/Code/step1.py

---

### PIERRE FREY (pierrefrey.com) — Requests+BS4

- **Method:** Requests+BS4
- **Card:** `div.resultListItem`
- **URL:** `a.resultListItem__link[href]` (urljoin with BASE_URL)
- **Image:** `img.resultListItem__img[src]`
- **Name:** `div.resultListItem__supTitle` + " " + `div.resultListItem__title`
- **SKU:** `div.resultListItem__subTitle`
- **Pagination:** `ul.pagination__list a.pagination__button--num` — visit each unvisited href
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Pierre Frey/Code/step1.py

---

### POWELL & BONNELL (powellandbonnell.com) — Selenium WooCommerce

- **Method:** Selenium + BS4 (infinite scroll only)
- **Card:** `li.product` (soup after scroll)
- **URL:** `a.woocommerce-LoopProduct-link[href]`
- **Name:** `h2.woocommerce-loop-product__title`
- **Image:** `div.product__img[style]` → CSS `url('...')` regex extract
- **Pagination:** Infinite scroll (no pagination buttons)
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** Powell & Bonnell/Code/step1.py

---

### QUATRINE (quatrine.com) — Requests+BS4

- **Method:** Requests+BS4
- **Card:** `article.product-list-item`
- **URL:** `h2.product-list-item-title a[href]` (urljoin)
- **Image:** `figure img[src]` (urljoin)
- **Name:** `h2.product-list-item-title a`
- **Pagination:** `li.next a[href]` → follow until None
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** Quatrine/Code/Step1.py

---

### REMAINS (remains.com) — Selenium, Custom Shopify-like

- **Method:** Selenium, hardcoded `C:/chromedriver.exe`
- **Wait:** `div.usf-results` container
- **Card:** `div.grid__item.grid-product`
- **URL:** `a.grid-product__link[href]`
- **Name:** `div.grid-product__title.grid-product__title--heading`
- **Image:** `img.grid-product__image` → `ensure_image_src()` (wait up to 4.5s for non-data: src)
- **Pagination:** `?page=N` via `build_page_url()`, stop when added==0 after one retry
- **Special:** `detach=True` — browser stays open even if Python exits
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** Remains (Scraping & Spot Checking)/Code/listpage.py

---

### ROWE (rowefurniture.com) — Requests+BS4

- **Method:** Requests+BS4
- **Card:** `div.product-item`
- **URL:** `div.picture a[href]` (urljoin)
- **Image:** `img.picture-img[src]`
- **Name:** `h2.product-title a`
- **SKU:** `div.sku`
- **Pagination:** `?pagenumber=N` or `&pagenumber=N` depending on existing `?` in URL
- **Hash strip:** `base_url = category_url.split("#", 1)[0]` before pagination
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Rowe/Code/Rowe_1.py

---

### SIEMON & SALAZAR (siemonandsalazar.com) — Playwright async (Wix)

- **Method:** `from playwright.async_api import async_playwright` — async!
- **Card:** `[data-hook="product-list-grid-item"]`
- **URL:** `[data-hook="product-item-root"] a[data-hook="product-item-container"][href]`
- **Name:** `[data-hook="product-item-name"]`
- **Image:** `wow-image img[src]` → `clean_image_url()` (keep only base Wix media URL)
  - Fallback: `wow-image[data-image-info]` JSON → `"uri":"..."` regex
- **SKU:** Generated from product name first word + CATEGORY_CODE + zero-padded index
- **Pagination:** None (infinite scroll — `scroll_to_bottom()` with up to 40 scrolls, 2.5s pause)
- **Wait:** `wait_until="networkidle"`, then `page.wait_for_selector('[data-hook="product-list-grid-item"]')`
- **Output:** `Product URL, Image URL, Product Name, SKU, Category`
- **File:** Siemon & Salazar/Code/Step1.py

---

### STAHL & BAND (stahlandband.com) — Requests+BS4

- **Method:** Requests+BS4
- **Card:** `div[class*='countergrid']` (lambda class check)
- **URL:** `a.box[href]`, **Image:** `img[src]`, **Name:** `h3`
- **SKU:** Generated — `STA-PU-{index}` (prefix changes per category)
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Stahl & Band/Code/Step1.py

---

### STUDIO TWENTY SEVEN (shop.studiotwentyseven.com) — Requests+BS4 Shopify

- **Method:** Requests+BS4
- **Card:** `div.boost-pfs-filter-product-item, div.block.grid-item`
- **URL:** `div.block-image a[href]` → prepend BASE_URL
- **Image:** `img.boost-pfs-filter-product-item-main-image` → `data-srcset` → `srcset` → `src`; `clean_image_url()` removes size suffix `_NNNx.ext`
- **Name:** `h2.block-title` (decompose `.block-vendor` span first), then optionally append vendor text
- **Price:** `span.boost-pfs-filter-product-item-regular-price`
- **SKU:** Generated — `VEN[:3] + category[:2] + index`
- **Pagination:** detect `a.boost-pfs-filter-next-btn, li.next a, a[rel='next']`
- **Output:** `Product URL, Image URL, Product Name, SKU, List Price (USD)`
- **File:** Studio Twenty Seven/Code/Step1.py

---

### SUDIO BEL VETRO (studiobelvetro.com) — Requests+BS4

- **Method:** Requests+BS4, single collections page
- **Card:** `div.col.c6`
- **URL:** `a[href]` (urljoin), **Image:** `img[src]` (urljoin), **Name:** `figcaption`
- **Pagination:** None (single page)
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** Sudio Bel Vetro/Code/Step1.py

---

### SUNPAN (sunpan.com) — Selenium Shopify

- **Method:** Selenium
- **Card:** `div.card-wrapper.product-card-wrapper`
- **URL:** XPath `./ancestor::a[1]` from card element
- **SKU:** `div.product__sku span.sku`
- **Name:** `h3.card__heading.h3`
- **Image:** `img.card-product-image[src]`
- **Pagination:** `base_url + "?page={N}"` (hardcoded `last_page` variable!)
- **Output:** `Product URL, Image URL, Product Name, SKU` (reordered in final DataFrame)
- **File:** Sunpan/Code/Sunpan_step1.py

---

### SURYA (surya.com) — Selenium + Chrome Debug Port

- **Method:** Selenium + `chromedriver_autoinstaller` + Chrome Debug Port 9222
- **Special:** `opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")` — connects to existing Chrome
- **Auto-launch:** `ensure_debug_chrome_running()` — starts Chrome with `--remote-debugging-port=9222` if not running
- **Card:** `div[data-test-selector='productGridItem']`
- **URL:** `a[data-test-selector='productImage'], a[data-test-selector='productDescriptionLink']`
- **Image:** `img[src/data-src/srcset]` → `picture source[srcset]` → `[style*='background-image']` CSS
- **SKU:** `span[data-test-selector='productStyleId']` etc. → `clean_sku()` (strip trailing `-p`, regex extract)
- **Scroll:** `smart_slow_scroll_to_bottom()` — SCROLL_STEP_PX=2000, STABLE_ROUNDS_LIMIT=45, max 7200s
- **⚠️ Output: `Product URL, Image URL, SKU` — NO Product Name in step1!**
- **File:** Surya (Scraping & Spot Checking)/Code/Surya_step1.py

---

### SUTHERLAND (sutherlandfurniture.com) — Selenium

- **Method:** Selenium + WebdriverManager
- **Card:** `div.grid__lg-quarter div.suthQuickViewCard.productCard__quickview`
- **URL:** `a.links__overlay[href]`
- **Image:** `img.js-dynamic-image.productCard__img[src|data-src]`
- **Name:** `div.productCard__name.notranslate`
- **Pagination:** `?pg=N` via `build_page_url()` (parse_qs + urlencode, doseq=True)
- **Wait:** `time.sleep(6)` per page
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** Sutherland/Code/Step1.py

---

### TROSCAN (troscandesign.com) — Requests+BS4

- **Method:** Requests+BS4, single page
- **Card:** `div.thumb`
- **URL:** `div.thumb[deeplink]` attr → `{URL}#{deeplink}` fragment URL
- **Name:** `div.name`
- **Image:** `img[src]` (prefix BASE_URL if starts with /)
- **SKU:** Generated — `TRO-SO-{index}` (prefix changes per category)
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Troscan/Code/Step1.py

---

### UTTER MOST (uttermost.com) — Playwright sync

- **Method:** `from playwright.sync_api import sync_playwright` — sync Playwright
- **Card:** `div.item-root-Chs`
- **URL:** `a.item-images--uD[href]` → urljoin BASE_URL
- **Name:** `a.item-name-LPg span`
- **SKU:** `p span.font-semibold`
- **Image:** `img[class*="rounded-"][data-src|src]`
- **Total pages:** `div.css-1m76rdz-singleValue` text → split "of" → int
- **Pagination:** `URL_TEMPLATE.format(page=page_number)`
- **Resource blocking:** `route.abort()` for stylesheet/font/media
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Utter Most/Code/Step1.py

---

### VERELLEN (verellen.biz) — Selenium, NEW FORMAT

- **Method:** Selenium + WebdriverManager, headless + disable images preference
- **Card:** `.product-widget, .product-item, .product-card` (multiple fallbacks)
- **Fallback:** link scan — `a[href*=verellen.biz]` where text is UPPERCASE + 2+ words
- **URL:** `a.p-0[href], a[href*='/products/'][href]`
- **Name:** `.product-name, [itemprop='name'], .name-wishlist-container a, h2, h3, a`
- **Pagination:** `.pagination .button-wrapper button` — read all page numbers (zero-padded: "01", "02"), click each
- **⚠️ NEW FORMAT output: `Manufacturer, Source, Product Name` — no Image or SKU!**
- **MANUFACTURER = "Verellen"** hardcoded
- **File:** Verellen/Code/step1.py

---

### VILL & HOUSE (vandh.com) — Requests+BS4 BigCommerce

- **Method:** Requests+BS4
- **Card:** `li.product`
- **URL:** `figure.card-figure a[href]` → urljoin "https://vandh.com"
- **Image:** `div.card-img-container img[src|data-src]`
- **Name:** `h3.card-title a`
- **SKU:** `div.card-text strong` → strip "SKU :" prefix
- **Pagination:** `?page=N` (first page = no param)
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Vill & House (Scraping & Spot Checking)/Code/Villa_house1.py

---

### WELLS ABBOTT (wellsabbott.com) — Playwright sync Shopify

- **Method:** `from playwright.sync_api import sync_playwright`
- **Card:** `li.usf-sr-product` (BS4 after `page.content()`)
- **URL+Name:** `h3.card__heading a.full-unstyled-link`
- **Image:** `.card__media img` or `img` → `//` → `https:` prefix fix
- **Pagination:** None (lazy scroll with 5 passes)
- **Output:** `Product URL, Image URL, Product Name` (no SKU)
- **File:** Wells Abbott/Code/Step1.py

---

### WORLDS AWAY (worlds-away.com) — Selenium WooCommerce + EXTRA COLUMNS

- **Method:** Selenium + `chromedriver_autoinstaller`
- **Card:** `li.product`
- **URL+Name:** `div.card-body a[href]` + text
- **Image:** tries 5 selectors: `img.card-image, img.card-img-top, img.attachment-woocommerce_thumbnail, img.wp-post-image, img` → `data-src` or `src`
- **Extra:** `div.card-dimensions-standard` → Size; `div.card-desc` → Description
- **Pagination:** Infinite scroll (3-round stable stop), MAX_SCROLL_ROUNDS=200
- **⚠️ Output: `Category, Product URL, Image URL, Product Name, Size, Description` — 2 EXTRA COLUMNS vs standard!**
- **File:** Worlds Away (Scraping & Spot Checking)/Code/Step1.py

---

### ZUO (zuomod.com) — Selenium Magento, NEW FORMAT

- **Method:** Selenium, CDP injection (`navigator.webdriver = undefined`)
- **Card:** `li.product-item, div.product-item, .product-item-info, li.item.product, div.item.product`
- **URL:** `a.product-item-link, a.product-item-photo, a[href*='zuomod.com']`
- **Name:** `a.product-item-link` or `.product-item-name a`
- **Image:** `img.product-image-photo[data-src|src]`
- **SKU:** extracted in step2 (step1 has no SKU)
- **Pagination:** `a.action.next, li.pages-item-next a` → next href; per-page limiter: set 160 option
- **⚠️ NEW FORMAT output: `Manufacturer, Source, Product Name, Image URL, Category`**
- **MANUFACTURER = "Zuo"** hardcoded
- **File:** Zuo/Code/Step1.py

---

### INVISIBLE COLLECTION (theinvisiblecollection.com) — Selenium Algolia

- **Method:** Selenium + WebdriverManager
- **Engine:** Algolia-powered JS rendering
- **Card:** `.ais-Hits-item article.hit`
- **URL+Name:** `a.vsz-product-container[href][data-product_name]`
- **Image:** `img.pro-front-img[src]`, fallback `img[src*='http']`
- **Price:** `.wcpbc-price` → "Price upon request" → 0; else `.woocommerce-Price-amount` → strip symbols
- **Pagination:** `.ais-Pagination-item--nextPage a.ais-Pagination-link` click (check disabled class first)
- **Special:** `nuke_popups()` JS function removes fixed/absolute overlays > z-index 1000
- **Output:** `Product URL, Image URL, Product Name, List Price`
- **File:** Invisible Collection/invisible_step1.py (at root, not in Code/ subfolder!)

---

### BENNETT (bennetttothetrade.com) — Requests+BS4 Shopify

- **Method:** Requests+BS4
- **Card:** `div.card-wrapper.product-card-wrapper`
- **URL:** `h3.card__heading.h5 a.full-unstyled-link[href]` → prepend base_domain
- **Image:** `div.media img[src]` → `//` → `https:` fix
- **Name:** same `a` element text
- **SKU:** = Product Name (site quirk — no separate SKU, name is used as SKU)
- **Pagination:** `?page=N`, stop on empty products or status != 200
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** Bennett (Scraping & Spot Checking)/Code/step1.py

---

### CR LAINE (crlaine.com) — Selenium Custom PHP

- **Method:** Selenium + WebdriverManager
- **Card:** `div.style_thumbs` or `div[stylename]` attribute
- **URL:** `a.pageLoc[href]` → urljoin BASE_URL
- **Image:** `img[src|lazyload|data-src]` → urljoin BASE_URL
- **Name:** `div.stylename` + optionally append last `div.stylenumber` text
- **SKU:** `div.stylenumber` — first SKU-like value (has digits, ≤12 chars or has hyphen)
- **Pagination:** Infinite scroll (scroll until count stable, MAX_SCROLLS=60)
- **Output:** `Product URL, Image URL, Product Name, SKU`
- **File:** CR Laine/Code/CR_Laine1.py

---

### HIGHLAND HOUSE (highlandhousefurniture.com) — Selenium ASP.NET

- **Method:** Selenium (no WDM)
- **Card:** `li.prodListingDiv div.prodSearchDiv`
- **URL:** `a[href]` → prepend BASE_URL
- **Image:** `img.prodSearchImage[src]` → prepend BASE_URL
- **SKU:** `div[style*='margin:5px 0;'] strong` text (no Product Name in step1!)
- **Pagination:** Try `span.viewAll.prodPageNavItem` click first (loads all), else `span.nextPage.prodPageNavItem` click
- **URL format:** `?TypeID=N` (category ID in URL)
- **⚠️ Column casing: `Product Url, Image Url, Sku` — lowercase Url!**
- **⚠️ No Product Name extracted in step1 — only URL, Image, SKU**
- **File:** Highland House/Highland House (Scraping & Spot Checking)/Code/Highland_House_1.py

---

### JOHNATHAN BROWN INC (jonathanbrowninginc.com) — Requests+BS4 Data Attributes

- **Method:** Requests+BS4, single page
- **Card:** `li.grid-item` with data attributes
- **URL:** built from `data-category` + `data-deeplink` → `{BASE_URL}/products/{category}?deep={deeplink}`
- **Image:** `data-resting` attr → prepend BASE_URL if starts with /
- **Name:** `data-product-name` attribute
- **SKU:** None (no SKU available)
- **Pagination:** None (single page, all products in DOM)
- **Output:** `Product URL, Image URL, Product Name`
- **File:** Johnathan Brown Inc/Code/Step1.py

---

### PHASE-05 (AU) SPECIAL LIBRARY SUMMARY

| Vendor | Library | Why |
|--------|---------|-----|
| Dana Creath | Playwright sync | WordPress with heavy JS rendering |
| Siemon & Salazar | Playwright async | Wix site requires real browser |
| Wells Abbott | Playwright sync | Shopify with filter JS |
| Utter Most | Playwright sync | Custom React frontend |
| Collier Webb | undetected_chromedriver | Cloudflare protection |
| Surya | Chrome Debug Port 9222 | Bot detection bypass |
| Janus et Cie | Chrome Debug Port 9222 | Bot detection bypass |
| Invisible Collection | Algolia JS + nuke_popups | Heavy popup overlays |

### PHASE-05 (AU) OUTPUT FORMAT SUMMARY

| Format | Vendors |
|--------|---------|
| **NEW** (Manufacturer + Source) | Hennepin Made, Verellen, Zuo |
| **NEW partial** (Source only) | Gabby (step1), Visual Comfort (step3) |
| **STANDARD OLD** (Product URL, Image URL, Product Name, SKU) | Most vendors |
| **Extra columns** | Worlds Away (+Size, Description), Studio Twenty Seven (+List Price), Siemon & Salazar (+Category) |
| **No Product Name step1** | Surya (only URL, Image, SKU), Highland House (only URL, Image, SKU) |
| **No SKU** | New Classics, Niermann Weeks, Powell & Bonnell, Quatrine, Remains, Sutherland, Sudio Bel Vetro, Wells Abbott, Johnathan Brown Inc |
