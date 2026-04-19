import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import random
from datetime import datetime

# ========= Settings =========
SHOW_PREVIEW_ROWS = 3
AUTOSAVE_EVERY = 25
TIMEOUT = 30

# ======== File paths ========
input_path = "theodore_alexander_Floor_lamps.xlsx"
output_path = "theodore_alexander_Floor_lamps_Final.xlsx"

# ======== Read input ========
df = pd.read_excel(input_path)

# ======== Ensure all columns in client-specified order ========
final_columns = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family Name",
    "Description", "List Price", "Weight", "Net Weight",
    "Width (in)", "Depth (in)", "Diameter (in)", "Height (in)",
    "Finish", "Collection", "Room / Type", "Main Materials",
    "Shapes Materials", "Seat Height", "Arm Height",
    "Inside Seat Depth", "Inside Seat Width"
]

for col in final_columns:
    if col not in df.columns:
        df[col] = None

df = df.reindex(columns=final_columns)

# ======== Helpers ========
HDRS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
}

def clean_text(s):
    if not s:
        return None
    return re.sub(r"\s+", " ", s).strip()

def extract_description(psoup):
    blk = psoup.select_one("div.product_detail_info_description > p")
    if blk:
        return clean_text(blk.get_text())
    block = psoup.select_one("div.product_desc")
    if block:
        return clean_text(block.get_text(" ", strip=True))
    nav_detail = psoup.select_one("#nav-detail .col-xl-8")
    if nav_detail:
        p = nav_detail.find("p")
        if p:
            return clean_text(p.get_text(" ", strip=True))
    p_any = psoup.find("p")
    return clean_text(p_any.get_text()) if p_any else None

def extract_dimensions_including_diameter(psoup):
    result = {"Width": None, "Depth": None, "Height": None, "Diameter": None}
    table = psoup.select_one("table.tableDimension")
    if not table:
        return result

    headers = []
    thead = table.find("thead")
    if thead:
        ths = thead.find_all("th")
        for th in ths:
            headers.append(clean_text(th.get_text()) or "")

    name_to_idx = {}
    for i, h in enumerate(headers):
        hl = (h or "").strip().lower()
        if hl in ["width", "depth", "height", "diameter"]:
            name_to_idx[hl.capitalize()] = i

    body_rows = table.select("tr.tableBodyDimension")
    in_row = None
    for tr in body_rows:
        th = tr.find("th")
        if th and (clean_text(th.get_text()) or "").lower() == "in":
            in_row = tr
            break
    if not in_row:
        return result

    cells = [in_row.find("th")] + in_row.find_all("td")

    for key, idx in name_to_idx.items():
        if idx < len(cells):
            result[key] = clean_text(cells[idx].get_text())

    return result

def extract_finish(psoup):
    for li in psoup.select("div.col-xl-8.col-md-12.w-100 li"):
        label = li.select_one("span.product_tab_content_detail-title")
        if label and "finish" in label.get_text(strip=True).lower():
            val = li.select_one("span.col-xl-8")
            if val:
                return clean_text(val.get_text())
    return None

# 🆕 Extract only Gross Weight
def extract_weight(psoup):
    """
    Extracts only Gross Weight (for Weight column).
    Returns None if not found.
    """
    rows = psoup.select("div.row.p-0.m-0.w-100")
    for row in rows:
        title_div = row.select_one(".product_tab_content_detail-title")
        if not title_div:
            continue
        title_text = clean_text(title_div.get_text(" ", strip=True)).lower()
        if "gross weight" in title_text:
            content_div = row.select_one(".product_tab_content_detail-content")
            if content_div:
                all_spans = content_div.find_all("span")
                for sp in all_spans:
                    val = clean_text(sp.get_text())
                    if val and re.search(r"\d", val):
                        return val
    return None

def extract_price(psoup):
    candidates = psoup.select(".price, .product-price, [data-price], .ta-price")
    for c in candidates:
        txt = clean_text(c.get_text())
        if txt and re.search(r"[\$£€]\s?\d", txt):
            return txt
    for tag in psoup.find_all(string=re.compile(r"list\s*price", re.I)):
        parent = tag.parent
        if parent:
            txt = clean_text(parent.get_text(" ", strip=True))
            if txt and re.search(r"[\$£€]\s?\d", txt):
                return txt
    return None

# ======== Extract from Details tab ========
def extract_details_tab_fields(psoup):
    fields = {
        "Collection": None,
        "Room / Type": None,
        "Main Materials": None,
        "Shapes Materials": None,
        "Net Weight": None,
        "Seat Height": None,
        "Arm Height": None,
        "Inside Seat Depth": None,
        "Inside Seat Width": None,
    }

    detail_rows = psoup.select("div#nav-detail div.row.p-0.m-0.w-100")
    for row in detail_rows:
        title = row.select_one("span.product_tab_content_detail-title, div.product_tab_content_detail-title span")
        value = row.select_one("span.col-xl-8, div.product_tab_content_detail-content")
        if not title or not value:
            continue

        label = clean_text(title.get_text(":"))
        val = clean_text(value.get_text(" ", strip=True))
        if not label or not val:
            continue

        for key in fields.keys():
            if key.lower() in label.lower():
                fields[key] = val
                break

    # Special handle for Room / Type links
    if not fields["Room / Type"]:
        room_el = psoup.select_one(
            "div#nav-detail span.product_tab_content_detail-title:-soup-contains('Room / Type')"
        )
        if room_el:
            next_span = room_el.find_next("span", class_="col-xl-8")
            if next_span:
                links = [a.get("title") for a in next_span.find_all("a", title=True)]
                if links:
                    fields["Room / Type"] = " / ".join(links)

    return fields

def short(s, n=90):
    if not s:
        return ""
    return s if len(s) <= n else s[:n-1] + "…"

def autosave(df, path_xlsx):
    ts = datetime.now().strftime("%H:%M:%S")
    df.to_excel(path_xlsx, index=False)
    print(f"💾 [{ts}] Autosaved to: {path_xlsx}", flush=True)

# ======== Scrape loop ========
try:
    from tqdm import tqdm
    iterator = tqdm(df.iterrows(), total=len(df), desc="Scraping")
except Exception:
    iterator = df.iterrows()

count = 0
for idx, row in iterator:
    url = row.get("Product URL")
    if not isinstance(url, str) or not url.strip():
        continue

    try:
        res = requests.get(url, headers=HDRS, timeout=TIMEOUT)
        res.raise_for_status()
        psoup = BeautifulSoup(res.text, "html.parser")

        # Extract core fields
        df.at[idx, "Description"] = extract_description(psoup)
        dims = extract_dimensions_including_diameter(psoup)
        df.at[idx, "Width (in)"]    = dims.get("Width")
        df.at[idx, "Depth (in)"]    = dims.get("Depth")
        df.at[idx, "Height (in)"]   = dims.get("Height")
        df.at[idx, "Diameter (in)"] = dims.get("Diameter")
        df.at[idx, "Finish"]        = extract_finish(psoup)
        df.at[idx, "Weight"]        = extract_weight(psoup)
        df.at[idx, "List Price"]    = extract_price(psoup)

        # Extract detail tab fields
        details = extract_details_tab_fields(psoup)
        for k, v in details.items():
            df.at[idx, k] = v

        # Product Family Name = Product Name
        if "Product Name" in df.columns:
            df.at[idx, "Product Family Name"] = df.at[idx, "Product Name"]
        elif "Product" in df.columns:
            df.at[idx, "Product Family Name"] = df.at[idx, "Product"]

        count += 1

        print(f"✅ {count}/{len(df)} | {url}")

        if count <= SHOW_PREVIEW_ROWS:
            print(f"— Preview —\nDescription: {df.at[idx,'Description']}\nGross Weight: {df.at[idx,'Weight']}\nNet Weight: {df.at[idx,'Net Weight']}\nPrice: {df.at[idx,'List Price']}\nFinish: {df.at[idx,'Finish']}\n— End —\n")

        if AUTOSAVE_EVERY and (count % AUTOSAVE_EVERY == 0):
            autosave(df, output_path)

        time.sleep(random.uniform(1.5, 3.5))

    except Exception as e:
        print(f"⚠️ Error scraping {url}: {e}", flush=True)

# ======== Final Save ========
autosave(df, output_path)
print(f"🎯 Done! Total scraped: {count}/{len(df)}")
