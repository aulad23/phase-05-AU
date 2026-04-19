import requests
from bs4 import BeautifulSoup
from urllib.parse import unquote
import pandas as pd
import re

# ─────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────
URL        = "https://www.troscandesign.com/products/seating/benches"
BASE_URL   = "https://www.troscandesign.com"
VENDOR     = "Troscan"
CATEGORY   = "Benches"
SKU_PREFIX = VENDOR[:3].upper() + "-" + CATEGORY[:2].upper()

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# ─────────────────────────────────────────
# FINISH MAP
# finishes="93,21,20,18,17,16,15,14,92,13,97,10,6,8,7,88,9"
# positionally matched with #finish-thumbs <li> items (document 5)
# ─────────────────────────────────────────
FINISH_MAP = {
    "93": "Dover - walnut",
    "21": "clay - walnut",
    "20": "atlantic grey - walnut",
    "18": "natural - walnut",
    "17": "butternut - walnut",
    "16": "irish brown - walnut",
    "15": "cinnamon brown - walnut",
    "14": "ebony brown - walnut",
    "92": "Squall - walnut",
    "13": "roasted brown - walnut",
    "97": "Winter Grey - oak",
    "10": "atlantic - oak",
    "6":  "roasted brown - oak",
    "8":  "irish - oak",
    "7":  "natural - oak",
    "88": "Eclipse - oak",
    "9":  "coastal - oak",
}


# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────
def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")


def product_family_id(product_name):
    return product_name.split(" - ")[0].strip()


def resolve_finishes(finish_ids_str):
    if not finish_ids_str:
        return ""
    ids = [i.strip() for i in finish_ids_str.split(",") if i.strip()]
    return " | ".join(FINISH_MAP[i] for i in ids if i in FINISH_MAP)


def get_finish_ids(thumb):
    """
    Try 3 sources to get finish IDs:
    1. finishes attribute on thumb div
    2. <div class="button finishes"> inside thumb
    3. data-finishes attribute
    """
    for attr in ["finishes", "data-finishes"]:
        val = thumb.get(attr, "").strip()
        if val:
            return val

    btn = thumb.find("div", class_="finishes")
    if btn:
        val = btn.get("finishes", "").strip()
        if val:
            return val

    return ""


def parse_dimensions(raw):
    """
    Decode URL-encoded dimension attribute.
    Extracts from FIRST dimension line:
      W → Width | D → Depth | Diam/Dia/DIA → Diameter | H → Height
    Separate lines: Seat Height, Arm Height
    """
    decoded = unquote(raw)
    text    = re.sub(r"<[^>]+>", "\n", decoded)
    lines   = [l.strip() for l in text.splitlines() if l.strip()]

    width = depth = diameter = height = seat_h = arm_h = dim_line = ""

    for line in lines:
        if re.search(r"\d+\.?\d*\s*(W|D|Diam|Dia|DIA|H)\b", line, re.I) and not dim_line:
            dim_line = line
            for val, unit in re.findall(r"([\d.]+)\s*(Diam|Dia|DIA|W|D|H)\b", line, re.I):
                u = unit.upper()
                if u == "W" and not width:       width = val
                elif u == "D" and not depth:     depth = val
                elif u in ("DIAM","DIA") and not diameter: diameter = val
                elif u == "H" and not height:    height = val

        elif re.search(r"seat\s*height", line, re.I) and not seat_h:
            m = re.search(r"([\d.]+)", line)
            seat_h = m.group(1) if m else ""

        elif re.search(r"arm\s*height", line, re.I) and not arm_h:
            m = re.search(r"([\d.]+)", line)
            arm_h = m.group(1) if m else ""

    return width, depth, diameter, height, seat_h, arm_h, dim_line


# ─────────────────────────────────────────
# TRY TO BUILD DYNAMIC FINISH MAP FROM PAGE
# If page's #finish-thumbs has same count as first product's IDs → zip them
# Otherwise use hardcoded FINISH_MAP
# ─────────────────────────────────────────
def build_finish_map(soup):
    try:
        first_thumb = soup.find("div", class_="thumb")
        ul          = soup.find(id="finish-thumbs")
        if not first_thumb or not ul:
            return FINISH_MAP

        ids_str = get_finish_ids(first_thumb)
        ids     = [i.strip() for i in ids_str.split(",") if i.strip()]
        lis     = ul.find_all("li")

        if ids and len(ids) == len(lis):
            dynamic = {}
            for fid, li in zip(ids, lis):
                name   = li.get("data-finish-name", "").strip()
                parent = li.get("data-finish-parent", "").strip()
                dynamic[fid] = f"{name} - {parent}" if parent else name
            print(f"Dynamic finish map: {len(dynamic)} entries")
            return dynamic
    except Exception as e:
        print(f"Dynamic map failed: {e}")

    print(f"Using hardcoded finish map: {len(FINISH_MAP)} entries")
    return FINISH_MAP


# ─────────────────────────────────────────
# MAIN SCRAPE
# ─────────────────────────────────────────
def scrape():
    print("Fetching page...")
    soup       = get_soup(URL)
    finish_map = build_finish_map(soup)

    thumbs = soup.find_all("div", class_="thumb")
    print(f"Found {len(thumbs)} products\n")

    rows = []
    for idx, thumb in enumerate(thumbs, start=1):
        name_tag     = thumb.find("div", class_="name")
        product_name = name_tag.get_text(strip=True) if name_tag else "N/A"

        deeplink    = thumb.get("deeplink", "").strip()
        product_url = f"{URL}#{deeplink}" if deeplink else URL

        img_tag   = thumb.find("img")
        img_src   = img_tag["src"] if img_tag and img_tag.get("src") else ""
        image_url = BASE_URL + img_src if img_src.startswith("/") else img_src

        sku       = f"{SKU_PREFIX}-{str(idx).zfill(3)}"
        family_id = product_family_id(product_name)

        dim_raw = thumb.get("dimensions", "")
        width, depth, diameter, height, seat_h, arm_h, dim_line = parse_dimensions(dim_raw)

        finish_ids_str = get_finish_ids(thumb)
        if finish_ids_str:
            ids      = [i.strip() for i in finish_ids_str.split(",") if i.strip()]
            finishes = " | ".join(finish_map[i] for i in ids if i in finish_map)
        else:
            finishes = ""

        rows.append({
            "Product URL":       product_url,
            "Image URL":         image_url,
            "Product Name":      product_name,
            "SKU":               sku,
            "Product Family Id": family_id,
            "Description":       "",
            "Weight":            "",
            "Width":             width,
            "Depth":             depth,
            "Diameter":          diameter,
            "Height":            height,
            "Seat Height":       seat_h,
            "Arm Height":        arm_h,
            "Finish":            finishes,
            "Dimension":         dim_line,
        })

        print(f"[{idx:02d}] {product_name}")
        print(f"      SKU:        {sku}")
        print(f"      Family ID:  {family_id}")
        print(f"      Dim:        {dim_line}")
        print(f"      W={width}  D={depth}  Dia={diameter}  H={height}  SeatH={seat_h}  ArmH={arm_h}")
        print(f"      Finish IDs: {finish_ids_str}")
        print(f"      Finishes:   {finishes[:80]}{'...' if len(finishes)>80 else ''}")
        print()

    df = pd.DataFrame(rows, columns=[
        "Product URL", "Image URL", "Product Name", "SKU",
        "Product Family Id", "Description",
        "Weight", "Width", "Depth", "Diameter", "Height",
        "Seat Height", "Arm Height", "Finish", "Dimension"
    ])

    df.to_excel("troscan_Benches.xlsx", index=False)
    print(f"Done! {len(rows)} products saved to troscan_Benches.xlsx")


if __name__ == "__main__":
    scrape()