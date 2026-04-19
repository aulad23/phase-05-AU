import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

BASE_URL = "https://honning.us"
INPUT_FILE = "Mirrors.xlsx"
OUTPUT_FILE = "Mirrors_final.xlsx"


def parse_dimensions(dimension_str: str) -> dict:
    result = {
        "Width":    "",
        "Depth":    "",
        "Diameter": "",
        "Length":   "",
        "Height":   "",
        "Weight":   "",
    }

    if not dimension_str or pd.isna(dimension_str):
        return result

    # ── Remove inch symbols & unicode quotes ──────────────────────
    clean = str(dimension_str)
    for ch in ['"', '\u2019', '\u201d', '\u2018', '\u201c', "'"]:
        clean = clean.replace(ch, '')

    clean = clean.upper()

    patterns = {
        "Width":    r"([\d.]+)\s*W\b",
        "Depth":    r"([\d.]+)\s*D\b",
        "Height":   r"([\d.]+)\s*H\b",
        "Length":   r"([\d.]+)\s*L\b",
        "Diameter": r"([\d.]+)\s*(?:DIA|DIAM)\b",
        "Weight":   r"([\d.]+)\s*(?:WEIGHT|WEIGTH|LBS?|IBS?)\b",
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, clean)
        if match:
            result[key] = match.group(1)

    return result


def is_dimension_line(text: str) -> bool:
    """Check if a line looks like a dimension string.
    Handles both:
      - Lines starting with digits:  '65W x 40D x 30H'
      - Lines starting with size names: 'Queen: 65W x 86.75D x 91.75H'
    """
    if not text:
        return False
    # Starts with a digit (original behavior)
    if text[0].isdigit():
        return True
    # Starts with a bed-size label like "Queen:", "King:", "CA King:", etc.
    if re.match(r"(?i)^(queen|king|ca\s*king|twin|full|standard)\s*:", text):
        return True
    return False


def scrape_product_details(product_url: str) -> dict:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
        )
    }

    try:
        resp = requests.get(product_url, headers=headers, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"  ❌ Error: {e}")
        return {
            "Description": "", "Dimension": "", "Tearsheet Link": "",
            "Width": "", "Depth": "", "Diameter": "",
            "Length": "", "Height": "", "Weight": "",
        }

    soup = BeautifulSoup(resp.text, "html.parser")
    figcaption = soup.select_one("figcaption.image-card-wrapper")

    description = ""
    dimension   = ""
    tearsheet   = ""

    if figcaption:
        subtitle_div = figcaption.select_one("div.image-subtitle")
        if subtitle_div:
            paras = subtitle_div.find_all("p")
            dimension_lines = []  # collect ALL dimension lines (beds have multiple)

            for p in paras:
                text = p.get_text(" ", strip=True)

                # ── Dimension ──────────────────────────────────────
                # FIX: Use is_dimension_line() instead of text[0].isdigit()
                if "sqsrte-small" in p.get("class", []) and is_dimension_line(text):
                    # Get text before first <br> (original logic)
                    raw_p = BeautifulSoup(str(p), "html.parser").find("p")
                    first_chunk = ""
                    for content in raw_p.children:
                        if hasattr(content, "name") and content.name == "br":
                            break
                        first_chunk += str(content)
                    line = BeautifulSoup(first_chunk, "html.parser").get_text(strip=True)
                    if line:
                        dimension_lines.append(line)

                # ── Description ────────────────────────────────────
                if "sqsrte-small" not in p.get("class", []) and text and "Download" not in text:
                    description = text

        # ── Combine all dimension lines ────────────────────────────
        # For beds: "Queen: 65W x 86.75D x 91.75H | King: 81W x 86.75D x 91.75H"
        # For others: just the single line
        dimension = " | ".join(dimension_lines) if dimension_lines else ""

        # ── Tearsheet Link ─────────────────────────────────────────
        a_tag = figcaption.find("a", href=True)
        if a_tag and "Tearsheet" in a_tag.get_text():
            href = a_tag["href"]
            tearsheet = BASE_URL + href if href.startswith("/") else href

    # Parse dimensions from the FIRST size variant (e.g. Queen for beds)
    first_dim = dimension_lines[0] if dimension_lines else ""
    # If it has a "Label:" prefix, strip it to get just the dimension part
    if ":" in first_dim:
        first_dim = first_dim.split(":", 1)[1].strip()
    dim_parts = parse_dimensions(first_dim)

    return {
        "Description":    description,
        "Dimension":      dimension,
        "Tearsheet Link": tearsheet,
        **dim_parts,
    }


if __name__ == "__main__":
    df = pd.read_excel(INPUT_FILE)
    print(f"📥 Loaded {len(df)} rows from {INPUT_FILE}\n")

    results = []

    for i, row in df.iterrows():
        product_url  = row["Product URL"]
        image_url    = row["Image URL"]
        product_name = row["Product Name"]
        sku          = row["SKU"]

        print(f"[{i+1:>3}] Scraping: {product_name}")
        details = scrape_product_details(product_url)

        results.append({
            "Product URL":       product_url,
            "Image URL":         image_url,
            "Product Name":      product_name,
            "SKU":               sku,
            "Product Family Id": product_name,
            "Description":       details["Description"],
            "Weight":            details["Weight"],
            "Width":             details["Width"],
            "Depth":             details["Depth"],
            "Diameter":          details["Diameter"],
            "Length":            details["Length"],
            "Height":            details["Height"],
            "Dimension":         details["Dimension"],
            "Tearsheet Link":    details["Tearsheet Link"],
        })

        print(f"       Dimension : {details['Dimension']}")
        print(f"       Width     : {details['Width']}")
        print(f"       Depth     : {details['Depth']}")
        print(f"       Height    : {details['Height']}")
        print(f"       Diameter  : {details['Diameter']}")
        print(f"       Length    : {details['Length']}")
        print(f"       Weight    : {details['Weight']}")
        print(f"       Tearsheet : {details['Tearsheet Link']}")
        print(f"       Desc      : {details['Description'][:80]}...")
        print()

        time.sleep(0.5)

    # ── Final Column Order ─────────────────────────────────────────
    out_df = pd.DataFrame(results, columns=[
        "Product URL",
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
        "Dimension",
        "Tearsheet Link",
    ])

    out_df.to_excel(OUTPUT_FILE, index=False)

    print(f"\n✅ {len(out_df)} টি product পাওয়া গেছে।")
    print(f"📄 Final data সেভ হয়েছে: {OUTPUT_FILE}")
    print(out_df.to_string(index=False))