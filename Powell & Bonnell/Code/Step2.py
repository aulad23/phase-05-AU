import re
import time
import pandas as pd
import requests
from bs4 import BeautifulSoup

INPUT_EXCEL = "powell_bonnell_Mirrors.xlsx"
OUTPUT_EXCEL = "powell_bonnell_Mirrors_OUTPUT.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
}

# -----------------------------
# Helpers
# -----------------------------
def clean_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def first_match(pattern, text, flags=re.I):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""

def to_number_str(val: str) -> str:
    if not val:
        return ""
    m = re.search(r"(\d+(?:\.\d+)?)", val)
    return m.group(1) if m else ""

def fetch_soup(url: str, tries=3, sleep=1.2) -> BeautifulSoup | None:
    for i in range(tries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 200 and r.text:
                return BeautifulSoup(r.text, "lxml")
        except Exception:
            pass
        time.sleep(sleep * (i + 1))
    return None

def is_dimension_like(text: str) -> bool:
    t = (text or "").lower()
    if not t:
        return False
    if "product" in t:
        return False
    # dimension-ish signals
    if re.search(r'(\d+(?:\.\d+)?)\s*["”]?\s*(h|w|d)\b', t, re.I):
        return True
    if re.search(r"\boah\b", t, re.I):
        return True
    if "diameter" in t or re.search(r"\bdia\.?\b", t, re.I):
        return True
    if "seat height" in t or "seat depth" in t or "arm height" in t:
        return True
    if "projection" in t:
        return True
    if "shade" in t:
        return True
    return False

# -----------------------------
# Extractors
# -----------------------------
def extract_description(summary_div: BeautifulSoup) -> str:
    p = summary_div.find("p")
    return clean_text(p.get_text(" ", strip=True)) if p else ""

def extract_sku(summary_div: BeautifulSoup) -> str:
    for h in summary_div.find_all("h5"):
        t = clean_text(h.get_text(" ", strip=True))
        if "PRODUCT" in t.upper():
            sku = first_match(r"PRODUCT\s*([A-Za-z0-9\-\/]+)", t)
            if sku:
                return sku
    # sometimes "PRODUCT 9150" appears as plain text
    txt = clean_text(summary_div.get_text(" | ", strip=True))
    sku2 = first_match(r"\bPRODUCT\s*([A-Za-z0-9\-\/]+)", txt)
    return sku2

def extract_dimension_block_text(summary_div: BeautifulSoup) -> str:
    """
    Primary: find dimension-ish text inside <h5> blocks.
    Fallback: if not found, scan all text lines inside summary and pick the best dimension-looking line.
    This fixes pages like:
    - "25” DIA. X 27” H" (plain text, not in h5.p1)
    """
    # Primary: h5 candidates
    h5s = summary_div.find_all("h5")
    candidates = []
    for h in h5s:
        t = clean_text(h.get_text(" | ", strip=True))
        if is_dimension_like(t):
            candidates.append(t)

    if candidates:
        candidates = sorted(candidates, key=lambda x: len(x), reverse=True)
        return candidates[0]

    # Fallback: scan all text lines
    lines = []
    for s in summary_div.stripped_strings:
        t = clean_text(str(s))
        if t:
            lines.append(t)

    # keep only dimension-like lines
    dim_lines = [l for l in lines if is_dimension_like(l)]
    if not dim_lines:
        return ""

    # pick best candidate (longest)
    dim_lines = sorted(dim_lines, key=lambda x: len(x), reverse=True)
    return dim_lines[0]

def extract_upholstery_options(summary_div: BeautifulSoup) -> str:
    h5s = summary_div.find_all("h5")
    for h in h5s:
        t = clean_text(h.get_text(" | ", strip=True))
        if "UPHOLSTERY OPTIONS" in t.upper():
            idx = t.upper().find("UPHOLSTERY OPTIONS")
            return clean_text(t[idx:])
    return ""

# -----------------------------
# Parsers
# -----------------------------
def parse_dimension_fields(dim_text: str) -> dict:
    """
    Requirements:
    - 'seat 17"h' -> Seat Height = 17
    - Shade Details -> only numbers comma-separated (e.g., "22,13")
    - 'oah' -> Height (e.g., 90" oah)
    - 'dia.' -> Diameter (e.g., 30"dia.)
    - Example: 90” oah x 30”dia. shade -> Height=90, Diameter=30
    """
    out = {
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Height": "",
        "Seat Height": "",
        "Seat Depth": "",
        "Arm Height": "",
        "Shade Details": "",
    }

    if not dim_text:
        return out

    raw = dim_text.replace("”", '"').replace("″", '"').replace("’", "'").replace("×", "x")
    parts = [clean_text(p) for p in re.split(r"\||<br>|;|\n", raw) if clean_text(p)]

    shade_nums = []

    for p in parts:
        low = p.lower()

        # seat 17"h pattern
        if not out["Seat Height"]:
            m_seat = re.search(r"\bseat\s*(\d+(?:\.\d+)?)\s*\"?\s*h\b", p, re.I)
            if m_seat:
                out["Seat Height"] = m_seat.group(1)

        # Seat/Arm normal patterns
        if "seat height" in low:
            out["Seat Height"] = to_number_str(p)
            continue
        if "seat depth" in low:
            out["Seat Depth"] = to_number_str(p)
            continue
        if "arm height" in low:
            out["Arm Height"] = to_number_str(p)
            continue

        # Weight
        if re.search(r"\b(weight|lbs|ibs|lb|kg)\b", low):
            out["Weight"] = to_number_str(p)
            continue

        # OAH -> Height (explicit)
        if not out["Height"]:
            m_oah = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*oah\b', p, re.I)
            if m_oah:
                out["Height"] = m_oah.group(1)

        # DIA -> Diameter (explicit)  ✅ (your requirement)
        if not out["Diameter"]:
            m_dia = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*dia\.?\b', p, re.I)
            if m_dia:
                out["Diameter"] = m_dia.group(1)

        # General diameter wording
        if (not out["Diameter"]) and ("diameter" in low or re.search(r"\bdia\.?\b", low)):
            out["Diameter"] = to_number_str(p)

        # h/w/d tokens
        if not out["Height"]:
            m_h = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*h\b', p, re.I)
            if m_h:
                out["Height"] = m_h.group(1)

        if not out["Width"]:
            m_w = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*w\b', p, re.I)
            if m_w:
                out["Width"] = m_w.group(1)

        if not out["Depth"]:
            m_d = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*d\b', p, re.I)
            if m_d:
                out["Depth"] = m_d.group(1)

        # projection => Depth
        if "projection" in low and not out["Depth"]:
            out["Depth"] = to_number_str(p)

        # overall height fallback
        if ("overall height" in low or low.endswith("height")) and not out["Height"]:
            out["Height"] = to_number_str(p)

        # Shade -> only numbers
        if "shade" in low:
            nums = re.findall(r"(\d+(?:\.\d+)?)", p)
            for n in nums:
                shade_nums.append(n)

    # Shade Details final: only numbers comma-separated, no duplicates
    if shade_nums:
        seen = set()
        ordered = []
        for n in shade_nums:
            if n not in seen:
                seen.add(n)
                ordered.append(n)
        out["Shade Details"] = ",".join(ordered)

    return out

def parse_upholstery_fields(up_text: str) -> dict:
    """
    COM -> only number
    COL -> only number
    CUSHION -> keep text (if present)
    """
    out = {"Com": "", "Col": "", "Cushion": ""}
    if not up_text:
        return out

    t = up_text.replace("\xa0", " ")
    t = re.sub(r"\s+", " ", t).strip()

    m = re.search(r"\bCOM\s*:\s*(.+?)(?=\s*\bCOL\s*:|\s*\bCUSHION\s*:|$)", t, re.I)
    if m:
        out["Com"] = to_number_str(m.group(1))

    m = re.search(r"\bCOL\s*:\s*(.+?)(?=\s*\bCOM\s*:|\s*\bCUSHION\s*:|$)", t, re.I)
    if m:
        out["Col"] = to_number_str(m.group(1))

    m = re.search(r"\bCUSHION\s*:\s*(.+?)(?=\s*\bCOM\s*:|\s*\bCOL\s*:|$)", t, re.I)
    if m:
        out["Cushion"] = clean_text(m.group(1))

    return out

# -----------------------------
# Scrape one page
# -----------------------------
def scrape_one(url: str, product_name: str) -> dict:
    result = {
        "SKU": "",
        "Product Family Id": product_name or "",
        "Description": "",
        "Dimension": "",
        "UPHOLSTERY OPTIONS": "",
        "Com": "",
        "Col": "",
        "Cushion": "",
        "Weight": "",
        "Width": "",
        "Depth": "",
        "Diameter": "",
        "Height": "",
        "Seat Height": "",
        "Seat Depth": "",
        "Arm Height": "",
        "Shade Details": "",
    }

    if not url or not isinstance(url, str):
        return result

    soup = fetch_soup(url)
    if not soup:
        return result

    summary = soup.select_one("div.summary.entry-summary")
    if not summary:
        return result

    result["Description"] = extract_description(summary)
    result["SKU"] = extract_sku(summary)

    dim_text = extract_dimension_block_text(summary)
    result["Dimension"] = dim_text
    result.update(parse_dimension_fields(dim_text))

    up_text = extract_upholstery_options(summary)
    result["UPHOLSTERY OPTIONS"] = up_text
    result.update(parse_upholstery_fields(up_text))

    return result

# -----------------------------
# Main
# -----------------------------
def main():
    df = pd.read_excel(INPUT_EXCEL)

    # Required input columns
    for col in ["Product URL", "Image URL", "Product Name"]:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    # Ensure all output columns exist
    out_cols = [
        "SKU", "Product Family Id", "Description",
        "Weight", "Width", "Depth", "Diameter", "Height",
        "Seat Height", "Seat Depth", "Arm Height", "Shade Details",
        "Com", "Col", "Cushion",
        "Dimension", "UPHOLSTERY OPTIONS"
    ]
    for c in out_cols:
        if c not in df.columns:
            df[c] = ""

    # Scrape loop
    for i, row in df.iterrows():
        url = str(row.get("Product URL", "")).strip()
        pname = str(row.get("Product Name", "")).strip()

        data = scrape_one(url, pname)
        for k, v in data.items():
            df.at[i, k] = v

        if (i + 1) % 20 == 0:
            print(f"Processed {i+1}/{len(df)}")

    # Final column order (your requirement)
    final_order = [
        "Product URL", "Image URL", "Product Name",
        "SKU", "Product Family Id", "Description",
        "Weight", "Width", "Depth", "Diameter", "Height",
        "Seat Height", "Seat Depth", "Arm Height", "Shade Details",
        "Com", "Col", "Cushion",
        "Dimension", "UPHOLSTERY OPTIONS"
    ]

    for c in final_order:
        if c not in df.columns:
            df[c] = ""

    remaining = [c for c in df.columns if c not in final_order]
    df = df[final_order + remaining]

    df.to_excel(OUTPUT_EXCEL, index=False)
    print("\nDONE ✅")
    print("Saved:", OUTPUT_EXCEL)

if __name__ == "__main__":
    main()
