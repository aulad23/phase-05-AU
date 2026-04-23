"""
SKILL: Data Extraction
Text parsing patterns: SKU, dimensions, descriptions, prices.
"""

import re
import json
from bs4 import BeautifulSoup


# ── SKU EXTRACTION ────────────────────────────────────────────────────────────
SKU_PATTERNS = [
    r"(?:SKU|Item #?|Product #?|Model #?|Item No\.?|Style #?)\s*:?\s*([A-Z0-9\-_]{3,30})",
    r"(?:sku|item_id|product_id)\s*[\"']?\s*:\s*[\"']?([A-Z0-9\-_]{3,30})",
]

def extract_sku(text: str) -> str:
    for pat in SKU_PATTERNS:
        m = re.search(pat, text or "", re.I)
        if m:
            return m.group(1).strip()
    return ""


# ── DIMENSION LABEL PARSER ────────────────────────────────────────────────────
LABEL_MAP = {
    "width": "Width", "w": "Width",
    "depth": "Depth",  "d": "Depth",
    "height": "Height", "h": "Height",
    "length": "Length", "l": "Length",
    "diameter": "Diameter", "dia": "Diameter",
    "weight": "Weight", "wt": "Weight",
    "seat width": "Seat Width",  "sw": "Seat Width",
    "seat depth": "Seat Depth",  "sd": "Seat Depth",
    "seat height": "Seat Height","sh": "Seat Height",
    "arm height": "Arm Height",  "ah": "Arm Height",
    "arm width":  "Arm Width",   "aw": "Arm Width",
}

def parse_spec_table(soup_or_text, label_selector: str = None) -> dict:
    """
    Parse a spec table or label:value text block.
    Returns dict of dimension fields.
    """
    out = {v: "" for v in set(LABEL_MAP.values())}

    if isinstance(soup_or_text, str):
        lines = soup_or_text.splitlines()
        for line in lines:
            if ":" in line:
                key, _, val = line.partition(":")
                mapped = LABEL_MAP.get(key.strip().lower())
                if mapped:
                    out[mapped] = val.strip()
        return out

    # BeautifulSoup element
    rows = soup_or_text.find_all("tr") if soup_or_text else []
    for row in rows:
        cells = row.find_all(["td", "th"])
        if len(cells) >= 2:
            key = cells[0].get_text(strip=True).lower()
            val = cells[1].get_text(strip=True)
            mapped = LABEL_MAP.get(key)
            if mapped:
                out[mapped] = val
    return out


# ── SHOPIFY JSON EXTRACTION ───────────────────────────────────────────────────
def extract_shopify_json(page_html: str) -> dict:
    """Extract product JSON from Shopify pages (window.ShopifyAnalytics or JSON-LD)."""
    # Try window.__st or ShopifyAnalytics
    m = re.search(r"var\s+meta\s*=\s*(\{.*?\});", page_html, re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass

    # JSON-LD
    soup = BeautifulSoup(page_html, "html.parser")
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
            if isinstance(data, dict) and data.get("@type") in ("Product", "product"):
                return data
        except Exception:
            pass
    return {}


# ── DESCRIPTION CLEANER ───────────────────────────────────────────────────────
_STOP_MARKERS = [
    "dimensions", "specifications", "materials", "shipping",
    "delivery", "return policy", "care instructions",
]

def extract_description(element, max_chars: int = 800) -> str:
    """Get clean product description, stop at spec sections."""
    if element is None:
        return ""
    text = element.get_text(separator=" ", strip=True)
    lower = text.lower()
    for marker in _STOP_MARKERS:
        idx = lower.find(marker)
        if idx > 50:
            text = text[:idx]
            break
    return re.sub(r"\s+", " ", text).strip()[:max_chars]


# ── PRICE PARSER ──────────────────────────────────────────────────────────────
def extract_price(text: str) -> str:
    """Extract numeric price string from text like '$1,234.00'."""
    m = re.search(r"\$?([\d,]+\.?\d{0,2})", str(text or "").replace(",", ""))
    return m.group(1) if m else ""
