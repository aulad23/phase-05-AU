"""
SKILL: Python Developer Utilities
Common Python patterns used across all vendor scripts.
"""

import re
import time
import random
import logging
from functools import wraps
from fractions import Fraction


# ── FRACTION → DECIMAL ────────────────────────────────────────────────────────
def frac_to_decimal(text: str) -> str:
    """Convert '26 3/4 inches' → '26.75'"""
    if not text:
        return ""
    text = str(text).strip()

    def replace_frac(m):
        whole = int(m.group(1) or 0)
        num, den = int(m.group(2)), int(m.group(3))
        return str(round(whole + num / den, 4))

    result = re.sub(r"(\d+)?\s*(\d+)/(\d+)", replace_frac, text)
    result = re.sub(r"['""]|inches?|cm|mm|ft|lbs?|kg", "", result, flags=re.I).strip()
    try:
        return str(round(float(result), 2))
    except ValueError:
        return result


# ── DIMENSION EXTRACTOR ───────────────────────────────────────────────────────
_DIM_PATTERNS = [
    r"(\d[\d\s/]*)\s*[\"'']?\s*[Ww](?:idth)?",
    r"(\d[\d\s/]*)\s*[\"'']?\s*[Dd](?:epth|iam)?",
    r"(\d[\d\s/]*)\s*[\"'']?\s*[Hh](?:eight|t)?",
    r"(\d[\d\s/]*)\s*[\"'']?\s*[Ll](?:ength|en)?",
    r"[Dd]ia(?:meter)?\s*:?\s*(\d[\d\s/]*)",
    r"[Ww]eight\s*:?\s*([\d\s/]+)\s*(?:lbs?|kg)?",
]

def extract_dims(text: str) -> dict:
    """Extract W/D/H/L/Dia/Weight from a dimension string."""
    out = {"Width": "", "Depth": "", "Height": "", "Length": "", "Diameter": "", "Weight": ""}
    keys = ["Width", "Depth", "Height", "Length", "Diameter", "Weight"]
    for key, pat in zip(keys, _DIM_PATTERNS):
        m = re.search(pat, text or "")
        if m:
            out[key] = frac_to_decimal(m.group(1))
    return out


# ── RETRY DECORATOR ───────────────────────────────────────────────────────────
def retry(times=3, delay=2, exceptions=(Exception,)):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            for attempt in range(1, times + 1):
                try:
                    return fn(*args, **kwargs)
                except exceptions as e:
                    if attempt == times:
                        raise
                    wait = delay * attempt + random.uniform(0, 1)
                    print(f"  Retry {attempt}/{times} after {wait:.1f}s — {e}")
                    time.sleep(wait)
        return wrapper
    return decorator


# ── RANDOM DELAY ──────────────────────────────────────────────────────────────
def polite_sleep(lo=1.0, hi=3.0):
    time.sleep(random.uniform(lo, hi))


# ── LOGGING SETUP ─────────────────────────────────────────────────────────────
def get_logger(name: str, level=logging.INFO) -> logging.Logger:
    logging.basicConfig(
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        level=level,
    )
    return logging.getLogger(name)


# ── CLEAN TEXT ────────────────────────────────────────────────────────────────
def clean_text(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


# ── PRICE CLEANER ─────────────────────────────────────────────────────────────
def clean_price(text: str) -> str:
    if not text:
        return ""
    m = re.search(r"[\d,]+\.?\d*", str(text).replace(",", ""))
    return m.group() if m else ""
