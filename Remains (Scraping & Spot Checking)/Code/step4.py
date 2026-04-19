import re
import pandas as pd
from pathlib import Path

# -------- CONFIG --------
IN_XLSX  = "remains_chandeliers_file.xlsx"
OUT_XLSX = "remains_chandeliers_final.xlsx"

# Number pattern
NUM = r'(?:\d+(?:\.\d+)?(?:-\d+/\d+)?|\d+/\d+)'

# Regex patterns
TOKEN_RE = re.compile(rf'(?P<val>{NUM})\s*"?\s*(?P<label>h|w|d)\.?\b', re.I)
DIAM_WORD_RE = re.compile(rf'(?P<val>{NUM})\s*"?\s*(?:diameter|dia)\b', re.I)
DEPTH_WORD_RE = re.compile(rf'(?P<val>{NUM})\s*"?\s*(?:depth|deep)\b', re.I)

CUR_HEIGHT_RE = re.compile(rf'^\s*current\s+height\s*:\s*(?P<val>{NUM})', re.I)
MIN_HEIGHT_RE = re.compile(rf'^\s*min(?:imum)?\s+height\s*:\s*(?P<val>{NUM})', re.I)

CANOPY_RE = re.compile(rf'^\s*canopy\s*:\s*(?P<val>{NUM})', re.I)
DEPTH_RE = re.compile(rf'^\s*depth\s*:\s*(?P<val>{NUM})', re.I)
DIAMETER_RE = re.compile(rf'^\s*diameter\s*:\s*(?P<val>{NUM})', re.I)
BACKPLATE_RE = re.compile(rf'^\s*backplate\s*:\s*(?P<val>{NUM})', re.I)
BASE_RE = re.compile(rf'^\s*base\s*:\s*(?P<val>{NUM})', re.I)


# ---------- Helper functions ----------

def clean_num(s: str) -> str:
    return (s or "").replace('"', '').strip()


def pick_height_from_lines(dimensions_text: str) -> str:
    """Prefer Current height, else Minimum height."""
    if not isinstance(dimensions_text, str) or not dimensions_text.strip():
        return ""
    for raw in dimensions_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = CUR_HEIGHT_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    for raw in dimensions_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = MIN_HEIGHT_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    return ""


def parse_overall(dimensions_text: str):
    """
    Parse 'Overall:' line into (height, width, depth, diameter)
    Works even if diameter appears inside the same line.
    """
    h = w = depth = diam = ""
    if not isinstance(dimensions_text, str) or not dimensions_text.strip():
        return h, w, depth, diam

    for raw in dimensions_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw.strip().strip('"')
        # only take if line truly starts with "Overall:"
        if not re.match(r"^\s*overall\s*:", line, re.I):
            continue

        core = line.split(":", 1)[1].strip()

        # Case 1: tokens like 6" h. / 10" w. / 5" d.
        for m in TOKEN_RE.finditer(core):
            val = clean_num(m.group("val"))
            lab = m.group("label").lower()
            if lab == "h" and not h:
                h = val
            elif lab == "w" and not w:
                w = val
            elif lab == "d" and not depth:
                depth = val

        # Case 2: "15 diameter" or "12 dia"
        if not diam:
            m2 = DIAM_WORD_RE.search(core)
            if m2:
                diam = clean_num(m2.group("val"))

        # Case 3: "4.5 depth"
        if not depth:
            m3 = DEPTH_WORD_RE.search(core)
            if m3:
                depth = clean_num(m3.group("val"))

        break  # only first true Overall line
    return h, w, depth, diam


def pick_canopy_from_lines(dimensions_text: str) -> str:
    if not isinstance(dimensions_text, str) or not dimensions_text.strip():
        return ""
    for raw in dimensions_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = CANOPY_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    return ""


def pick_depth_from_lines(dimensions_text: str) -> str:
    if not isinstance(dimensions_text, str) or not dimensions_text.strip():
        return ""
    for raw in dimensions_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = DEPTH_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    return ""


def pick_diameter_from_lines(dimensions_text: str) -> str:
    if not isinstance(dimensions_text, str) or not dimensions_text.strip():
        return ""
    for raw in dimensions_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = DIAMETER_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    return ""


def pick_backplate_from_specs(specs_text: str) -> str:
    """Extract 'Backplate' value from Specifications column"""
    if not isinstance(specs_text, str) or not specs_text.strip():
        return ""
    for raw in specs_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = BACKPLATE_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    return ""


def pick_base_from_specs(specs_text: str) -> str:
    """Extract 'Base' value from Specifications column"""
    if not isinstance(specs_text, str) or not specs_text.strip():
        return ""
    for raw in specs_text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = BASE_RE.search(raw)
        if m:
            return clean_num(m.group("val"))
    return ""


# ---------- Main ----------

def main():
    if not Path(IN_XLSX).exists():
        raise FileNotFoundError(f"Input file not found: {IN_XLSX}")

    df = pd.read_excel(IN_XLSX)
    if "Dimensions" not in df.columns:
        raise KeyError('Expected a "Dimensions" column in the input file.')

    out_widths, out_depths, out_diams, out_heights, out_canopy, out_backplate, out_base = [], [], [], [], [], [], []

    for _, row in df.iterrows():
        dims = str(row.get("Dimensions", ""))
        specs = str(row.get("Specifications", ""))

        height = pick_height_from_lines(dims)
        overall_h, overall_w, overall_depth, overall_diam = parse_overall(dims)

        if not height:
            height = overall_h

        canopy = pick_canopy_from_lines(dims)
        depth = overall_depth or pick_depth_from_lines(dims)
        diam = overall_diam or pick_diameter_from_lines(dims)
        backplate = pick_backplate_from_specs(specs)
        base = pick_base_from_specs(specs)

        out_widths.append(overall_w)
        out_depths.append(depth)
        out_diams.append(diam)
        out_heights.append(height)
        out_canopy.append(canopy)
        out_backplate.append(backplate)
        out_base.append(base)

    # ---- final output columns ----
    df["Width"] = out_widths
    df["Depth"] = out_depths
    df["Diameter"] = out_diams
    df["Height"] = out_heights
    df["Canopy"] = out_canopy
    df["Backplate"] = out_backplate
    df["Base"] = out_base

    df.to_excel(OUT_XLSX, index=False)
    print(f"✅ Done. Saved {len(df)} rows → {OUT_XLSX}")


if __name__ == "__main__":
    main()
