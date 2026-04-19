import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import re


# ---------------- DRIVER SETUP ----------------
def setup_driver():
    options = Options()
    # options.add_argument("--headless=new")
    service = Service(r"C:\chromedriver-win64\chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)
    return driver


# ---------------- CM ? INCH CONVERTER ----------------
def cm_to_inch(dim_cm: str) -> str:
    """
    Convert dimension string from cm to inches.
    Handles:
    - European comma: 31,5 -> 31.5
    - "Diam. 150x36 cm" -> "Diam. 59.06 x 14.17 inch" (preserves "Diam.")
    - "LxPxHt 160x42x42 cm" -> "160x42x42 cm" (removes "LxPxHt")
    - "Structure - Diam. 29 x 146,5 cm?Base - Diam. 55/50 x 60 cm" -> extracts only Structure part
    - "55/50" -> takes first number (55)
    - Different separators
    Returns: "31.5 x 15.94 x 90.55 inch"
    """
    try:
        if not dim_cm or str(dim_cm).strip() == "":
            return ""

        s = str(dim_cm).strip()

        # ---------- SPECIAL CASE: Extract only "Structure" part if multiple parts exist ----------
        # Check for pipe separator (? or |) and "Structure" keyword
        if ('?' in s or '|' in s) and 'structure' in s.lower():
            # Split by pipe (both types)
            parts = re.split(r'[?|]', s)
            # Find the part containing "Structure"
            for part in parts:
                if 'structure' in part.lower():
                    # Extract dimension after "Structure -"
                    structure_match = re.search(r'structure\s*-\s*(.+)', part, flags=re.IGNORECASE)
                    if structure_match:
                        s = structure_match.group(1).strip()
                        break

        # Check if starts with "Diam." / "Diameter"
        diam_prefix = ""
        diam_match = re.match(r'^(Diam\.?|DIAM\.?|Diameter)\s*', s, flags=re.IGNORECASE)
        if diam_match:
            diam_prefix = "Diam. "  # Store prefix
            s = s[diam_match.end():].strip()  # Remove from string

        # Remove "LxPxHt" / "LxDxH" / similar prefixes (case insensitive)
        s = re.sub(r'^[LlWw]x[PpDd]x[HhTt]+\s*', '', s).strip()

        # Replace European comma with dot
        s = s.replace(",", ".")

        # Remove "cm" (any case)
        s = re.sub(r'\s*cm\s*', '', s, flags=re.IGNORECASE).strip()

        # Split by x (allow X too)
        parts = re.split(r'\s*[xX]\s*', s)
        parts = [p.strip() for p in parts if p.strip()]

        # Convert each numeric part
        inches = []
        for p in parts:
            # Handle slash in numbers (e.g., "55/50") - take first number
            if '/' in p:
                p = p.split('/')[0].strip()

            # Extract number (properly handles decimals like "146.5")
            # Pattern: one or more digits, optionally followed by dot and more digits
            num_match = re.search(r'\d+\.?\d*', p)
            if not num_match:
                return ""  # if conversion uncertain, return blank

            num_str = num_match.group(0)
            inches.append(round(float(num_str) / 2.54, 2))

        result = " x ".join(map(str, inches)) + " inch"

        # Re-add "Diam." prefix if it existed
        if diam_prefix:
            result = diam_prefix + result

        return result

    except Exception as e:
        print(f"   ?? Dimension conversion error: '{dim_cm}' -> {e}")
        return ""


# ---------------- PARSE DIMENSIONS ----------------
def parse_dimensions(dimension_str: str):
    """
    Extract: Weight, Width, Depth, Diameter, Height from Dimension.
    Rules:
    - If weight unit exists (kg/lb/lbs/ibs/ib) => keep Weight blank (do not store value)
    - Special case: "Diam. 150x36 inch" => Diameter=150, Height=36
    - Unlabeled:
        3 values => W, D, H
        2 values => W, H
    - Labeled:
        D => Depth, W => Width, H/HT => Height
        DIA/DIAM/DIAMETER => Diameter
    """
    result = {
        "weight": "",
        "width": "",
        "depth": "",
        "diameter": "",
        "height": ""
    }

    if not dimension_str or str(dimension_str).strip() in ["", "N/A"]:
        return result

    s = str(dimension_str).strip()

    # Remove trailing "inch" if present
    s = re.sub(r'\s*inch\s*$', '', s, flags=re.IGNORECASE).strip()

    # ---------- 1) WEIGHT: if present, DO NOT STORE VALUE (keep blank), just remove from string ----------
    weight_pattern = r'([\d.]+)\s*(kg|lbs?|lb|ibs?|ib)\b'
    if re.search(weight_pattern, s, flags=re.IGNORECASE):
        s = re.sub(weight_pattern, '', s, flags=re.IGNORECASE).strip()
        s = s.strip('xX').strip()

    # ---------- 2) SPECIAL CASE: "Diam. 150x36" or "Diameter 150 x 36" ----------
    diam_pattern = r'^(Diam\.?|DIAM\.?|Diameter)\s*([\d.]+)\s*[xX]?\s*([\d.]+)?'
    diam_match = re.match(diam_pattern, s, flags=re.IGNORECASE)

    if diam_match:
        result["diameter"] = diam_match.group(2).strip()
        if diam_match.group(3):  # If there's a second number (height)
            result["height"] = diam_match.group(3).strip()
        return result

    # ---------- 3) Detect labeled dims ----------
    has_labels = bool(re.search(r'[\d.]+\s*(DIA|DIAM|DIAMETER|D|W|H|HT)\b', s, flags=re.IGNORECASE))

    if has_labels:
        parts = re.split(r'\s*[xX]\s*', s)
        parts = [p.strip() for p in parts if p.strip()]

        for part in parts:
            m = re.match(r'([\d.]+)\s*([a-zA-Z.]+)', part)
            if not m:
                continue

            value = m.group(1).strip()
            label = m.group(2).strip().upper().replace(".", "")

            # Normalize common diameter labels
            if label in ["DIA", "DIAM", "DIAMETER"]:
                result["diameter"] = value
            elif label in ["D", "DEPTH"]:
                result["depth"] = value
            elif label in ["W", "WIDTH"]:
                result["width"] = value
            elif label in ["H", "HEIGHT", "HT"]:
                result["height"] = value

        return result

    # ---------- 4) Unlabeled (position-based) ----------
    nums = re.findall(r'\d+\.?\d*', s)
    if len(nums) >= 3:
        result["width"] = nums[0]
        result["depth"] = nums[1]
        result["height"] = nums[2]
    elif len(nums) == 2:
        result["width"] = nums[0]
        result["height"] = nums[1]
    elif len(nums) == 1:
        result["width"] = nums[0]

    return result


# ---------------- SCRAPE PRODUCT PAGE ----------------
def scrape_product_details(driver, url):
    driver.get(url)
    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # Description
    desc_tag = soup.select_one(".product-content__description p")
    description = desc_tag.get_text(strip=True) if desc_tag else ""

    # Dimension (from cm -> inch)
    dimension = ""
    dim_tag = soup.select_one(".product-content__size__description p")
    if dim_tag:
        dim_text = dim_tag.get_text(strip=True)
        dimension = cm_to_inch(dim_text)

    if not dimension:
        dimension = "N/A"

    return description, dimension


# ---------------- MAIN ----------------
if __name__ == "__main__":
    INPUT_FILE = "studioliaigre_Objects.xlsx"
    OUTPUT_FILE = "studioliaigre_Objects_STEP2.xlsx"

    df = pd.read_excel(INPUT_FILE)

    driver = setup_driver()

    data_rows = []

    VENDOR_CODE = "LIA"
    CATEGORY_CODE = "Ob"

    try:
        for index, row in df.iterrows():
            print(f"\n{'=' * 70}")
            print(f"[{index + 1}/{len(df)}] {row['Product Name']}")
            print(f"URL: {row['Product URL']}")

            description, dimension = scrape_product_details(driver, row["Product URL"])

            sku = f"{VENDOR_CODE}{CATEGORY_CODE}{str(index + 1).zfill(3)}"

            parsed = parse_dimensions(dimension)

            row_data = {
                "Product URL": row["Product URL"],
                "Image URL": row["Image URL"],
                "Product Name": row["Product Name"],
                "SKU": sku,
                "Product Family Id": row["Product Name"],
                "Description": description,
                "Weight": parsed["weight"],  # will remain blank if kg/lb etc exists
                "Width": parsed["width"],
                "Depth": parsed["depth"],
                "Diameter": parsed["diameter"],
                "Height": parsed["height"],
                "Dimension": dimension
            }

            data_rows.append(row_data)

            print(f"   ? SKU: {sku}")
            print(f"   ? Dimension: {dimension}")
            print(
                f"   ? Parsed: W={parsed['width']} D={parsed['depth']} DIA={parsed['diameter']} H={parsed['height']} Weight='{parsed['weight']}'")

    except Exception as e:
        print(f"\n? Error: {e}")
        import traceback

        traceback.print_exc()
    finally:
        driver.quit()

    # Final column order (exactly as you requested)
    FINAL_COLS = [
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
        "Height",
        "Dimension",
    ]

    final_df = pd.DataFrame(data_rows)
    final_df = final_df.reindex(columns=FINAL_COLS)
    final_df.to_excel(OUTPUT_FILE, index=False)

    print(f"\n{'=' * 70}")
    print(f"? COMPLETED! Saved to: {OUTPUT_FILE}")
    print(f"{'=' * 70}\n")