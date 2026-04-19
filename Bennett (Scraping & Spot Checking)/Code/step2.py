import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import time
import re

# 🏠 Base URL
base_domain = "https://www.bennetttothetrade.com"

# 📁 File paths
script_dir = os.path.dirname(os.path.abspath(__file__))
source_path = os.path.join(script_dir, "bennet_sideboards.xlsx")
final_path = os.path.join(script_dir, "bennet_sideboards_final.xlsx")

# ================================
# STEP 1: SCRAPING PRODUCT DETAILS
# ================================

print("=" * 60)
print("STEP 1: SCRAPING PRODUCT DETAILS")
print("=" * 60)

# ✅ Required columns
columns = [
    "Product URL", "Image URL", "Product Name", "SKU",
    "Description", "Weight", "Product Family Id", "List Price"
]

# 🔹 Load previous progress or create new
if os.path.exists(final_path):
    df = pd.read_excel(final_path)
    print(f"📂 Previous progress loaded: {len(df)} rows found")
else:
    source_df = pd.read_excel(source_path)
    df = source_df.copy()
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df.to_excel(final_path, index=False)
    print(f"🆕 New file created from source.")

# 🔹 Get pending products
pending_df = df[df["Description"].isna() | (df["Description"] == "")]
print(f"🔎 {len(pending_df)} products pending to scrape.\n")


# ✅ Extract product details
def extract_details(soup):
    desc_block = soup.select_one("div.product__description.rte.quick-add-hidden")
    description = ""
    weight = ""
    product_family_id = ""
    list_price = ""

    # Extract Product Family Id
    name_tag = soup.find("h1")
    if name_tag:
        product_family_id = name_tag.get_text(strip=True)

    # Extract Description
    if desc_block:
        description = "\n".join([p.get_text(separator="\n", strip=True) for p in desc_block.find_all("p")])

    # Extract List Price (try multiple selectors and extract only numbers)
    price_text = ""

    # Try multiple possible price selectors
    price_selectors = [
        "span.price-item.price-item--regular",
        ".price-item--regular",
        "#price-template--17007414902996__main .price__container .price__regular .price-item--regular",
        ".price__regular .price-item"
    ]

    for selector in price_selectors:
        price_block = soup.select_one(selector)
        if price_block:
            price_text = price_block.get_text(strip=True)
            break

    # Extract only numbers from price (remove $, USD, commas, etc.)
    if price_text:
        # Remove currency symbols, "USD", and extract numbers with decimal
        price_match = re.search(r'([0-9,]+\.?[0-9]*)', price_text)
        if price_match:
            list_price = price_match.group(1).replace(',', '')  # Remove commas from number

    return description, weight, product_family_id, list_price


# --- Scraping loop ---
batch_size = 5
processed = 0
total = len(pending_df)

for index, row in pending_df.iterrows():
    product_url = row["Product URL"]
    if not isinstance(product_url, str) or not product_url.startswith("http"):
        continue

    processed += 1
    print(f"🔎 [{processed}/{total}] Scraping: {product_url}")

    try:
        response = requests.get(product_url, headers={"User-Agent": "Mozilla/5.0"})
        if response.status_code != 200:
            print(f"⚠️ Failed to fetch: {product_url}")
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        desc, weight, product_family_id, list_price = extract_details(soup)

        # Update dataframe
        df.at[index, "Description"] = desc
        df.at[index, "Weight"] = weight
        df.at[index, "Product Family Id"] = product_family_id
        df.at[index, "List Price"] = list_price

        print(f"✅ Done: {product_url}")

        # Auto-save every 5 products
        if processed % batch_size == 0:
            df.to_excel(final_path, index=False)
            print(f"💾 Auto-saved after {processed} products...\n")

        time.sleep(1.5)

    except Exception as e:
        print(f"❌ Error processing {product_url}: {e}")
        continue

# 🔹 Save scraping results
df.to_excel(final_path, index=False)
print(f"\n✅ Scraping completed and saved!")
print(f"📁 File location: {final_path}\n")

# ================================
# STEP 2: EXTRACT DIMENSIONS
# ================================

print("=" * 60)
print("STEP 2: EXTRACTING DIMENSIONS")
print("=" * 60)

# ✅ Ensure dimension columns exist
for col in ["Length", "Width", "Depth", "Diameter", "Height", "Seat Height", "COM"]:
    if col not in df.columns:
        df[col] = ""


# ✅ Dimension extraction function
def extract_dimensions(text):
    length = width = depth = diameter = height = seat_height = com = ""

    if not isinstance(text, str):
        return length, width, depth, diameter, height, seat_height, com

    # Normalize text
    text = text.replace("â€³", '"').replace("Ã—", "x").replace("  ", " ").strip()

    # --- Extract COM FIRST and REMOVE from text to avoid interference ---
    com_match = re.search(r'COM-?\s*([0-9\.]+)\s*(?:yards?)?', text, re.IGNORECASE)
    if com_match:
        com = com_match.group(1)
        # Remove COM from text so it doesn't get picked up as a dimension
        text = re.sub(r'COM-?\s*[0-9\.]+\s*(?:yards?)?', '', text, flags=re.IGNORECASE)

    # --- Extract Seat Height (multiple patterns) ---
    # Pattern 1: "to seat" or "to Seat" - e.g., (19" to seat) or 17.25"H to Seat
    seat_match1 = re.search(r'(?:\()?([0-9\.]+)"?\s*(?:H)?\s*to\s+[Ss]eat', text, re.IGNORECASE)
    if seat_match1:
        seat_height = seat_match1.group(1)

    # Pattern 2: "Seat Height:" format
    if not seat_height:
        seat_match2 = re.search(r'Seat\s+Height[:\s]*([0-9\.]+)"?', text, re.IGNORECASE)
        if seat_match2:
            seat_height = seat_match2.group(1)

    # Pattern 3: Inside parentheses at end of dimension string - e.g., 20"x18.5"x32" (19" to seat)
    if not seat_height:
        seat_match3 = re.search(r'\(([0-9\.]+)"?\s+to\s+seat\)', text, re.IGNORECASE)
        if seat_match3:
            seat_height = seat_match3.group(1)

    # Remove seat height patterns from text to avoid interference
    if seat_height:
        text = re.sub(r'(?:\()?[0-9\.]+"?\s*(?:H)?\s*to\s+[Ss]eat(?:\))?', '', text, flags=re.IGNORECASE)
        text = re.sub(r'Seat\s+Height[:\s]*[0-9\.]+"?', '', text, flags=re.IGNORECASE)

    # --- Extract dimensions like 66"x81"x70.5"H or 20"x18.5"x32" ---
    if 'x' in text.lower():
        # First extract the main dimension part (before any parentheses or extra info)
        dim_part = re.search(r'([0-9\.]+)"?\s*x\s*([0-9\.]+)"?\s*x\s*([0-9\.]+)"?\s*([HhDd])?', text)
        if dim_part:
            length = dim_part.group(1)
            width = dim_part.group(2)
            height_or_depth = dim_part.group(3)
            label = dim_part.group(4)

            # If labeled with H, it's height
            if label and label.upper() == 'H':
                height = height_or_depth
            # If labeled with D, it's depth
            elif label and label.upper() == 'D':
                depth = height_or_depth
            # Otherwise, assume third dimension is height
            else:
                height = height_or_depth

    # --- General dimension extraction for labeled values ---
    matches = re.findall(r'([0-9\.]+)\s*"?\s*([A-Za-z]*)', text)
    for num, label in matches:
        label = label.upper().strip()
        if "H" in label and not height:
            height = num
        elif "W" in label and not width:
            width = num
        elif "L" in label and not length:
            length = num
        elif "DIA" in label:
            diameter = num
        elif "D" in label and not depth:
            depth = num

    # --- Fallback for Length if still empty ---
    if not length:
        all_nums = re.findall(r'([0-9]+\.?[0-9]*)', text)
        if all_nums:
            length = all_nums[-1]

    # --- Ensure all values contain only numeric part ---
    length = re.sub(r'[^\d.]+', '', length) if length else ""
    width = re.sub(r'[^\d.]+', '', width) if width else ""
    depth = re.sub(r'[^\d.]+', '', depth) if depth else ""
    height = re.sub(r'[^\d.]+', '', height) if height else ""
    seat_height = re.sub(r'[^\d.]+', '', seat_height) if seat_height else ""
    diameter = re.sub(r'[^\d.]+', '', diameter) if diameter else ""
    com = re.sub(r'[^\d.]+', '', com) if com else ""

    return length, width, depth, diameter, height, seat_height, com


# ✅ Apply extraction to each row
print(f"🔎 Extracting dimensions from {len(df)} products...\n")
for index, row in df.iterrows():
    description = row.get("Description", "")
    length, width, depth, diameter, height, seat_height, com = extract_dimensions(description)

    df.at[index, "Length"] = length
    df.at[index, "Width"] = width
    df.at[index, "Depth"] = depth
    df.at[index, "Diameter"] = diameter
    df.at[index, "Height"] = height
    df.at[index, "Seat Height"] = seat_height
    df.at[index, "COM"] = com

# ✅ Reorder columns
column_order = [
    "Product URL", "Image URL", "Product Name", "SKU", "Product Family Id",
    "Description", "List Price", "Weight",
    "Width", "Depth", "Diameter", "Length", "Height", "Seat Height", "COM"
]

# Ensure all columns exist
for col in column_order:
    if col not in df.columns:
        df[col] = ""

df = df[column_order]

# ✅ Save final file
df.to_excel(final_path, index=False)
print("✅ All dimensions extracted and saved successfully!")
print(f"📁 Final file: {final_path}")
print("\n" + "=" * 60)
print("🎉 PROCESS COMPLETED!")
print("=" * 60)
print(f"📌 Source file: {source_path}")
print(f"📌 Final file: {final_path}")