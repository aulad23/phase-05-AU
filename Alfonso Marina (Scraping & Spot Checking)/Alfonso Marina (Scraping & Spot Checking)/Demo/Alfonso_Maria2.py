from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from fractions import Fraction

# --------------------------------------------------------
# Convert Fraction / Mixed fraction to Decimal
# --------------------------------------------------------
def convert_fraction_to_decimal(value):
    if not value:
        return ""
    try:
        value = value.replace('"', '').strip()

        # If value is a dash (–, -, —) or empty → return empty
        if value in ("–", "-", "—", ""):
            return ""

        # Mixed fraction example: "26 3/4" or "19-3/4"
        if re.match(r'^\d+[-\s]\d+/\d+$', value):
            whole, frac = re.split(r'[-\s]', value, maxsplit=1)
            return str(float(whole) + float(Fraction(frac)))

        # Simple fraction: 1/2
        if re.match(r'^\d+/\d+$', value):
            return str(float(Fraction(value)))

        # Number (integer or decimal)
        if re.match(r'^\d+(\.\d+)?$', value):
            return value

    except:
        pass

    return ""

# --------------------------------------------------------
# Selenium Setup
# --------------------------------------------------------
chrome_options = Options()
# chrome_options.add_argument("--headless")  # uncomment to run headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)
driver.set_page_load_timeout(60)
wait = WebDriverWait(driver, 30)

# --------------------------------------------------------
# Input / Output Files
# --------------------------------------------------------
input_path = "alfonsomarina_Boxes.xlsx"
df_input = pd.read_excel(input_path)

final_path = "alfonso_Boxes_final.xlsx"

# Create empty sheet with Finish column
columns = [
    "Product URL", "Image URL", "Product Name", "SKU",
    "Product Family Id", "Description", "Weight", "Width",
    "Depth", "Diameter", "Height", "Finish"
]
pd.DataFrame(columns=columns).to_excel(final_path, index=False)

# --------------------------------------------------------
# Normalize column names
# --------------------------------------------------------
df_input.columns = [c.strip().replace(" ", "_").lower() for c in df_input.columns]

# --------------------------------------------------------
# Start Scraping
# --------------------------------------------------------
total = len(df_input)
print(f"🟢 Starting scraping for {total} products...\n")

all_rows = []

for idx, row in df_input.iterrows():
    product_url = row.get("product_url", "")
    product_name = str(row.get("product_name", "")).strip()
    image_url = row.get("image_url", "")

    if not product_url:
        print(f"[{idx + 1}/{total}] Skipping row, Product URL missing")
        continue

    # First letter upper, rest lower → "LILLE CIRCULAR LAMP TABLE" → "Lille circular lamp table"
    product_name = product_name.capitalize()
    # Product Family Id = everything before first - , . or _
    product_family_id = re.split(r'[-,._]', product_name, maxsplit=1)[0].strip()

    try:
        driver.get(product_url)
        # Wait for main content
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "elementor-widget-container")))
        time.sleep(1)
        soup = BeautifulSoup(driver.page_source, "html.parser")
    except Exception as e:
        print(f"⚠️ Failed to load {product_url}: {e}")
        continue

    # -------------------------------
    # SKU
    # -------------------------------
    sku_tag = soup.find(string=re.compile("PRODUCT CODE"))
    sku = sku_tag.split(":")[-1].strip() if sku_tag else ""

    # -------------------------------
    # Dimensions (INCHES section only)
    # -------------------------------
    Weight = Width = Depth = Height = Diameter = ""

    # Find "DIMENSIONS IN" text to locate the inches section
    dim_tags = soup.find_all("div", class_="elementor-widget-container")

    in_section = False  # Flag to track if we're in the INCHES section

    for tag in dim_tags:
        text = tag.get_text(strip=True)

        # Detect start of INCHES section
        if "DIMENSIONS IN" in text:
            in_section = True
            continue

        # Detect start of next section (stop reading dimensions)
        if in_section and ("Character" in text or "Lead Time" in text or "As Shown" in text):
            in_section = False
            break

        # Only parse dimensions from INCHES section
        if in_section:
            if text.startswith("W:"):
                raw = text.replace("W:", "").strip()
                Width = convert_fraction_to_decimal(raw)
            elif text.startswith("D:"):
                raw = text.replace("D:", "").strip()
                Depth = convert_fraction_to_decimal(raw)
            elif text.startswith("H:"):
                raw = text.replace("H:", "").strip()
                Height = convert_fraction_to_decimal(raw)
            elif text.startswith("Dia") or text.startswith("Diameter"):
                raw = re.sub(r'^Dia(?:meter)?:\s*', '', text).strip()
                Diameter = convert_fraction_to_decimal(raw)

        # Weight (outside of dimension sections)
        if text.lower().startswith("weight:"):
            Weight = convert_fraction_to_decimal(text.replace("Weight:", "").strip())

    # -------------------------------
    # Description Section
    # -------------------------------
    description = ""
    desc_header = soup.find("h2",
        class_="elementor-heading-title elementor-size-default",
        string=re.compile("DETAILS", re.I)
    )

    if desc_header:
        next_section = desc_header.find_next("section")
        if next_section:
            desc_header.decompose()
            for unwanted in next_section.find_all(["a", "button", "script", "style"]):
                unwanted.decompose()
            description = next_section.get_text(separator="\n", strip=True)

    # -------------------------------
    # Finish Section (all finishes, title case)
    # -------------------------------
    Finish = ""
    finish_links = soup.find_all("a", class_="jet-listing-dynamic-terms__link")
    if finish_links:
        finish_names = []
        for link in finish_links:
            name = link.get_text(strip=True)
            if name:
                # Convert "31 FRANZ MAYER" → "31 Franz Mayer"
                parts = name.split(" ", 1)
                if len(parts) == 2 and parts[0].isdigit():
                    name = parts[0] + " " + parts[1].title()
                else:
                    name = name.title()
                finish_names.append(name)
        Finish = ", ".join(finish_names)

    # -------------------------------
    # Append row
    # -------------------------------
    all_rows.append({
        "Product URL": product_url,
        "Image URL": image_url,
        "Product Name": product_name,
        "SKU": sku,
        "Product Family Id": product_family_id,
        "Description": description,
        "Weight": Weight,
        "Width": Width,
        "Depth": Depth,
        "Diameter": Diameter,
        "Height": Height,
        "Finish": Finish
    })

    print(f"[{idx + 1}/{total}] Scraped: {product_name} | SKU: {sku} | W: {Width} | D: {Depth} | Dia: {Diameter} | H: {Height} | Finish: {Finish}")

# Write all rows at once (enforce column order)
pd.DataFrame(all_rows, columns=columns).to_excel(final_path, index=False)
driver.quit()
print(f"\n✅ Scraping complete! Final Excel saved at:\n{final_path}")