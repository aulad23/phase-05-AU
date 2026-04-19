import requests
import pandas as pd
import re
from bs4 import BeautifulSoup

INPUT_FILE = "studiobelvetro_collections.xlsx"
OUTPUT_FILE = "studiobelvetro_products_FINAL.xlsx"

HEADERS = {"User-Agent": "Mozilla/5.0"}

# ---------------- HELPERS ----------------

def extract(pattern, text):
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def clean(v):
    return v.replace('"', '').replace('”', '').strip()

def parse_illumination(text):
    data = {"Wattage": "", "Color Temperature": "", "Base": ""}

    watt = extract(r"(\d+(\.\d+)?)\s*w", text)
    if watt:
        data["Wattage"] = watt

    temp = extract(r"(\d{3,5})\s*K", text)
    if temp:
        data["Color Temperature"] = temp

    bases = []
    for b in ["LED", "GU10", "MR16"]:
        if re.search(rf"\b{b}\b", text, re.IGNORECASE):
            bases.append(b)

    data["Base"] = ", ".join(bases)
    return data

def parse_details(details):
    data = {
        "Diameter": clean(extract(r"DIA:\s*([\d\-\.]+)", details)),
        "Height": clean(extract(r"\bH:\s*([\d\-\.]+)", details)),
        "Width": clean(extract(r"\bW:\s*([\d\-\.]+)", details)),
        "Depth": clean(extract(r"\bD:\s*([\d\-\.]+)", details)),
        "Weight": "",
        "Conpoy": "",
        "Finish": "",
        "Illumination": "",
        "Wattage": "",
        "Color Temperature": "",
        "Base": ""
    }

    weight = extract(r"(WEIGHT|IBS|IB|LBS|LB)[^\d]*([\d\-\.]+)", details)
    if weight:
        data["Weight"] = clean(weight)

    cd = clean(extract(r"CANOPY\s*DIA:\s*([\d\-\.]+)", details))
    ch = clean(extract(r"CANOPY\s*H:\s*([\d\-\.]+)", details))
    data["Conpoy"] = ", ".join([v for v in [cd, ch] if v])

    data["Finish"] = extract(r"METAL FINISH OPTIONS:\s*(.+?)(?:\||$)", details)

    illumination = extract(r"ILLUMINATION:\s*(.+?)(?:\||$)", details)
    data["Illumination"] = illumination

    data.update(parse_illumination(illumination))
    return data

# ---------------- MAIN ----------------

df = pd.read_excel(INPUT_FILE)
rows = []

for _, r in df.iterrows():
    product_url = r["Product URL"]
    image_url = r["Image URL"]
    product_name = r["Product Name"]

    res = requests.get(product_url, headers=HEADERS, timeout=30)
    soup = BeautifulSoup(res.text, "html.parser")

    description = ""
    d = soup.find("div", class_="product-description")
    if d and d.find("p"):
        description = d.find("p").get_text(strip=True)

    sku = ""
    s = soup.find("span", class_="model-number")
    if s:
        sku = s.get_text(strip=True)

    detail_parts = []
    specs = soup.find("div", class_="col c8")
    if specs:
        for li in specs.find_all("li"):
            detail_parts.append(li.get_text(" ", strip=True))

    details = " | ".join(detail_parts)
    parsed = parse_details(details)

    rows.append([
        product_url,
        image_url,
        product_name,
        sku,
        product_name,
        description,
        parsed["Weight"],
        parsed["Width"],
        parsed["Depth"],
        parsed["Diameter"],
        parsed["Height"],
        parsed["Conpoy"],
        parsed["Finish"],
        parsed["Wattage"],
        parsed["Color Temperature"],
        parsed["Base"],
        details,
        parsed["Illumination"]
    ])

# ---------------- SAVE ----------------

final_df = pd.DataFrame(rows, columns=[
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
    "Conpoy",
    "Finish",
    "Wattage",
    "Color Temperature",
    "Base",
    "Details",
    "Illumination"
])

final_df.to_excel(OUTPUT_FILE, index=False)

print("✅ FINAL EXCEL CREATED WITH REQUESTED ORDER:", OUTPUT_FILE)
