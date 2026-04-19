"""
Hennepin Made - Step 1 (Product URL Collection)
Requirements: pip install requests openpyxl
Run: python hennepinmade_step1.py
Output: hennepinmade_Chandeliers.xlsx (or per category)
"""

import requests
import time
import re
from openpyxl import Workbook

BASE_URL = "https://hennepinmade.com"
MANUFACTURER = "Hennepin Made"

# ─── ALL CATEGORIES ───────────────────────────────────────────────
CATEGORIES = {
    "Pendants": [
        "pendant",
    ]
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}


def get_products_from_collection(collection_handle):
    products = []
    page = 1
    while True:
        url = f"{BASE_URL}/collections/{collection_handle}/products.json?limit=250&page={page}"
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            if resp.status_code != 200:
                print(f"  ⚠ Status {resp.status_code} for {collection_handle} page {page}")
                break
            data = resp.json()
            batch = data.get("products", [])
            if not batch:
                break
            products.extend(batch)
            print(f"  Page {page}: {len(batch)} products")
            page += 1
            time.sleep(0.5)
        except Exception as e:
            print(f"  ❌ Error: {e}")
            break
    return products


def save_to_excel(rows, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    headers = ["Manufacturer", "Source", "Product Name"]
    ws.append(headers)
    for row in rows:
        ws.append([row["Manufacturer"], row["Source"], row["Product Name"]])
    wb.save(filename)
    print(f"\n✅ Saved {len(rows)} products to {filename}")


def main():
    for category_name, collection_handles in CATEGORIES.items():
        print(f"\n{'='*60}")
        print(f"  CATEGORY: {category_name}")
        print(f"{'='*60}")

        seen_handles = set()
        all_rows = []

        for handle in collection_handles:
            print(f"\n  Collection: {handle}")
            products = get_products_from_collection(handle)

            for p in products:
                p_handle = p.get("handle", "")
                if p_handle in seen_handles:
                    continue
                seen_handles.add(p_handle)

                product_name = p.get("title", "").strip()
                product_url = f"{BASE_URL}/products/{p_handle}"

                all_rows.append({
                    "Manufacturer": MANUFACTURER,
                    "Source": product_url,
                    "Product Name": product_name,
                })
                print(f"    ✓ {product_name}")

        if all_rows:
            filename = f"hennepinmade_{category_name}.xlsx"
            save_to_excel(all_rows, filename)
        else:
            print(f"  ⚠ No products found for {category_name}")

    print(f"\n{'='*60}")
    print("  ALL CATEGORIES DONE!")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()