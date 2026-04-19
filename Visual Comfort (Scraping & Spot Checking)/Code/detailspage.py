import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import sys
import os

# ============== Config ==============
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_XLSX = os.path.join(SCRIPT_DIR, "visualcomfort_bulbs.xlsx")
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, "visualcomfort_bulbs_full_data.xlsx")

HEADLESS = False
BATCH_SIZE = 5   # Save progress after every 5 newly processed products

# ============== Driver Setup ==============
options = webdriver.ChromeOptions()
if HEADLESS:
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)


def wait_css(css, timeout=25):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css))
    )


def get_soup(url, wait_time=1.0):
    """
    Load a URL in the browser and return BeautifulSoup of the page.
    """
    print(f"Loading: {url}")
    driver.get(url)
    try:
        wait_css("#product-specifications-bottom, #product-specifications, div.product-container, main, body")
    except Exception:
        pass
    time.sleep(wait_time)
    return BeautifulSoup(driver.page_source, "html.parser")


def clean_text(s: str) -> str:
    """
    Normalize whitespace in text.
    """
    return " ".join((s or "").split())


# ============== Helper: get value from configurable select (by label text, BeautifulSoup) ==============
def get_select_value_by_label(soup: BeautifulSoup, label_text: str) -> str:
    """
    <div class="field configurable required">
        <label class="label"><span>Finish / Color Temperature</span></label>
        <div class="control">
            <select class="super-attribute-select">
                <option value="">Choose an Option...</option>
                <option value="...">2700K</option>
                ...
            </select>
        </div>
    </div>

    → এই block থেকে selected / first non-"Choose an Option..." option এর text return করবে
    """
    label_text = label_text.strip().lower()
    # ✅ required ছাড়াও configurable field ধরব
    fields = soup.select("div.field.configurable")
    for fld in fields:
        lab = fld.select_one("label span")
        if not lab:
            continue
        lab_txt = lab.get_text(strip=True).lower()
        # ✅ exact না হোক, শুধু ভিতরে থাকলেই হবে
        if label_text not in lab_txt:
            continue

        sel = fld.select_one("select.super-attribute-select")
        if not sel:
            continue

        # 1) selected option খুঁজি
        for opt in sel.find_all("option"):
            sel_attr = opt.get("selected")
            if sel_attr is not None or opt.has_attr("selected"):
                txt = clean_text(opt.get_text())
                if txt and txt.lower() != "choose an option...":
                    return txt

        # 2) fallback: first non-"Choose an Option..." option
        for opt in sel.find_all("option"):
            txt = clean_text(opt.get_text())
            if txt and txt.lower() != "choose an option...":
                return txt

    return ""


# ============== Helper: get Color Temperature using Selenium after variation click ==============
def get_color_temp_selenium(timeout: int = 5) -> str:
    """
    Variation click করার পর live DOM থেকে 'Color Temperature'
    field-এর select থেকে current selected / first non-default option নেয়া হবে.
    Extra wait যোগ করা হয়েছে যাতে JS late load হলেও dhore.
    """
    try:
        # ✅ wait until Color Temperature field exists (if it will exist)
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//div[contains(@class,'field') and contains(@class,'configurable')]"
                    "[.//label/span[contains(translate(normalize-space(.),"
                    " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                    " 'color temperature')]]"
                    "//select[contains(@class,'super-attribute-select')]"
                )
            )
        )
    except Exception:
        # jodi thakbena eirokom page hoy, tahole blank return korte dibo
        pass

    try:
        # ✅ required ছাড়াও configurable field ধরব
        fields = driver.find_elements(By.CSS_SELECTOR, "div.field.configurable")
    except Exception:
        return ""

    for fld in fields:
        try:
            label_span = fld.find_element(By.CSS_SELECTOR, "label span")
        except Exception:
            continue

        lab_txt = label_span.text.strip().lower()
        if "color temperature" not in lab_txt:
            continue

        try:
            sel = fld.find_element(By.CSS_SELECTOR, "select.super-attribute-select")
        except Exception:
            continue

        options = sel.find_elements(By.TAG_NAME, "option")
        selected = ""

        # 1) selected option
        for opt in options:
            if opt.get_attribute("selected") or opt.is_selected():
                t = clean_text(opt.text)
                if t and t.lower() != "choose an option...":
                    selected = t
                    break

        # 2) fallback: first non-"Choose an Option..." option
        if not selected:
            for opt in options:
                t = clean_text(opt.text)
                if t and t.lower() != "choose an option...":
                    selected = t
                    break

        return selected

    return ""


def scrape_main_image() -> str:
    """
    Scrape the main product image URL from the page
    (before clicking on any variation).
    Prefer the active fotorama image and strip any query string.
    """
    try:
        img_el = None
        css_candidates = [
            ".fotorama__stage__frame.fotorama__active img",
            ".fotorama__stage__frame.fotorama__active .fotorama__img",
            "div.product.media img",
            ".gallery-placeholder__image",
            "img.product-image-photo",
        ]
        for css in css_candidates:
            try:
                img_el = driver.find_element(By.CSS_SELECTOR, css)
                if img_el:
                    break
            except Exception:
                continue

        if img_el:
            src = img_el.get_attribute("src") or ""
            src = src.strip()
            # Strip any query string like ?$product_page_image_medium$
            if "?" in src:
                src = src.split("?", 1)[0]
            return src
    except Exception:
        pass
    return ""


# ============== Specifications (old + new layouts) ==============
def scrape_specs_block(soup: BeautifulSoup):
    """
    Specifications can appear in two layouts:
    1) Old layout: #product-specifications-bottom (label + pure-value)
    2) New layout: #spec-inch-tab table.product-attribute-specs-table
    Returns:
      specs_text, extension, backplate, base
    """
    specs_lines = []
    extension_value = ""
    backplate_value = ""
    base_value = ""

    # ---------- Case 1: Old layout (#product-specifications-bottom) ----------
    rows = soup.select("#product-specifications-bottom table.options tbody tr")
    if not rows:
        rows = soup.select("#product-specifications-bottom tbody tr")

    for row in rows:
        label_el = row.select_one(".label")
        value_el = row.select_one(".pure-value")

        if label_el and value_el:
            label = clean_text(label_el.get_text(strip=True))
            value = clean_text(value_el.get_text(strip=True))
            if label and value:
                specs_lines.append(f"{label}: {value}")

                if label.lower() == "extension":
                    extension_value = value
                elif label.lower() == "backplate":
                    backplate_value = value
                elif label.lower() == "base":
                    base_value = value

                continue

        cells = row.find_all(["td", "th"])
        if cells:
            text = clean_text(" ".join(c.get_text(" ", strip=True) for c in cells))
            if text and text.lower() not in {"", ":", "—", "-"}:
                specs_lines.append(text)

    # ---------- Case 2: New layout (#spec-inch-tab) ----------
    if not specs_lines:
        inch_cells = soup.select(
            "#spec-inch-tab table.product-attribute-specs-table tbody tr td"
        )
        for cell in inch_cells:
            text = clean_text(cell.get_text(" ", strip=True))
            if not text:
                continue

            specs_lines.append(text)

            lower = text.lower()
            if lower.startswith("extension:"):
                extension_value = clean_text(text.split(":", 1)[1])
            elif lower.startswith("backplate:"):
                backplate_value = clean_text(text.split(":", 1)[1])
            elif lower.startswith("base:"):
                base_value = clean_text(text.split(":", 1)[1])

    # Final fallback
    if not specs_lines:
        alt_items = soup.select("#product-specifications-bottom li, #product-specifications-bottom p")
        for it in alt_items:
            text = clean_text(it.get_text(" ", strip=True))
            if text:
                specs_lines.append(text)

    return "\n".join(specs_lines), extension_value, backplate_value, base_value


def scrape_description(soup: BeautifulSoup) -> str:
    """
    Scrape the product description paragraph.
    """
    desc_div = soup.select_one("div.additional-description div.content")
    return desc_div.get_text(strip=True) if desc_div else ""


def scrape_rating(soup: BeautifulSoup) -> str:
    """
    Scrape rating information if present.
    """
    rating_wrapper = soup.select_one("div.rating-wrapper")
    if rating_wrapper:
        items = rating_wrapper.select("div.rating-flex div.rating-item p")
        ratings = [clean_text(p.get_text()) for p in items if p.get_text(strip=True)]
        return ", ".join(ratings)
    return ""


def scrape_shade_details(soup: BeautifulSoup) -> str:
    """
    Scrape the "Shade Details" row if present.
    """
    rows = soup.select("#product-specifications-bottom table.options tbody tr")
    if not rows:
        rows = soup.select("#product-specifications-bottom tbody tr")

    for row in rows:
        label_el = row.select_one("td.label")
        value_el = row.select_one("td.pure-value")

        if label_el and value_el:
            label = clean_text(label_el.get_text(strip=True))
            if label.lower() == "shade details":
                return clean_text(value_el.get_text(strip=True))

    alt = soup.find("td", string=lambda t: t and "shade details" in t.lower())
    if alt:
        next_td = alt.find_next("td")
        if next_td:
            return clean_text(next_td.get_text(strip=True))

    return ""


def scrape_finish_details(soup: BeautifulSoup) -> str:
    """
    Product-level finish options list (all options together) – optional info.
    Per-variation finish is handled separately in scrape_variations().
    """
    finish_div = soup.select_one(
        "div.field.configurable.required select.super-attribute-select"
    )
    if finish_div:
        options = [
            opt.get_text(strip=True)
            for opt in finish_div.find_all("option")
            if opt.get_text(strip=True) not in {"Choose an Option...", ""}
        ]
        if options:
            return ", ".join(options)
    return ""


# ============== Variation Navigation Helper ==============
def click_variation_thumbnail_with_carousel(sku: str, max_attempts: int = 20) -> bool:
    """
    Try to click the variation thumbnail for a given SKU.
    At most ~4 thumbnails are visible at a time. If the thumbnail
    is not in the current visible set, click the owl-next button to slide
    the carousel and then try again, up to max_attempts times.
    """
    attempts = 0

    while attempts < max_attempts:
        try:
            thumbs = driver.find_elements(
                By.CSS_SELECTOR,
                ".product-item-variation-carousel-wrapper a.configurable-thumbnail"
            )

            target_el = None
            for el in thumbs:
                sku_attr = (el.get_attribute("data-product-sku") or
                            el.get_attribute("data-clp-sku") or "").strip()
                if sku_attr == sku:
                    target_el = el
                    break

            if target_el:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_el)
                driver.execute_script("arguments[0].click();", target_el)

                try:
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "span.price"))
                    )
                except Exception:
                    pass
                time.sleep(0.4)
                return True

        except Exception:
            pass

        try:
            owl_next = driver.find_element(By.CSS_SELECTOR, "button.owl-next")
        except Exception:
            owl_next = None

        if not owl_next:
            break

        classes = owl_next.get_attribute("class") or ""
        if "disabled" in classes:
            break

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", owl_next)
            owl_next.click()
            time.sleep(0.4)
        except Exception:
            break

        attempts += 1

    print(f"⚠️ Could not click variation thumbnail for SKU '{sku}' after {attempts} attempts")
    return False


# ============== Variations + Product Family Id ==============
def scrape_variations():
    """
    variations_list: list of dicts for each variation:
        sku, name, price, qty, status, image, finish, url, color_temp
    """
    soup_initial = BeautifulSoup(driver.page_source, "html.parser")
    thumb_tags = soup_initial.select(
        ".product-item-variation-carousel-wrapper a.configurable-thumbnail"
    )

    if not thumb_tags:
        print("⚠️ No variation thumbnails found on this page")
        return [], []

    variation_meta = []

    for tag in thumb_tags:
        sku = (tag.get("data-product-sku") or tag.get("data-clp-sku") or "").strip()
        if not sku:
            continue

        name = (tag.get("data-product-name") or tag.get("data-mainproduct-name") or "").strip()
        alt_name = (tag.get("data-clp-image-alt") or "").strip()
        if not alt_name:
            alt_name = name

        qty = (tag.get("data-qty-value") or "").strip()
        status = (tag.get("data-filter-message") or "").strip()
        image_attr = (tag.get("data-clp-image") or tag.get("data-pdp-medium-image") or "").strip()
        price_data = (tag.get("data-clp-price") or tag.get("data-full-price") or tag.get("data-product-price") or "").strip()
        finish_attr = (tag.get("data-clp-finish") or "").strip()

        variation_meta.append({
            "sku": sku,
            "name": name,
            "alt_name": alt_name,
            "qty": qty,
            "status": status,
            "image_attr": image_attr,
            "price_data": price_data,
            "finish_attr": finish_attr,
        })

    print(f"🔢 Found {len(variation_meta)} variations")

    variations_list = []

    for v in variation_meta:
        sku = v["sku"]

        # click variation
        click_variation_thumbnail_with_carousel(sku)

        # choto wait, then fresh DOM
        time.sleep(0.3)
        current_url = driver.current_url
        soup_var = BeautifulSoup(driver.page_source, "html.parser")

        # === 1) Price ===
        price = v["price_data"]
        try:
            price_el = driver.find_element(By.CSS_SELECTOR, "span.price")
            price_text = price_el.text.strip()
            if price_text:
                price = price_text
        except Exception:
            pass

        # === 2) Finish ===
        finish_text = v["finish_attr"]
        try:
            sel = driver.find_element(By.CSS_SELECTOR, "select.super-attribute-select")
            options = sel.find_elements(By.TAG_NAME, "option")
            selected = ""

            for opt in options:
                if opt.get_attribute("selected") or opt.is_selected():
                    t = clean_text(opt.text)
                    if t and t.lower() != "choose an option...":
                        selected = t
                        break

            if not selected and options:
                first_text = clean_text(options[0].text)
                if first_text and first_text.lower() != "choose an option...":
                    selected = first_text

            if selected:
                finish_text = selected
        except Exception:
            pass

        if not finish_text:
            nm = v["alt_name"] or v["name"]
            if nm and " in " in nm:
                finish_text = nm.split(" in ", 1)[1].strip()

        # === 3) Color Temperature (Selenium-based with wait) ===
        color_temp = get_color_temp_selenium()
        if not color_temp:
            color_temp = get_select_value_by_label(soup_var, "color temperature")

        # === 4) Image ===
        image = v["image_attr"]
        try:
            img_el = None
            css_candidates = [
                ".fotorama__stage__frame.fotorama__active img",
                ".fotorama__stage__frame.fotorama__active .fotorama__img",
                "div.product.media img",
                ".gallery-placeholder__image",
                "img.product-image-photo",
            ]
            for css in css_candidates:
                try:
                    img_el = driver.find_element(By.CSS_SELECTOR, css)
                    if img_el:
                        break
                except Exception:
                    continue

            if img_el:
                src = img_el.get_attribute("src") or ""
                src = src.strip()
                if "?" in src:
                    src = src.split("?", 1)[0]
                image = src
        except Exception:
            pass

        alt_name = v["alt_name"] or v["name"]

        variations_list.append({
            "sku": sku,
            "name": alt_name,
            "price": price,
            "qty": v["qty"],
            "status": v["status"],
            "image": image,
            "finish": finish_text,
            "color_temp": color_temp,
            "url": current_url,
        })

    return [], variations_list


def scrape_details(url: str):
    """
    Scrape all common fields + all variations for a single product URL.
    Returns:
      common_fields (dict),
      family_id (str),
      variations_list (list[dict])
    """
    try:
        soup = get_soup(url)
    except Exception as e:
        print(f"❌ Failed to load: {url} -> {e}", file=sys.stderr)
        return {}, "", []

    # ---- Product Name + Product Family Id main h1 theke ----
    product_name = ""
    title_el = soup.select_one("h1.page-title span, h1.page-title")
    if title_el:
        product_name = clean_text(title_el.get_text(" ", strip=True))

    family_main = ""
    if product_name:
        if " in " in product_name:
            family_main = product_name.split(" in ", 1)[0].strip()
        else:
            family_main = product_name

    specifications, extension, backplate, base = scrape_specs_block(soup)
    description = scrape_description(soup)
    rating = scrape_rating(soup)
    shade_details = scrape_shade_details(soup)
    finish_details = scrape_finish_details(soup)

    # Base Color Temperature from initial DOM (if dropdown exists)
    color_temp_base = get_select_value_by_label(soup, "color temperature")

    # Main product image from initial page load (before variation clicks)
    main_image_url = scrape_main_image()

    # Variations
    _, variations_list = scrape_variations()

    product_family_id = family_main

    common_fields = {
        "Product Name": product_name,
        "Specifications": specifications,
        "Description": description,
        "Rating": rating,
        "Shade Details": shade_details,
        "Finish": finish_details,   # default, will be overridden per variation
        "Extension": extension,
        "Backplate": backplate,
        "Base": base,
        "Product Family Id": product_family_id,
        "Image URL": main_image_url,   # main product image for base row
        "Color Temperature": color_temp_base,  # base-level color temp (if any)
    }

    return common_fields, product_family_id, variations_list


# ============== Main Logic (with RESUME) ==============
def main():
    df_input = pd.read_excel(INPUT_XLSX)
    if "Product URL" not in df_input.columns:
        print("❌ 'Product URL' column not found in input file.", file=sys.stderr)
        driver.quit()
        return

    total_products = len(df_input)

    # ---- Resume support ----
    output_rows = []
    processed_indices = set()

    if os.path.exists(OUTPUT_XLSX):
        print(f"📂 Found existing output file: {OUTPUT_XLSX} — resuming...")
        df_existing = pd.read_excel(OUTPUT_XLSX)

        output_rows = df_existing.to_dict("records")

        if "Input Index" in df_existing.columns:
            try:
                processed_indices = set(
                    int(x) for x in df_existing["Input Index"].dropna().tolist()
                )
            except Exception:
                processed_indices = set()
        else:
            processed_indices = set()

        if processed_indices:
            print(f"✅ Already processed {len(processed_indices)} products (by Input Index).")
    else:
        print("🆕 No previous output file found — starting fresh...")

    print(f"🔢 Total products in input: {total_products}")

    for i in range(total_products):
        if i in processed_indices:
            print(f"⏭️ Skipping product {i + 1}/{total_products} (already processed)")
            continue

        url = df_input.at[i, "Product URL"]
        print(f"\n🔎 Scraping product {i + 1}/{total_products} -> {url}")

        common_fields, family_id, variations_list = scrape_details(url)
        base_row = df_input.iloc[i].to_dict()

        # Ensure List Price key exists (used as fallback for variations only)
        if "List Price" not in base_row:
            base_row["List Price"] = base_row.get("List Price", "")

        # Default price (for variations fallback): start from input List Price
        default_price = base_row.get("List Price", "")

        # If still empty, try from first variation
        if not default_price and variations_list:
            default_price = variations_list[0].get("price", "") or ""

        # If still empty, try reading from main page price
        if not default_price:
            try:
                main_price_el = driver.find_element(By.CSS_SELECTOR, "span.price")
                default_price = main_price_el.text.strip()
            except Exception:
                default_price = base_row.get("List Price", "")

        # 👉 1) BASE PRODUCT ROW
        base_product_row = base_row.copy()
        base_product_row.update(common_fields)

        base_product_row["Input Index"] = i

        if "Product URL" in base_product_row:
            base_product_row["Product URL"] = url

        output_rows.append(base_product_row)

        # 👉 2) VARIATION ROWS
        if variations_list:
            for var in variations_list:
                row = base_row.copy()
                row.update(common_fields)

                var_sku = var.get("sku", "")
                var_price = var.get("price", "")
                var_finish = var.get("finish", "")
                var_url = var.get("url", "")
                var_color_temp = var.get("color_temp", "")

                row["Input Index"] = i

                sku_cols = ["SKU", "Sku", "sku", "Product SKU", "Product Sku", "product_sku"]
                found_sku_col = False
                for col in sku_cols:
                    if col in row:
                        row[col] = var_sku
                        found_sku_col = True
                if not found_sku_col:
                    row["SKU"] = var_sku

                row["List Price"] = var_price if var_price else default_price

                if var_finish:
                    row["Finish"] = var_finish

                if var_color_temp:
                    row["Color Temperature"] = var_color_temp

                if "Product URL" in row and var_url:
                    row["Product URL"] = var_url

                image_url_col = None
                for col in ["Image URL", "Image", "image_url", "image"]:
                    if col in row:
                        image_url_col = col
                        break

                if image_url_col is None:
                    image_url_col = "Image URL"

                row[image_url_col] = var.get("image", "")

                output_rows.append(row)

        processed_indices.add(i)

        if len(processed_indices) % BATCH_SIZE == 0 or i == total_products - 1:
            df_out = pd.DataFrame(output_rows)

            if "Image URL" in df_out.columns and "Image" in df_out.columns:
                df_out.drop(columns=["Image"], inplace=True)

            df_out.to_excel(OUTPUT_XLSX, index=False)
            print(f"💾 Progress saved after product {i + 1}/{total_products} -> {OUTPUT_XLSX}")

    print(f"\n✅ Done! All data (main product row + variation rows, with Color Temperature + resume) saved to '{OUTPUT_XLSX}'")


if __name__ == "__main__":
    try:
        main()
    finally:
        driver.quit()
