from playwright.sync_api import sync_playwright
import pandas as pd
import time
import re
from urllib.parse import urljoin, urlparse

# ================ CONFIG ===============
INPUT_FILE = "uttermost_Rugs.xlsx"
OUTPUT_FILE = "uttermost_Rugs_details.xlsx"
HEADLESS = False
BASE_URL = "https://uttermost.com"
BATCH_SIZE = 10  # Save after every 10 products


# ========================================

def normalize_image_url(image_url, base_url=BASE_URL):
    """Normalize image URL by adding base URL if it's a relative path."""
    if not image_url:
        return ""
    if image_url.startswith(('http://', 'https://')):
        return image_url
    return urljoin(base_url, image_url)


def extract_product_family_id(product_name):
    """Extract the common part of product name (without variation)"""
    cleaned = re.sub(r'\b(Side|Coffee|End|Accent)\s+(Table|Desk)?\b', '', product_name, flags=re.IGNORECASE)
    cleaned = re.sub(r'\d+[WwHhDd]\s*[xX×]\s*\d+[WwHhDd].*', '', cleaned)
    cleaned = re.sub(r'\(.*?\)', '', cleaned)
    return cleaned.strip()


def parse_dimensions(dimension_str):
    """Parse dimension string and extract Width, Height, Depth, Length, Diameter"""
    result = {"Width": "", "Height": "", "Depth": "", "Length": "", "Diameter": ""}

    if not dimension_str or dimension_str == "N/A":
        return result

    # Check for Diameter first (now extracts just the number)
    diameter_pattern = r'(\d+\.?\d*)\s*(?:Dia\.?|Diam\.?|DIA|Diameter)'
    diameter_match = re.search(diameter_pattern, dimension_str, re.IGNORECASE)
    if diameter_match:
        result["Diameter"] = diameter_match.group(1)
        # Still extract height if present with diameter
        height_pattern = r'(\d+\.?\d*)\s*[Hh](?:\s|$|\(|[xX×])'
        height_match = re.search(height_pattern, dimension_str)
        if height_match:
            result["Height"] = height_match.group(1)
        return result

    # Parse other dimensions
    depth_pattern = r'(\d+\.?\d*)\s*[Dd](?:\s|$|\()'
    depth_match = re.search(depth_pattern, dimension_str)
    if depth_match:
        result["Depth"] = depth_match.group(1)

    width_pattern = r'(\d+\.?\d*)\s*[Ww](?:\s|$|\()'
    width_match = re.search(width_pattern, dimension_str)
    if width_match:
        result["Width"] = width_match.group(1)

    height_pattern = r'(\d+\.?\d*)\s*[Hh](?:\s|$|\()'
    height_match = re.search(height_pattern, dimension_str)
    if height_match:
        result["Height"] = height_match.group(1)

    length_pattern = r'(\d+\.?\d*)\s*[Ll](?:\s|$|\()'
    length_match = re.search(length_pattern, dimension_str)
    if length_match:
        result["Length"] = length_match.group(1)

    return result


def parse_weight(weight_str):
    """Extract only the numeric value from weight"""
    if not weight_str or weight_str == "N/A":
        return ""
    number_match = re.search(r'(\d+\.?\d*)', weight_str)
    if number_match:
        return number_match.group(1)
    return ""


def parse_details(details_str):
    """Parse details string and extract all attributes into separate columns"""
    result = {}
    if not details_str or details_str == "N/A":
        return result

    parts = details_str.split("|")
    for part in parts:
        part = part.strip()
        if ":" in part:
            key, value = part.split(":", 1)
            key = key.strip()
            value = value.strip()

            if "Seat Size" in key:
                size_match = re.search(r'(\d+\.?\d*)\s*[*xX×]\s*(\d+\.?\d*)', value)
                if size_match:
                    result["Seat Width"] = size_match.group(1)
                    result["Seat Depth"] = size_match.group(2)
            elif "Seat Height" in key:
                number_match = re.search(r'(\d+\.?\d*)', value)
                if number_match:
                    result["Seat Height"] = number_match.group(1)
            elif "Arm Height" in key:
                number_match = re.search(r'(\d+\.?\d*)', value)
                if number_match:
                    result["Arm Height"] = number_match.group(1)
            else:
                result[key] = value

    return result


def safe_click_variation(page, option_type, value, max_retries=3):
    """
    Safely click a variation option with better error handling.
    Returns True if successful, False otherwise.
    """
    for attempt in range(max_retries):
        try:
            # First, scroll to the options section to ensure it's visible
            options_section = page.locator(f"div.option-root-JXw:has(span.option-title-KCu:has-text('{option_type}'))")

            # Check if this option type exists on the page
            if options_section.count() == 0:
                print(f"     ℹ {option_type} option not available on this page")
                return True  # Not an error, just not applicable

            if options_section.count() > 0:
                try:
                    options_section.first.scroll_into_view_if_needed(timeout=5000)
                    time.sleep(0.5)
                except:
                    # If scroll fails, try without it
                    pass

            # Try finding button by title first
            button = page.locator(
                f"div.option-root-JXw:has(span.option-title-KCu:has-text('{option_type}')) "
                f"button[title='{value}']"
            )

            # If not found by title, try by span text
            if button.count() == 0:
                button = page.locator(
                    f"div.option-root-JXw:has(span.option-title-KCu:has-text('{option_type}')) "
                    f"button:has(span:has-text('{value}'))"
                )

            if button.count() > 0:
                # Check if already selected
                first_button = button.first
                class_attr = first_button.get_attribute("class") or ""

                if "selected" in class_attr.lower():
                    print(f"     → {option_type} '{value}' already selected")
                    return True

                # Scroll button into view with shorter timeout
                try:
                    first_button.scroll_into_view_if_needed(timeout=5000)
                    time.sleep(0.3)
                except:
                    # If scroll fails, continue anyway
                    pass

                # Try clicking with force if needed
                try:
                    first_button.click(timeout=5000)
                except:
                    # If regular click fails, try force click
                    try:
                        first_button.click(force=True, timeout=3000)
                    except:
                        print(f"     ⚠ Could not click {option_type}: {value}")
                        if attempt < max_retries - 1:
                            continue
                        return False

                # Wait for page to update
                time.sleep(2)

                # Verify the click worked
                try:
                    page.wait_for_load_state("networkidle", timeout=5000)
                except:
                    # Timeout is okay, page might already be loaded
                    pass

                print(f"     → Clicked {option_type}: {value}")
                return True
            else:
                if attempt == 0:
                    print(f"     ℹ {option_type} '{value}' not found (might not be available for this variation)")
                return False

        except Exception as e:
            error_msg = str(e)[:100]
            if "Timeout" not in error_msg or attempt == max_retries - 1:
                print(f"     ⚠ Attempt {attempt + 1} failed for {option_type} '{value}': {error_msg}")

            if attempt < max_retries - 1:
                time.sleep(1)
                # Only reload on last retry
                if attempt == max_retries - 2:
                    try:
                        page.reload(wait_until="domcontentloaded", timeout=10000)
                        time.sleep(2)
                    except:
                        pass
            else:
                return False

    return False


def scrape_product_details(page, url):
    """Scrape details from a single product page"""
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)

        try:
            page.wait_for_selector("section.productFullDetail-title-lo9", state="visible", timeout=15000)
            time.sleep(1.5)
        except:
            print(f"  ⚠ Product details not loading for {url}")
            return None

        # Extract Description
        description = ""
        try:
            desc_locator = page.locator("span.productFullDetail-shortDescription--SS div.richContent-root-Ddk")
            if desc_locator.count() > 0:
                description = desc_locator.first.text_content().strip()
        except:
            pass

        # Extract Dimensions
        dimension = ""
        try:
            specs_section = page.locator("section.productFullDetail-details-Glq")
            if specs_section.count() > 0:
                dim_items = specs_section.locator("li").all()
                for item in dim_items:
                    text = item.text_content()
                    if "Dimensions" in text or "Dimension" in text:
                        parts = text.split(":")
                        if len(parts) > 1:
                            dimension = parts[1].strip()
                            break
        except:
            pass

        # Extract Weight
        weight = ""
        try:
            specs_section = page.locator("section.productFullDetail-details-Glq")
            if specs_section.count() > 0:
                weight_items = specs_section.locator("li").all()
                for item in weight_items:
                    text = item.text_content()
                    if "Weight" in text:
                        parts = text.split(":")
                        if len(parts) > 1:
                            weight = parts[1].strip()
                            break
        except:
            pass

        # Extract More Details
        details = ""
        try:
            view_more_btn = page.locator("button:has-text('View More'), button:has-text('view more')")
            if view_more_btn.count() > 0:
                view_more_btn.first.click()
                time.sleep(1)

            details_list = []
            custom_attrs = page.locator("div.customAttributes-root-MXb ul.customAttributes-list-qDg li")
            if custom_attrs.count() > 0:
                for attr in custom_attrs.all():
                    try:
                        label_el = attr.locator("div.text-label-daH, p.font-bold").first
                        label = label_el.text_content().strip() if label_el.count() > 0 else ""

                        value_el = attr.locator("div.text-content-Mcy, p.ml-2").first
                        value = value_el.text_content().strip() if value_el.count() > 0 else ""

                        if label and value:
                            label = label.replace(":", "").strip()
                            details_list.append(f"{label}: {value}")
                    except:
                        continue

            if details_list:
                details = " | ".join(details_list)
        except Exception as e:
            pass

        # Check for variations
        variations = []
        try:
            # Check for Color variations
            color_section = page.locator("div.option-root-JXw:has(span.option-title-KCu:has-text('Color'))")
            colors = []
            if color_section.count() > 0:
                color_buttons = color_section.locator("button.tile-root-8ZR").all()
                for btn in color_buttons:
                    color_name = btn.get_attribute("title") or ""
                    if color_name:
                        colors.append(color_name.strip())

            # Check for Size variations
            size_section = page.locator("div.option-root-JXw:has(span.option-title-KCu:has-text('Size'))")
            sizes = []
            if size_section.count() > 0:
                size_buttons = size_section.locator("button.tile-root-8ZR").all()
                for btn in size_buttons:
                    size_text = ""
                    span_el = btn.locator("span")
                    if span_el.count() > 0:
                        size_text = span_el.text_content().strip()
                    if not size_text:
                        size_text = btn.get_attribute("title") or ""
                    if size_text:
                        sizes.append(size_text.strip())

            # Create variations based on what's available
            if colors and sizes:
                # Both color and size exist - create all combinations
                for color in colors:
                    for size in sizes:
                        variations.append({"color": color, "size": size})
            elif colors:
                # Only colors exist - create color variations WITHOUT size
                for color in colors:
                    variations.append({"color": color, "size": None})
            elif sizes:
                # Only sizes exist - create size variations WITHOUT color
                for size in sizes:
                    variations.append({"color": None, "size": size})
        except Exception as e:
            print(f"  ⚠ Error extracting variations: {e}")

        return {
            "description": description,
            "dimension": dimension,
            "weight": weight,
            "details": details,
            "variations": variations if variations else None
        }
    except Exception as e:
        print(f"  ✗ Error scraping {url}: {e}")
        return None


# Read input Excel
try:
    df_input = pd.read_excel(INPUT_FILE)
    print(f"✓ Loaded {len(df_input)} products from {INPUT_FILE}")
except Exception as e:
    print(f"✗ Error reading input file: {e}")
    exit(1)

# Prepare output data
all_output = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=HEADLESS)
    context = browser.new_context(
        viewport={"width": 1400, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    )
    page = context.new_page()
    page.route("**/*", lambda route: route.abort() if route.request.resource_type in ["stylesheet", "font",
                                                                                      "media"] else route.continue_())

    total_products = len(df_input)
    products_processed = 0

    for index, row in df_input.iterrows():
        product_url = row['Product URL']
        image_url = row['Image URL']
        product_name = row['Product Name']
        sku = row['SKU']

        print(f"\n[{index + 1}/{total_products}] Processing: {product_name}")
        print(f"  URL: {product_url}")

        product_family_id = extract_product_family_id(product_name)
        details = scrape_product_details(page, product_url)

        if details:
            description = details['description']
            dimension = details['dimension']
            weight = details['weight']
            more_details = details['details']
            variations = details['variations']

            parsed_dims = parse_dimensions(dimension)
            parsed_weight = parse_weight(weight)
            parsed_details = parse_details(more_details)

            if variations:
                print(f"  ✓ Found {len(variations)} variations")

                for idx, var in enumerate(variations):
                    try:
                        # Reload page for fresh state
                        page.goto(product_url, wait_until="domcontentloaded", timeout=60000)
                        page.wait_for_selector("section.productFullDetail-title-lo9", state="visible", timeout=15000)
                        page.wait_for_load_state("networkidle", timeout=10000)
                        time.sleep(1.5)

                        print(f"\n  → Processing variation {idx + 1}/{len(variations)}")

                        # Click color first if exists
                        color_clicked = False
                        if var['color']:
                            success = safe_click_variation(page, "Color", var['color'])
                            if not success:
                                print(f"     ⚠ Skipping variation - couldn't select color")
                                continue
                            color_clicked = True

                        # Then click size if exists (only if this variation has a size)
                        size_clicked = False
                        if var['size']:
                            success = safe_click_variation(page, "Size", var['size'])
                            if success:
                                size_clicked = True
                            else:
                                # If size was expected but couldn't click, skip this variation
                                print(f"     ⚠ Skipping variation - couldn't select size")
                                continue

                        # Wait for all updates to complete
                        time.sleep(1.5)
                        page.wait_for_load_state("networkidle", timeout=10000)

                        # Get updated information
                        var_url = page.url

                        var_name_el = page.locator("section.productFullDetail-title-lo9 h1")
                        var_name = var_name_el.text_content().strip() if var_name_el.count() > 0 else product_name

                        var_sku_el = page.locator("p.productFullDetail-productSku-vjY")
                        var_sku = sku
                        if var_sku_el.count() > 0:
                            sku_text = var_sku_el.text_content()
                            var_sku = sku_text.replace("SKU:", "").strip()

                        # Get updated image
                        var_image = image_url
                        try:
                            img_el = page.locator(
                                "img.carousel-image-bYB, img[alt*='product'], img.productFullDetail-image-, "
                                "div.carousel-root button.carousel-imageContainer img"
                            ).first
                            if img_el.count() > 0:
                                raw_img_url = img_el.get_attribute("src") or img_el.get_attribute("data-src") or ""
                                if raw_img_url:
                                    var_image = normalize_image_url(raw_img_url)
                        except:
                            pass

                        # Get updated dimension
                        var_dimension = ""
                        try:
                            specs_section = page.locator("section.productFullDetail-details-Glq")
                            if specs_section.count() > 0:
                                dim_items = specs_section.locator("li").all()
                                for item in dim_items:
                                    text = item.text_content()
                                    if "Dimensions" in text or "Dimension" in text:
                                        parts = text.split(":")
                                        if len(parts) > 1:
                                            var_dimension = parts[1].strip()
                                            break
                        except:
                            pass

                        # Get updated weight
                        var_weight = ""
                        try:
                            specs_section = page.locator("section.productFullDetail-details-Glq")
                            if specs_section.count() > 0:
                                weight_items = specs_section.locator("li").all()
                                for item in weight_items:
                                    text = item.text_content()
                                    if "Weight" in text:
                                        parts = text.split(":")
                                        if len(parts) > 1:
                                            var_weight = parts[1].strip()
                                            break
                        except:
                            pass

                        # Get updated description
                        var_description = description
                        try:
                            desc_locator = page.locator(
                                "span.productFullDetail-shortDescription--SS div.richContent-root-Ddk")
                            if desc_locator.count() > 0:
                                var_description = desc_locator.first.text_content().strip()
                        except:
                            pass

                        # Get updated details
                        var_details = more_details
                        try:
                            view_more_btn = page.locator("button:has-text('View More'), button:has-text('view more')")
                            if view_more_btn.count() > 0:
                                view_more_btn.first.click()
                                time.sleep(1)

                            details_list = []
                            custom_attrs = page.locator("div.customAttributes-root-MXb ul.customAttributes-list-qDg li")
                            if custom_attrs.count() > 0:
                                for attr in custom_attrs.all():
                                    try:
                                        label_el = attr.locator("div.text-label-daH, p.font-bold").first
                                        label = label_el.text_content().strip() if label_el.count() > 0 else ""

                                        value_el = attr.locator("div.text-content-Mcy, p.ml-2").first
                                        value = value_el.text_content().strip() if value_el.count() > 0 else ""

                                        if label and value:
                                            label = label.replace(":", "").strip()
                                            details_list.append(f"{label}: {value}")
                                    except:
                                        continue

                            if details_list:
                                var_details = " | ".join(details_list)
                        except:
                            pass

                        var_parsed_dims = parse_dimensions(var_dimension)
                        var_parsed_weight = parse_weight(var_weight)
                        var_parsed_details = parse_details(var_details)

                        row_data = {
                            "Product URL": var_url,
                            "Image URL": var_image,
                            "Product Name": var_name,
                            "SKU": var_sku,
                            "Product Family Id": product_family_id,
                            "Description": var_description,
                            "Weight": var_parsed_weight,
                            "Width": var_parsed_dims["Width"],
                            "Depth": var_parsed_dims["Depth"],
                            "Diameter": var_parsed_dims["Diameter"],
                            "Length": var_parsed_dims["Length"],
                            "Height": var_parsed_dims["Height"]
                        }
                        row_data.update(var_parsed_details)
                        all_output.append(row_data)

                        variation_label = f"{var['color'] or ''} {var['size'] or ''}".strip()
                        print(f"     ✓ Completed: {variation_label} | SKU: {var_sku}")

                    except Exception as e:
                        print(f"     ✗ Error processing variation: {str(e)[:100]}")
                        continue
            else:
                normalized_image = normalize_image_url(image_url)
                row_data = {
                    "Product URL": product_url,
                    "Image URL": normalized_image,
                    "Product Name": product_name,
                    "SKU": sku,
                    "Product Family Id": product_family_id,
                    "Description": description,
                    "Weight": parsed_weight,
                    "Width": parsed_dims["Width"],
                    "Depth": parsed_dims["Depth"],
                    "Diameter": parsed_dims["Diameter"],
                    "Length": parsed_dims["Length"],
                    "Height": parsed_dims["Height"]
                }
                row_data.update(parsed_details)
                all_output.append(row_data)
                print(f"  ✓ Details scraped (no variations)")
        else:
            normalized_image = normalize_image_url(image_url)
            row_data = {
                "Product URL": product_url,
                "Image URL": normalized_image,
                "Product Name": product_name,
                "SKU": sku,
                "Product Family Id": extract_product_family_id(product_name),
                "Description": "", "Weight": "", "Width": "", "Depth": "",
                "Diameter": "", "Length": "", "Height": ""
            }
            all_output.append(row_data)
            print(f"  ✗ Failed to scrape details")

        # Save after every BATCH_SIZE products
        products_processed += 1
        if products_processed % BATCH_SIZE == 0 and all_output:
            df_temp = pd.DataFrame(all_output)
            df_temp.to_excel(OUTPUT_FILE, index=False)
            print(f"\n💾 Progress saved: {len(all_output)} rows written to {OUTPUT_FILE}")

    browser.close()

# Save to Excel
if all_output:
    df_output = pd.DataFrame(all_output)
    df_output.to_excel(OUTPUT_FILE, index=False)
    print("\n" + "=" * 50)
    print("✅ Scraping complete!")
    print(f"Total rows in output: {len(df_output)}")
    print(f"Excel saved as: {OUTPUT_FILE}")
    print("=" * 50)
else:
    print("\n⚠️ No data collected.")