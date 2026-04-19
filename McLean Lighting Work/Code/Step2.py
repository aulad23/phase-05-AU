import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re

# Read the Excel file from step 1
df = pd.read_excel("mclean_lighting.xlsx")

headers = {
    "User-Agent": "Mozilla/5.0"
}


def get_product_family_id(product_name):
    """Extract Product Family ID from Product Name (before comma or hyphen)"""
    if not product_name:
        return ""

    # Split by comma first, then by hyphen
    if ',' in product_name:
        return product_name.split(',')[0].strip()
    elif '-' in product_name:
        return product_name.split('-')[0].strip()
    else:
        return product_name.strip()


def extract_dimension(description_html):
    """Extract dimension in various formats:
    - W9″ H9″ D8″ (Width, Height, Depth)
    - 21″H 12″W 6″D (number-first)
    - Diam 10″, Dia 12″, Dia. 8″ (Diameter)
    - Weight 5 Ibs, 3 Ib, 10 IBS (Weight)
    """
    if not description_html:
        return ""

    dimensions = []

    # Pattern 1: Letter-first with full words (Width, Height, Depth, Diameter, Weight)
    # W9″, H21″, D8″, Diam 10″, Dia 12″, Weight 5 Ibs
    pattern_letter = r'(?:W|H|D|Diam\.?|Dia\.?|Width|Height|Depth|Diameter|Weight)\s*\d+(?:\.\d+)?[″"]?(?:\s*(?:Ibs?|IBS|lbs?))?'
    matches_letter = re.findall(pattern_letter, description_html, re.IGNORECASE)
    dimensions.extend(matches_letter)

    # Pattern 2: Number-first (21″H, 12″W, 6″D)
    pattern_number = r'\d+(?:\.\d+)?[″"]?\s*(?:W|H|D|Diam\.?|Dia\.?)'
    matches_number = re.findall(pattern_number, description_html, re.IGNORECASE)
    dimensions.extend(matches_number)

    # If we found dimensions, join them with spaces
    if dimensions:
        # Remove duplicates while preserving order
        seen = set()
        unique_dims = []
        for dim in dimensions:
            dim_clean = dim.strip()
            if dim_clean and dim_clean not in seen:
                seen.add(dim_clean)
                unique_dims.append(dim_clean)

        # Return first 5 dimensions (to avoid getting too much)
        return ' '.join(unique_dims[:5])

    return ""


def parse_dimensions(dimension_str):
    """Parse dimension string into separate Width, Height, Depth, Diameter, Weight"""
    result = {
        'Width': '',
        'Height': '',
        'Depth': '',
        'Diameter': '',
        'Weight': ''
    }

    if not dimension_str:
        return result

    # Extract Width
    width_pattern = r'(?:W|Width)[\s:]*(\d+(?:\.\d+)?)[″"]?'
    width_match = re.search(width_pattern, dimension_str, re.IGNORECASE)
    if width_match:
        result['Width'] = width_match.group(1)

    # Number-first width pattern (12″W)
    width_num_first = r'(\d+(?:\.\d+)?)[″"]?\s*W'
    width_nf_match = re.search(width_num_first, dimension_str, re.IGNORECASE)
    if width_nf_match and not result['Width']:
        result['Width'] = width_nf_match.group(1)

    # Extract Height
    height_pattern = r'(?:H|Height)[\s:]*(\d+(?:\.\d+)?)[″"]?'
    height_match = re.search(height_pattern, dimension_str, re.IGNORECASE)
    if height_match:
        result['Height'] = height_match.group(1)

    # Number-first height pattern (21″H)
    height_num_first = r'(\d+(?:\.\d+)?)[″"]?\s*H'
    height_nf_match = re.search(height_num_first, dimension_str, re.IGNORECASE)
    if height_nf_match and not result['Height']:
        result['Height'] = height_nf_match.group(1)

    # Extract Depth
    depth_pattern = r'(?:D|Depth)[\s:]*(\d+(?:\.\d+)?)[″"]?'
    depth_match = re.search(depth_pattern, dimension_str, re.IGNORECASE)
    if depth_match:
        result['Depth'] = depth_match.group(1)

    # Number-first depth pattern (6″D)
    depth_num_first = r'(\d+(?:\.\d+)?)[″"]?\s*D'
    depth_nf_match = re.search(depth_num_first, dimension_str, re.IGNORECASE)
    if depth_nf_match and not result['Depth']:
        result['Depth'] = depth_nf_match.group(1)

    # Extract Diameter
    diam_pattern = r'(?:Diam\.?|Dia\.?|Diameter)[\s:]*(\d+(?:\.\d+)?)[″"]?'
    diam_match = re.search(diam_pattern, dimension_str, re.IGNORECASE)
    if diam_match:
        result['Diameter'] = diam_match.group(1)

    # Extract Weight
    weight_pattern = r'(?:Weight)[\s:]*(\d+(?:\.\d+)?)\s*(?:Ibs?|IBS|lbs?)?'
    weight_match = re.search(weight_pattern, dimension_str, re.IGNORECASE)
    if weight_match:
        result['Weight'] = weight_match.group(1)

    return result


def extract_list_price(description_html):
    """Extract price number only (e.g., $450 pair -> 450)"""
    if not description_html:
        return ""

    # Pattern to match $450 or $450 pair or similar
    pattern = r'\$(\d+(?:,\d+)?)'
    match = re.search(pattern, description_html)

    if match:
        return match.group(1).replace(',', '')

    return ""


def extract_finishes(description_html):
    """Extract all finishes and return as comma-separated string"""
    if not description_html:
        return ""

    soup = BeautifulSoup(description_html, 'html.parser')
    text = soup.get_text()

    finishes = []

    # Find "Finishes:" section
    if 'Finishes:' in text or 'Finish:' in text:
        # Split by "Finishes:" or "Finish:"
        parts = re.split(r'Finishes?:', text, flags=re.IGNORECASE)
        if len(parts) > 1:
            finish_section = parts[1]

            # Split by next major section or end
            finish_section = re.split(r'(\(See Finish|Configurations|Sizes?:|<hr>)', finish_section)[0]

            # Extract lines starting with •
            lines = finish_section.split('\n')
            for line in lines:
                line = line.strip()
                if line.startswith('•'):
                    finish = line.replace('•', '').strip()
                    # Remove any trailing info like "(Shown)"
                    finish = re.sub(r'\s*\(.*?\)\s*$', '', finish)
                    if finish:
                        finishes.append(finish)

    return ', '.join(finishes)


# Lists to store detailed data
detailed_data = []

print("Starting detailed scraping...")

for index, row in df.iterrows():
    product_url = row['Product URL']
    product_name = row['Product Name']
    sku = row['SKU']
    image_url = row['Image URL']

    print(f"\n[{index + 1}/{len(df)}] Scraping: {product_name}")
    print(f"URL: {product_url}")

    try:
        response = requests.get(product_url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract Product Family ID
        product_family_id = get_product_family_id(product_name)

        # Extract Description (plain text from div with itemprop="description")
        description_div = soup.find('div', {'itemprop': 'description'})
        if description_div:
            # Get text content without HTML tags
            description = description_div.get_text(separator='\n', strip=True)
            # Clean up multiple newlines
            description = re.sub(r'\n\s*\n', '\n', description)
        else:
            description = ""

        # Extract Dimension
        dimension = extract_dimension(description)

        # Parse dimensions into separate fields
        parsed_dims = parse_dimensions(dimension)

        # Extract List Price
        list_price = extract_list_price(description)

        # Extract Finishes
        finish = extract_finishes(description)

        detailed_data.append({
            'Product URL': product_url,
            'Image URL': image_url,
            'Product Name': product_name,
            'SKU': sku,
            'Product Family Id': product_family_id,
            'Description': description,
            'Weight': parsed_dims['Weight'],
            'Width': parsed_dims['Width'],
            'Depth': parsed_dims['Depth'],
            'Diameter': parsed_dims['Diameter'],
            'Height': parsed_dims['Height'],
            'List Price': list_price,
            'Finish': finish,
            'Dimension': dimension
        })

        print(f"✓ Product Family ID: {product_family_id}")
        print(f"✓ Width: {parsed_dims['Width']}, Height: {parsed_dims['Height']}, Depth: {parsed_dims['Depth']}")
        print(f"✓ Diameter: {parsed_dims['Diameter']}, Weight: {parsed_dims['Weight']}")
        print(f"✓ List Price: {list_price}")
        print(f"✓ Finish: {finish[:50]}..." if len(finish) > 50 else f"✓ Finish: {finish}")

        time.sleep(1)  # Be polite to the server

    except Exception as e:
        print(f"✗ Error scraping {product_url}: {str(e)}")
        # Add empty data for failed scrapes
        detailed_data.append({
            'Product URL': product_url,
            'Image URL': image_url,
            'Product Name': product_name,
            'SKU': sku,
            'Product Family Id': get_product_family_id(product_name),
            'Description': "",
            'Weight': "",
            'Width': "",
            'Depth': "",
            'Diameter': "",
            'Height': "",
            'List Price': "",
            'Finish': "",
            'Dimension': ""
        })
        time.sleep(1)

# Create DataFrame with proper column order
df_detailed = pd.DataFrame(detailed_data, columns=[
    'Product URL',
    'Image URL',
    'Product Name',
    'SKU',
    'Product Family Id',
    'Description',
    'Weight',
    'Width',
    'Depth',
    'Diameter',
    'Height',
    'List Price',
    'Finish',
    'Dimension'
])

# Save to Excel
output_file = "mclean_lighting_detailed.xlsx"
df_detailed.to_excel(output_file, index=False)

print(f"\n✅ Detailed scraping complete!")
print(f"📊 Total products scraped: {len(detailed_data)}")
print(f"💾 File saved: {output_file}")