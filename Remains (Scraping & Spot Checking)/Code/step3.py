# remains_specs_splitter.py
import re
import pandas as pd

# =============================
# CONFIG — edit if you like
# =============================
INPUT_XLSX  = "remains_flush-mounts_detils.xlsx"      # your Step-2 output file
OUTPUT_XLSX = "remains_flush-mounts_file.xlsx"

# Subsections to extract from the "Specifications" column (case-insensitive)
TARGET_HEADINGS = [
    "Dimensions",
    "Application",
    "Lamping",
    "Junction Box",
]

# =============================
# Helpers
# =============================
def normalize_text(s):
    if s is None:
        return ""
    # Ensure string & normalize newlines
    s = str(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return s.strip()

def parse_spec_sections(spec_text, headings):
    """
    Parse the 'Specifications' blob into sections keyed by the headings list.
    Accepts headings written as a standalone line (with optional colon) and
    collects all following lines until the next heading or end of text.
    Also supports content on the same line after the colon.

    Returns: dict {Heading: text or ""} for all provided headings.
    """
    out = {h: "" for h in headings}
    text = normalize_text(spec_text)
    if not text:
        return out

    # Build a single regex that matches any heading at start of a line.
    # Example match lines:
    # "Dimensions:" / "Dimensions" / "Dimensions: Overall: ..."
    head_alts = "|".join(re.escape(h) for h in headings)
    pattern = re.compile(rf"(?im)^\s*({head_alts})\s*:?\s*(.*)$")

    # Find all heading matches with their positions
    matches = list(pattern.finditer(text))
    if not matches:
        return out

    for i, m in enumerate(matches):
        name = m.group(1)  # the canonical heading as written in text
        inline_content = m.group(2).strip()  # content on same line after heading, if any
        start = m.end()  # start of following lines after heading line
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        block_after = text[start:end].strip()

        # Combine inline content + following block (if both exist)
        section_text = inline_content
        if block_after:
            section_text = (section_text + "\n" + block_after).strip() if section_text else block_after

        # Store under the exact target key case from TARGET_HEADINGS
        # (e.g., prefer "Dimensions" capitalization exactly)
        for canon in headings:
            if canon.lower() == name.lower():
                out[canon] = section_text
                break

    return out

# =============================
# Main
# =============================
def main():
    df = pd.read_excel(INPUT_XLSX)

    if "Specifications" not in df.columns:
        raise ValueError("No 'Specifications' column found in the input file.")

    # Create new columns (if they already exist, we'll overwrite with parsed values)
    for col in TARGET_HEADINGS:
        if col not in df.columns:
            df[col] = ""

    # Parse each row's Specifications
    for i, spec in enumerate(df["Specifications"].tolist()):
        sections = parse_spec_sections(spec, TARGET_HEADINGS)
        for col in TARGET_HEADINGS:
            df.at[i, col] = sections.get(col, "")

    # Write out new file with all original columns + new ones
    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"Done. Wrote {len(df)} rows to '{OUTPUT_XLSX}'.")
    print(f"Added/updated columns: {', '.join(TARGET_HEADINGS)}")

if __name__ == "__main__":
    main()
