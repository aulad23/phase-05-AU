import pandas as pd
from pathlib import Path


INPUT_XLSX = Path.home() / "Downloads" / "HVAC_FL_1to10_2025-09-07_09-22-24-480.xlsx"
SECOND_XLSX = Path("D:/SC-AU/Resul_first.xlsx")

main_df = pd.read_excel(INPUT_XLSX)
resul_df = pd.read_excel(SECOND_XLSX)
resul_df = resul_df.iloc[:, 1:]  

url_column_candidates = [col for col in main_df.columns if "url" in col.lower()]
if not url_column_candidates:
    raise ValueError("No URL column found in the main DataFrame.")
main_url_col = url_column_candidates[0]

resul_df = resul_df.drop_duplicates(subset=["URL"])

main_df[main_url_col] = main_df[main_url_col].astype(str).str.strip()
resul_df["URL"] = resul_df["URL"].astype(str).str.strip()

merged_df = main_df.merge(
    resul_df,
    left_on=main_url_col,
    right_on="URL",
    how="left"
)
if "URL" in merged_df.columns:
    merged_df = merged_df.drop(columns=["URL"])

output_file = Path.home() / "Downloads" / "HVAC_FL_1to10_99_66_updated.xlsx"
merged_df.to_excel(output_file, index=False)

print(f"✅ Merge complete. Updated file saved at: {output_file}")
