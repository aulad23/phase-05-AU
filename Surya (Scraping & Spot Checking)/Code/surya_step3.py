import pandas as pd
import os
import re

# যেই folder এ এই python file আছে
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Input Excel (same folder)
input_file = os.path.join(BASE_DIR, "surya_Pillows_details.xlsx")

# Output Excel (same folder)
output_file = os.path.join(BASE_DIR, "surya_Pillows_details_updated.xlsx")

# Read Excel
df = pd.read_excel(input_file)

# Polyester extract করা
df["Polyester"] = df["Description"].str.extract(
    r'(\d+%\s*Polyester)', expand=False
)

# Save Excel
df.to_excel(output_file, index=False)

print("✅ Output saved in same folder as the code")
print("📁 Location:", BASE_DIR)
