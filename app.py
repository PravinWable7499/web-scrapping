import pandas as pd
import json

df = pd.read_excel("hospital_data_midc.xlsx")

# Clean column names
df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

# Convert to JSON
data = df.to_dict(orient="records")

with open("companies.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=4, ensure_ascii=False)

print("JSON file created!")