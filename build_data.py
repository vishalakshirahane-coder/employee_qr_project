import pandas as pd
import json
import qrcode
import os

BASE_URL = "https://YOUR_GITHUB_USERNAME.github.io/employee_qr_project/view.html"
QR_FOLDER = "QR_Output"

os.makedirs(QR_FOLDER, exist_ok=True)

# Auto-detect Excel
excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
if len(excel_files) != 1:
    raise Exception("Keep exactly ONE Excel file")

excel_file = excel_files[0]
print(f"Using Excel: {excel_file}")

df = pd.read_excel(excel_file)

# Convert Excel → JSON (preserve Marathi / English headers)
records = df.fillna("").to_dict(orient="records")

with open("employees.json", "w", encoding="utf-8") as f:
    json.dump(records, f, ensure_ascii=False, indent=2)

print("employees.json created")

# Generate QR codes
for i in range(len(records)):
    qr_url = f"{BASE_URL}?id={i}"
    img = qrcode.make(qr_url)
    img.save(f"{QR_FOLDER}/employee_{i}.png")

print("All QR codes generated")
