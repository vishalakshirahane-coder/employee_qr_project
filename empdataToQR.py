import pandas as pd
import qrcode
from qrcode.constants import ERROR_CORRECT_L
import os
import re
from weasyprint import HTML

output_folder = "QR_Output"
os.makedirs(output_folder, exist_ok=True)

# Find Excel file automatically
excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

if len(excel_files) == 0:
    raise FileNotFoundError("❌ No Excel file found in current folder.")
elif len(excel_files) > 1:
    raise Exception("❌ Multiple Excel files found. Keep only ONE Excel file.")

excel_file = excel_files[0]
print(f"📄 Using Excel file: {excel_file}")

# Read Excel
df = pd.read_excel(excel_file)

# Detect NAME column (Marathi + English)
name_column = None
for col in df.columns:
    col_lower = str(col).lower()
    if "name" in col_lower or "नाव" in col or "संपूर्ण" in col:
        name_column = col
        break

if name_column is None:
    raise Exception("❌ Name column not found.")

print(f"✅ Detected name column: {name_column}")

# HTML template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{name} - Employee Info</title>
<style>
body {{
    font-family: Arial, sans-serif;
    max-width: 600px;
    margin: 50px auto;
    padding: 20px;
    border: 1px solid #ccc;
    border-radius: 12px;
}}
h1 {{
    text-align: center;
    color: #2c3e50;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 20px;
}}
td, th {{
    padding: 10px;
    border: 1px solid #ddd;
}}
th {{
    background-color: #f4f4f4;
    text-align: left;
}}
</style>
</head>
<body>
<h1>{name}</h1>
<table>
{rows}
</table>
</body>
</html>
"""

# Generate QR, HTML, PDF per employee
for _, row in df.iterrows():
    employee_name = str(row[name_column]).strip()
    if employee_name == "" or employee_name.lower() == "nan":
        continue

    # Safe filename
    safe_name = re.sub(r"[^\w\d-]", "_", employee_name)

    # Avoid overwrite
    count = 1
    original_name = safe_name
    while os.path.exists(os.path.join(output_folder, f"{safe_name}.html")):
        safe_name = f"{original_name}_{count}"
        count += 1

    # Build HTML rows
    rows_html = ""
    for col in df.columns:
        value = row[col]
        if pd.notna(value):
            rows_html += f"<tr><th>{col}</th><td>{value}</td></tr>\n"

    html_content = html_template.format(
        name=employee_name,
        rows=rows_html
    )

    # Save HTML
    html_file_path = os.path.join(output_folder, f"{safe_name}.html")
    with open(html_file_path, "w", encoding="utf-8") as f:
       f.write(html_content)

    # ✅ CONVERT HTML → PDF (ADDED)
    pdf_file_path = os.path.join(output_folder, f"{safe_name}.pdf")
    HTML(string=html_content).write_pdf(pdf_file_path)

    # Generate QR (still points to local HTML)
    '''qr_text = "Employee Details\n"
    qr_text += "-" * 25 + "\n\n"

    for col in df.columns:
        value = row[col]
        if pd.notna(value):
            qr_text += f"• {col}: {value}\n\n"
            
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_L,
        box_size=3,
        border=2
    )

    qr.add_data(qr_text.strip())
    qr.make(fit=True)
    '''

    # Your GitHub Pages base URL
    #BASE_URL = "https://vishalakshirahane-coder.github.io/employee_qr_project/QR_Output"
    BASE_URL = "https://employee-qr-project.vercel.app/QR_Output"

    employee_url = f"{BASE_URL}/{safe_name}.html"

    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_L,
        box_size=3,
        border=2
    )

    qr.add_data(employee_url)
    qr.make(fit=True)


    img_text_qr = qr.make_image(
        fill_color="black",
        back_color="white"
    )

    qr_file_path = os.path.join(
        output_folder,
        f"{safe_name}.png"
    )
    img_text_qr.save(qr_file_path)

    print(f"✅ HTML: {html_file_path}")
    print(f"✅ PDF : {pdf_file_path}")
    print(f"✅ QR  : {qr_file_path}")

print("🎉 PDF + QR generation completed!")
