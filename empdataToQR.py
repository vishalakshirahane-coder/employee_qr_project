import pandas as pd
import qrcode
from qrcode.constants import ERROR_CORRECT_L
import os
import re
import shutil
import subprocess
from weasyprint import HTML
import hashlib


# =====================================================
# CONFIG
# =====================================================

OUTPUT_FOLDER = "QR_Output"
HASH_FILE = ".excel_md5"

GIT_BRANCH = "main"
COMMIT_MSG = "- Deploying Updated QR codes of employees"

BASE_URL = "https://employee-qr-project.vercel.app/QR_Output"



# =====================================================
# UTILITY: CALCULATE FILE MD5
# =====================================================

def calculate_file_md5(file_path):

    md5 = hashlib.md5()

    with open(file_path, "rb") as f:

        for chunk in iter(lambda: f.read(8192), b""):
            md5.update(chunk)

    return md5.hexdigest()


# =====================================================
# STEP 1: FIND EXCEL FILE
# =====================================================

excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

if len(excel_files) == 0:
    raise FileNotFoundError("❌ No Excel file found")

if len(excel_files) > 1:
    raise Exception("❌ Keep only ONE Excel file")

excel_file = excel_files[0]

print(f"📄 Using Excel: {excel_file}")



# =====================================================
# STEP 1.1.: CHECK IF EXCEL CHANGED
# =====================================================

print("\n🔍 Checking Excel file hash...")

current_md5 = calculate_file_md5(excel_file)


# Read old hash
if os.path.exists(HASH_FILE):

    with open(HASH_FILE, "r") as f:
        old_md5 = f.read().strip()

else:
    old_md5 = None


print("📌 Old MD5:", old_md5)
print("📌 New MD5:", current_md5)


# If no change → exit
if current_md5 == old_md5:

    print("\nℹ️ Excel not changed. Nothing to do.")
    exit(0)


print("\n✅ Excel changed. Regenerating files...")


# =====================================================
# STEP 2: CLEAN OLD QR_Output
# =====================================================

if os.path.exists(OUTPUT_FOLDER):
    print("🧹 Removing old QR_Output folder...")
    shutil.rmtree(OUTPUT_FOLDER)

os.makedirs(OUTPUT_FOLDER)
print("✅ Fresh QR_Output created")


# =====================================================
# STEP 3: READ EXCEL
# =====================================================

df = pd.read_excel(excel_file)


# =====================================================
# STEP 4: FIND NAME COLUMN
# =====================================================
name_column = None

for col in df.columns:
    col_lower = str(col).lower()
    if (
        "name" in col_lower
        or "नाव" in col
    ):
        name_column = col
        break

if name_column:
    print(f"✅ Name column detected: {name_column}")
else:
    print("⚠️ Name column not found. Using auto-generated Employee IDs.")



# =====================================================
# STEP 5: HTML TEMPLATE
# =====================================================

html_template = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{name}</title>

<style>
body {{
    font-family: Arial;
    max-width: 600px;
    margin: 40px auto;
    padding: 20px;
    border: 1px solid #ccc;
    border-radius: 10px;
}}

h1 {{
    text-align: center;
}}

table {{
    width: 100%;
    border-collapse: collapse;
}}

th, td {{
    border: 1px solid #ddd;
    padding: 8px;
}}

th {{
    background: #f2f2f2;
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


# =====================================================
# STEP 6: GENERATE FILES
# =====================================================

print("\n⚙️ Generating QR files...\n")


for index, row in df.iterrows():

    # Get employee name safely
    if name_column:
        employee_name = str(row[name_column]).strip()

        if employee_name == "" or employee_name.lower() == "nan":
            employee_name = f"Participant_{index+1}"
    else:
        employee_name = f"Participant_{index+1}"


    # Safe filename
    safe_name = re.sub(r"[^\w\d-]", "_", employee_name)


    # Build rows
    rows_html = ""

    for col in df.columns:

        value = row[col]

        if pd.notna(value):
            rows_html += f"<tr><th>{col}</th><td>{value}</td></tr>\n"


    # Create HTML
    html_content = html_template.format(
        name=employee_name,
        rows=rows_html
    )


    # Save HTML
    html_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.html")

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)


    # Create PDF
    pdf_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.pdf")

    HTML(string=html_content).write_pdf(pdf_path)


    # Create QR
    employee_url = f"{BASE_URL}/{safe_name}.html"


    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_L,
        box_size=4,
        border=2
    )

    qr.add_data(employee_url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    qr_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.png")

    img.save(qr_path)


    print(f"✅ Created: {safe_name}")


print("\n🎉 All QR files generated")


# =====================================================
# STEP 7: SAVE NEW HASH
# =====================================================

with open(HASH_FILE, "w") as f:
    f.write(current_md5)


print("💾 Excel hash updated")

# =====================================================
# STEP 8: AUTO GIT PUSH
# =====================================================

def auto_git_push():

    try:

        print("\n🚀 Pushing to GitHub...")


        subprocess.run(
            ["git", "add", OUTPUT_FOLDER, HASH_FILE],
            check=True
        )


        subprocess.run(
            ["git", "commit", "-m", COMMIT_MSG],
            check=True
        )


        subprocess.run(
            ["git", "push", "origin", GIT_BRANCH],
            check=True
        )


        print("✅ Git push successful")


    except subprocess.CalledProcessError as e:

        print("❌ Git push failed")
        print(e)

auto_git_push()
