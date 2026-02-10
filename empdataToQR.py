import pandas as pd
import qrcode
from qrcode.constants import ERROR_CORRECT_L
import os
import re
import shutil
import subprocess
from weasyprint import HTML


# =====================================================
# CONFIG
# =====================================================

OUTPUT_FOLDER = "QR_Output"
GIT_BRANCH = "main"
COMMIT_MSG = "- Deploying Updated QR codes of employees"

BASE_URL = "https://employee-qr-project.vercel.app/QR_Output"


# =====================================================
# STEP 1: CLEAN OLD QR_Output
# =====================================================

if os.path.exists(OUTPUT_FOLDER):
    print("🧹 Removing old QR_Output folder...")
    shutil.rmtree(OUTPUT_FOLDER)

os.makedirs(OUTPUT_FOLDER)
print("✅ Fresh QR_Output created")


# =====================================================
# STEP 2: FIND EXCEL FILE
# =====================================================

excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

if len(excel_files) == 0:
    raise FileNotFoundError("❌ No Excel file found")

if len(excel_files) > 1:
    raise Exception("❌ Keep only ONE Excel file")

excel_file = excel_files[0]

print(f"📄 Using Excel: {excel_file}")


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

'''
name_column = None

for col in df.columns:
    col_lower = str(col).lower()

    if "name" in col_lower or "नाव" in col or "संपूर्ण" in col:
        name_column = col
        break

if name_column is None:
    raise Exception("❌ Name column not found")

print(f"✅ Name column: {name_column}")

'''

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

for index, row in df.iterrows():

    # Get employee name safely
    if name_column:
        employee_name = str(row[name_column]).strip()

        if employee_name == "" or employee_name.lower() == "nan":
            employee_name = f"Employee_{index+1}"
    else:
        employee_name = f"Employee_{index+1}"


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
# STEP 7: AUTO GIT PUSH
# =====================================================

def auto_git_push():

    try:

        print("\n🔍 Checking for changes in QR_Output...")

        # Check if QR_Output has changes
        result = subprocess.run(
            ["git", "status", "--porcelain", OUTPUT_FOLDER],
            capture_output=True,
            text=True
        )

        if result.stdout.strip() == "":
            print("ℹ️ No changes detected. Nothing to commit.")
            return


        print("✅ Changes found. Pushing to GitHub...")


        subprocess.run(
            ["git", "add", OUTPUT_FOLDER],
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

        print("🚀 Git Push Successful")


    except subprocess.CalledProcessError as e:

        print("❌ Git Push Failed")
        print(e)


auto_git_push()
