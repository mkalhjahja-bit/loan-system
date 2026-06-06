from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
from flask import Flask, render_template, request, redirect, session, send_file, flash
import sqlite3, os, ast, zipfile, json, subprocess
from io import BytesIO
from num2words import num2words
from docxtpl import DocxTemplate
from PyPDF2 import PdfMerger

app = Flask(__name__)
app.secret_key = "loan123"

BASE = os.path.dirname(os.path.abspath(__file__))
WORD_DIR = os.path.join(BASE, "word_templates")
OUTPUT = os.path.join(BASE, "output")

FOLDER_ID = "1sTAxZNmR-VKw9ULNiaoQWH68lnA05PsV"

os.makedirs(WORD_DIR, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)

import os

info = {
    "type": os.environ["GOOGLE_TYPE"],
    "project_id": os.environ["GOOGLE_PROJECT_ID"],
    "private_key_id": os.environ["GOOGLE_PRIVATE_KEY_ID"],
    "private_key": os.environ["GOOGLE_PRIVATE_KEY"].replace("\\n", "\n"),
    "client_email": os.environ["GOOGLE_CLIENT_EMAIL"],
    "client_id": os.environ["GOOGLE_CLIENT_ID"],
    "token_uri": "https://oauth2.googleapis.com/token"
}
from google.oauth2 import service_account

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]
credentials = service_account.Credentials.from_service_account_info(
    info,
    scopes=SCOPES
)

drive_service = build(
    "drive",
    "v3",
    credentials=credentials
)



# ================= NO CACHE =================

@app.after_request
def add_header(response):

    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"

    return response

# ================= LOGIN =================

@app.route("/", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form.get("username")
        password = request.form.get("password")

        if username == "admin" and password == "1234":

            session["user"] = "admin"

            return redirect("/home")

    return render_template("login.html")

# ================= HOME =================

@app.route("/home")
def home():

    return render_template("home.html")

# ================= PAGES =================

@app.route("/first-loan")
def first_loan():

    return render_template(
        "first_loan.html",
        data=None
    )

@app.route("/continue-loan")
def continue_loan():

    return render_template(
        "continue_loan.html",
        data=None
    )

@app.route("/card")
def card():

    return render_template(
        "card.html",
        data=None
    )

@app.route("/calculator")
def calculator():

    return render_template(
        "calculator.html"
    )

# ================= CLIENTS PAGE =================

@app.route("/clients")
def clients():

    return render_template("clients.html")

# ================= SAVE CLIENT FILE =================

@app.route("/save-client", methods=["POST"])
def save_client():

    data = dict(request.form)

    # ================= amount to words =================

    try:

        amount = (
            data.get("FacilityAmount")
            or data.get("facility_amount")
            or data.get("loan_amount")
            or 0
        )

        number = int(float(str(amount).replace(",", "")))

        data["FacilityAmountWords"] = num2words(
            number,
            lang="ar"
        )

    except:

        data["FacilityAmountWords"] = ""

    # ================= filename =================

    client_name = data.get(
        "ClientName_AR",
        "client"
    ).strip()

    # ================= json =================

    json_data = json.dumps(
        data,
        ensure_ascii=False,
        indent=4
    )

    file_stream = BytesIO()

    file_stream.write(
        json_data.encode("utf-8")
    )

    file_stream.seek(0)

    filename = f"{client_name}.json"

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/json"
    )

# ================= OPEN CLIENT FILE =================

@app.route("/upload-client", methods=["POST"])
def upload_client():

    file = request.files.get("client_file")

    if not file:
        return redirect("/clients")

    data = json.load(file)

    # ================= amount to words =================

    try:

        amount = (
            data.get("FacilityAmount")
            or data.get("facility_amount")
            or data.get("loan_amount")
            or 0
        )

        number = int(float(str(amount).replace(",", "")))

        data["FacilityAmountWords"] = num2words(
            number,
            lang="ar"
        )

    except:

        data["FacilityAmountWords"] = ""

    mode = request.form.get("mode")

    if mode == "first":

        return render_template(
            "first_loan.html",
            data=data
        )

    if mode == "continue":

        return render_template(
            "continue_loan.html",
            data=data
        )

    if mode == "card":

        return render_template(
            "card.html",
            data=data
        )

    return redirect("/clients")

# ================= GENERATE ZIP =================

def generate_zip(data, forms):

    # ================= amount to words =================

    try:

        amount = (
            data.get("FacilityAmount")
            or data.get("facility_amount")
            or data.get("loan_amount")
            or 0
        )

        number = int(float(str(amount).replace(",", "")))

        data["FacilityAmountWords"] = num2words(
            number,
            lang="ar"
        )

    except:

        data["FacilityAmountWords"] = ""

    word_files = []
    pdf_files = []

    # ================= CREATE WORD FILES =================

    for f in forms:

        src = os.path.join(
            WORD_DIR,
            f
        )

        if not os.path.isfile(src):
            continue

        # تعبئة الوورد

        doc = DocxTemplate(src)

        doc.render(data)

        word_path = os.path.join(
         OUTPUT,
            f
        )

        doc.save(word_path)

        word_files.append(word_path)

        # ================= CONVERT WORD TO PDF =================

        try:

            subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                OUTPUT,
                word_path
            ])

            pdf_path = word_path.replace(
                ".docx",
                ".pdf"
            )

            if os.path.exists(pdf_path):

                pdf_files.append(pdf_path)

        except Exception as e:

            print("PDF ERROR:", e)

    # ================= MERGE ALL PDFS =================

    final_pdf = os.path.join(
        OUTPUT,
        "PRINT_ALL.pdf"
    )

    merger = PdfMerger()

    for pdf in pdf_files:

        merger.append(pdf)

    merger.write(final_pdf)

    merger.close()

    # ================= CREATE ZIP =================

    zip_path = os.path.join(
        OUTPUT,
        "forms_result.zip"
    )

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        # ملفات Word

        for w in word_files:

            zipf.write(
                w,
                os.path.basename(w)
            )

        # PDF النهائي

        if os.path.exists(final_pdf):

            zipf.write(
                final_pdf,
                "PRINT_ALL.pdf"
            )

    return zip_path

def upload_to_drive(file_obj, filename):

    file_obj.seek(0)

    file_metadata = {
        "name": filename,
        "parents": [FOLDER_ID]
    }

    media = MediaIoBaseUpload(
        io.BytesIO(file_obj.read()),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    uploaded = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name, parents, webViewLink"
    ).execute()

    print("🔥 UPLOAD RESPONSE:")
    print(uploaded)

    return uploaded
# ================= FIRST LOAN =================

@app.route("/create-first", methods=["POST"])
def create_first():

    forms = [
        "form1.docx",
        "form10.docx"
    ]

    zip_file = generate_zip(
        dict(request.form),
        forms
    )

    return send_file(
        zip_file,
        as_attachment=True
    )

# ================= CONTINUE LOAN =================

@app.route("/create-continue", methods=["POST"])
def create_continue():

    data = dict(request.form)

    forms = [
        "form1.docx",
        "form2.docx",
        "form3.docx",
        "form4.docx",
        "form5.docx",
        "form6.docx",
        "form7.docx",
        "form8.docx",
        "form9.docx",
        "form10.docx",
        "form11.docx"
    ]

    if data.get("debt_card"):

        if "form5.docx" in forms:
            forms.remove("form5.docx")

    else:

        if "form6.docx" in forms:
            forms.remove("form6.docx")

    if not data.get("campaign"):

        if "form7.docx" in forms:
            forms.remove("form7.docx")

    zip_file = generate_zip(
        data,
        forms
    )

    return send_file(
        zip_file,
        as_attachment=True
    )

# ================= CARD =================

@app.route("/create-card", methods=["POST"])
def create_card():

    forms = [
        "form1.docx",
        "form2.docx",
        "form9.docx",
        "form10.docx",
        "form11.docx"
    ]

    zip_file = generate_zip(
        dict(request.form),
        forms
    )

    return send_file(
        zip_file,
        as_attachment=True
    )

# ================= EXCEL FILES =================

@app.route("/upload-excels", methods=["POST"])
def upload_excels():

    file_no = request.form.get("file_no")
    client_name = request.form.get("client_name")

    salary_file = request.files.get("salary_excel")
    debt_file = request.files.get("debt_excel")

    if not salary_file or not debt_file:
        flash("يرجى اختيار الملفين")
        return redirect("/excel-files")

    try:

        upload_to_drive(
            salary_file,
            f"{file_no}_{client_name}_salary.xlsx"
        )

        upload_to_drive(
            debt_file,
            f"{file_no}_{client_name}_debt.xlsx"
        )

        flash("تم رفع الملفات إلى Google Drive بنجاح ✅")

    except Exception as e:

        flash(str(e))

    return redirect("/excel-files")
    
# ================= LOGOUT =================

@app.route("/logout")
def logout():

    session.clear()

    return redirect("/")

@app.route("/excel-files")
def excel_files():
    return render_template("excel_files.html")

# ================= RUN =================

if __name__ == "__main__":

    app.run(debug=True)
