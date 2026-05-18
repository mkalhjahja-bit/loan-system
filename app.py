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

os.makedirs(WORD_DIR, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)

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
    return render_template("first_loan.html", data=None)

@app.route("/continue-loan")
def continue_loan():
    return render_template("continue_loan.html", data=None)

@app.route("/card")
def card():
    return render_template("card.html", data=None)

@app.route("/calculator")
def calculator():
    return render_template("calculator.html")

# ================= CLIENTS =================

@app.route("/clients")
def clients():
    return render_template("clients.html")

# ================= SAVE CLIENT =================

@app.route("/save-client", methods=["POST"])
def save_client():
    data = dict(request.form)

    try:
        amount = (
            data.get("FacilityAmount")
            or data.get("facility_amount")
            or data.get("loan_amount")
            or 0
        )

        number = int(float(str(amount).replace(",", "")))

        data["FacilityAmountWords"] = num2words(number, lang="ar")

    except:
        data["FacilityAmountWords"] = ""

    client_name = data.get("ClientName_AR", "client").strip()

    json_data = json.dumps(data, ensure_ascii=False, indent=4)

    file_stream = BytesIO()
    file_stream.write(json_data.encode("utf-8"))
    file_stream.seek(0)

    filename = f"{client_name}.json"

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/json"
    )

# ================= UPLOAD CLIENT =================

@app.route("/upload-client", methods=["POST"])
def upload_client():
    file = request.files.get("client_file")

    if not file:
        return redirect("/clients")

    data = json.load(file)

    try:
        amount = (
            data.get("FacilityAmount")
            or data.get("facility_amount")
            or data.get("loan_amount")
            or 0
        )

        number = int(float(str(amount).replace(",", "")))

        data["FacilityAmountWords"] = num2words(number, lang="ar")

    except:
        data["FacilityAmountWords"] = ""

    mode = request.form.get("mode")

    if mode == "first":
        return render_template("first_loan.html", data=data)

    if mode == "continue":
        return render_template("continue_loan.html", data=data)

    if mode == "card":
        return render_template("card.html", data=data)

    return redirect("/clients")

# ================= GENERATE ZIP (FIXED) =================

def generate_zip(data, forms):

    try:
        amount = (
            data.get("FacilityAmount")
            or data.get("facility_amount")
            or data.get("loan_amount")
            or 0
        )

        number = int(float(str(amount).replace(",", "")))

        data["FacilityAmountWords"] = num2words(number, lang="ar")

    except:
        data["FacilityAmountWords"] = ""

    word_files = []
    pdf_files = []

    # ================= CREATE WORD FILES =================

    for f in forms:

        src = os.path.join(WORD_DIR, f)

        if not os.path.isfile(src):
            continue

        doc = DocxTemplate(src)
        doc.render(data)

        word_path = os.path.join(OUTPUT, f)
        doc.save(word_path)

        word_files.append(word_path)

        # ================= CONVERT TO PDF (FIXED) =================

        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                OUTPUT,
                word_path
            ],
            capture_output=True,
            text=True
        )

        print(result.stdout)
        print(result.stderr)

        # 🔥 FIX: correct PDF name detection
        base_name = os.path.splitext(os.path.basename(word_path))[0]
        pdf_path = os.path.join(OUTPUT, base_name + ".pdf")

        if os.path.exists(pdf_path):
            pdf_files.append(pdf_path)

    # ================= MERGE PDFs =================

    print("PDF FILES READY:", pdf_files)

    final_pdf = os.path.join(OUTPUT, "PRINT_ALL.pdf")

    merger = PdfMerger()

    for pdf in pdf_files:
        merger.append(pdf)

    merger.write(final_pdf)
    merger.close()

    # ================= CREATE ZIP =================

    zip_path = os.path.join(OUTPUT, "forms_result.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:

        for w in word_files:
            zipf.write(w, os.path.basename(w))

        if os.path.exists(final_pdf):
            zipf.write(final_pdf, "PRINT_ALL.pdf")

    return zip_path

# ================= ROUTES =================

@app.route("/create-first", methods=["POST"])
def create_first():

    forms = ["form1.docx", "form10.docx"]

    zip_file = generate_zip(dict(request.form), forms)

    return send_file(zip_file, as_attachment=True)


@app.route("/create-continue", methods=["POST"])
def create_continue():

    data = dict(request.form)

    forms = [
        "form1.docx","form2.docx","form3.docx","form4.docx",
        "form5.docx","form6.docx","form7.docx","form8.docx",
        "form9.docx","form10.docx","form11.docx"
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

    zip_file = generate_zip(data, forms)

    return send_file(zip_file, as_attachment=True)


@app.route("/create-card", methods=["POST"])
def create_card():

    forms = ["form1.docx","form2.docx","form9.docx","form10.docx","form11.docx"]

    zip_file = generate_zip(dict(request.form), forms)

    return send_file(zip_file, as_attachment=True)


# ================= LOGOUT =================

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


# ================= RUN =================

if __name__ == "__main__":
    app.run(debug=True)
