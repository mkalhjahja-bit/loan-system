from flask import Flask, render_template, request, redirect, session, send_file
import os, zipfile, json, subprocess
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

# ================= GENERATE ZIP =================

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

    # ================= CREATE WORD + PDF =================

    for f in forms:

        src = os.path.join(WORD_DIR, f)

        if not os.path.isfile(src):
            continue

        # Word file
        doc = DocxTemplate(src)
        doc.render(data)

        word_path = os.path.join(OUTPUT, f)
        doc.save(word_path)

        word_files.append(word_path)

        # ================= WORD -> PDF =================

        try:
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

            base_name = os.path.splitext(os.path.basename(word_path))[0]
            pdf_path = os.path.join(OUTPUT, base_name + ".pdf")

            if os.path.exists(pdf_path):
                pdf_files.append(pdf_path)
            else:
                print("❌ PDF NOT CREATED:", word_path)

        except Exception as e:
            print("PDF ERROR:", e)

    # ================= PREVENT SERVER CRASH =================

    if not pdf_files:
        raise Exception("No PDFs were generated. Check LibreOffice installation or conversion.")

    # ================= MERGE PDFs =================

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

    zip_file = generate_zip(data, forms)

    return send_file(zip_file, as_attachment=True)


@app.route("/create-card", methods=["POST"])
def create_card():

    forms = ["form1.docx","form2.docx","form9.docx","form10.docx","form11.docx"]

    zip_file = generate_zip(dict(request.form), forms)

    return send_file(zip_file, as_attachment=True)


# ================= BASIC PAGES =================

@app.route("/")
def home():
    return "System Running"

# ================= RUN =================

if __name__ == "__main__":
    app.run(debug=True)
