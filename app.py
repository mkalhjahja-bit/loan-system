from flask import Flask, render_template, request, redirect, session, send_file, flash
import sqlite3, os, ast, zipfile, subprocess
from docxtpl import DocxTemplate
from pypdf import PdfMerger

app = Flask(__name__)
app.secret_key = "loan123"

# ================= PATHS =================

BASE = os.path.dirname(os.path.abspath(__file__))
WORD_DIR = os.path.join(BASE, "word_templates")
OUTPUT = os.path.join(BASE, "output")

os.makedirs(WORD_DIR, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)

DB = os.path.join(BASE, "database.db")

# ================= DATABASE =================

def db():
    return sqlite3.connect(DB)

with db() as con:
    con.execute("""
    CREATE TABLE IF NOT EXISTS clients(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        data TEXT
    )
    """)

# ================= LOGIN =================

@app.route("/", methods=["GET","POST"])
def login():
    if request.method == "POST":
        if request.form.get("username") == "admin" and request.form.get("password") == "1234":
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
    rows = db().execute("SELECT id,name FROM clients ORDER BY id DESC").fetchall()
    return render_template("clients.html", rows=rows)

@app.route("/delete-client/<int:id>")
def delete_client(id):
    with db() as con:
        con.execute("DELETE FROM clients WHERE id=?", (id,))
    return redirect("/clients")

# ================= SAVE CLIENT =================

@app.route("/save-client", methods=["POST"])
def save_client():
    name = request.form.get("ClientName_AR","")
    data = str(dict(request.form))

    with db() as con:
        con.execute("INSERT INTO clients(name,data) VALUES(?,?)",(name,data))

    flash("تم حفظ العميل ✅")
    return redirect("/clients")

# ================= LOAD CLIENT =================

@app.route("/load-client/<int:id>/<mode>")
def load_client(id, mode):

    row = db().execute("SELECT data FROM clients WHERE id=?", (id,)).fetchone()
    if not row:
        return redirect("/clients")

    data = ast.literal_eval(row[0])

    if mode == "first":
        return render_template("first_loan.html", data=data)

    if mode == "continue":
        return render_template("continue_loan.html", data=data)

    if mode == "card":
        return render_template("card.html", data=data)

    return redirect("/clients")

# ================= ZIP GENERATOR =================

def generate_zip(data, forms):

    word_files = []
    pdf_files = []

    for f in forms:

        src = os.path.join(WORD_DIR, f)

        if not os.path.isfile(src):
            print("Missing:", f)
            continue

        doc = DocxTemplate(src)
        doc.render(data)

        word_path = os.path.join(OUTPUT, f)
        doc.save(word_path)

        word_files.append(word_path)

        # تحويل Word إلى PDF باستخدام LibreOffice
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            word_path,
            "--outdir",
            OUTPUT
        ])

        pdf_path = word_path.replace(".docx",".pdf")

        if os.path.isfile(pdf_path):
            pdf_files.append(pdf_path)

    # دمج ملفات PDF
    merger = PdfMerger()

    for pdf in pdf_files:
        merger.append(pdf)

    final_pdf = os.path.join(OUTPUT,"PRINT_ALL.pdf")

    if pdf_files:
        merger.write(final_pdf)
        merger.close()

    # إنشاء ZIP
    zip_path = os.path.join(OUTPUT, "forms_result.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:

        for w in word_files:
            zipf.write(w, os.path.basename(w))

        if os.path.isfile(final_pdf):
            zipf.write(final_pdf,"PRINT_ALL.pdf")

    return zip_path

# ================= FIRST LOAN =================

@app.route("/create-first", methods=["POST"])
def create_first():

    forms = [
        "form1.docx",
        "form10.docx"
    ]

    zip_file = generate_zip(dict(request.form), forms)
    return send_file(zip_file, as_attachment=True)

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

    zip_file = generate_zip(data, forms)
    return send_file(zip_file, as_attachment=True)

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
