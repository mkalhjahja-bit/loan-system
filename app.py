from flask import Flask, render_template, request, redirect, session, send_file
import sqlite3, os, ast
import pythoncom   # ✅ حل مشكلة COM
from docxtpl import DocxTemplate
from docx2pdf import convert
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
app.secret_key = "loan123"

BASE = os.path.dirname(os.path.abspath(__file__))
WORD_DIR = os.path.join(BASE, "word_templates")
OUTPUT = os.path.join(BASE, "output")
os.makedirs(OUTPUT, exist_ok=True)

DB = os.path.join(BASE, "database.db")

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

# ---------------- LOGIN ----------------

@app.route("/", methods=["GET","POST"])
def login():
    if request.method=="POST":
        if request.form["username"]=="admin" and request.form["password"]=="1234":
            session["user"]="admin"
            return redirect("/home")
    return render_template("login.html")

# ---------------- HOME ----------------

@app.route("/home")
def home():
    return render_template("home.html")

# ---------------- PAGES ----------------

@app.route("/first-loan")
def first_loan():
    return render_template("first_loan.html", data=None)

@app.route("/continue-loan")
def continue_loan():
    return render_template("continue_loan.html", data=None)

# ================== قائمة العملاء ==================

@app.route("/clients")
def clients():
    con = db()
    rows = con.execute("SELECT id,name FROM clients").fetchall()
    con.close()
    return render_template("clients.html", rows=rows)

@app.route("/delete-client/<int:id>")
def delete_client(id):
    with db() as con:
        con.execute("DELETE FROM clients WHERE id=?", (id,))
    return redirect("/clients")

# ---------------- CLIENT SAVE ----------------

@app.route("/save-client", methods=["POST"])
def save_client():
    name = request.form.get("ClientName_AR","")
    data = str(dict(request.form))
    with db() as con:
        con.execute("INSERT INTO clients(name,data) VALUES(?,?)",(name,data))
    return redirect("/clients")

# ---------------- CLIENT LOAD ----------------

@app.route("/load-client/<int:id>/<mode>")
def load_client(id, mode):
    row = db().execute("SELECT data FROM clients WHERE id=?", (id,)).fetchone()
    if not row:
        return redirect("/clients")

    data = ast.literal_eval(row[0])

    if mode == "first":
        return render_template("first_loan.html", data=data)

    return render_template("continue_loan.html", data=data)

@app.route("/open-first/<int:id>")
def open_first(id):
    row = db().execute("SELECT data FROM clients WHERE id=?", (id,)).fetchone()
    data = ast.literal_eval(row[0])
    return render_template("first_loan.html", data=data)

@app.route("/open-continue/<int:id>")
def open_continue(id):
    row = db().execute("SELECT data FROM clients WHERE id=?", (id,)).fetchone()
    data = ast.literal_eval(row[0])
    return render_template("continue_loan.html", data=data)

# ================== حاسبة القروض ==================

@app.route("/calculator", methods=["GET","POST"])
def calculator():
    table = []

    if request.method == "POST":
        try:
            P = float(request.form["amount"])
            r = float(request.form["rate"]) / 100 / 12
            n = int(request.form["months"])

            pay = P * (r*(1+r)**n) / ((1+r)**n - 1)
            bal = P

            for i in range(1, n+1):
                interest = bal * r
                principal = pay - interest
                bal -= principal
                table.append((i, round(pay,2), round(interest,2),
                              round(principal,2), round(bal,2)))
        except:
            pass

    return render_template("calculator.html", table=table)

# ---------------- PDF ENGINE ----------------

def make_pdf(data, forms):
    pythoncom.CoInitialize()   # ✅ الحل النهائي للخطأ

    writer = PdfWriter()

    for f in forms:
        doc = DocxTemplate(os.path.join(WORD_DIR, f))
        doc.render(data)

        d = os.path.join(OUTPUT, "_" + f)
        p = d.replace(".docx",".pdf")

        doc.save(d)
        convert(d,p)

        for page in PdfReader(p).pages:
            writer.add_page(page)

    final = os.path.join(OUTPUT, "final.pdf")
    with open(final,"wb") as out:
        writer.write(out)

    return final

# ---------------- FIRST LOAN ----------------

@app.route("/create-first", methods=["POST"])
def create_first():
    pdf = make_pdf(dict(request.form), ["form1.docx","form10.docx"])
    return send_file(pdf, as_attachment=True)

# ---------------- CONTINUE LOAN ----------------

@app.route("/create-continue", methods=["POST"])
def create_continue():
    data = dict(request.form)

    forms = [f for f in os.listdir(WORD_DIR) if f.endswith(".docx")]

    if data.get("debt_card"):
        if "form5.docx" in forms: forms.remove("form5.docx")
    else:
        if "form6.docx" in forms: forms.remove("form6.docx")

    if not data.get("campaign"):
        if "form7.docx" in forms: forms.remove("form7.docx")

    pdf = make_pdf(data, forms)
    return send_file(pdf, as_attachment=True)


@app.route("/card-request")
def card_request():
    return render_template("card_request.html", data=None)


@app.route("/create-card", methods=["POST"])
def create_card():
    data = dict(request.form)

    forms = [
        "form1.docx",
        "form2.docx",
        "form9.docx",
        "form10.docx",
        "form11.docx"
    ]

    pdf = make_pdf(data, forms)
    return send_file(pdf, as_attachment=True)

@app.route("/load-client/<int:id>/card")
def open_card(id):
    row = db().execute("SELECT data FROM clients WHERE id=?", (id,)).fetchone()
    data = ast.literal_eval(row[0])
    return render_template("card_request.html", data=data)

# ---------------- LOGOUT ----------------

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ---------------- RUN ----------------

app.run(debug=True)
