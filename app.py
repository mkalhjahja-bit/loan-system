from flask import Flask, render_template, request, redirect, session, flash, send_file
import sqlite3, os, ast, zipfile
from docxtpl import DocxTemplate

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
        if request.form.get("username")=="admin" and request.form.get("password")=="1234":
            session["user"]="admin"
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

    flash("✅ تم حفظ العميل")
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

# ================= WORD → ZIP ENGINE =================

def generate_and_zip(data, forms):

    zip_path = os.path.join(OUTPUT, "loan_forms.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:

        for f in forms:
            template = os.path.join(WORD_DIR, f)

            if not os.path.exists(template):
                print("Missing:", template)
                continue

            doc = DocxTemplate(template)
            doc.render(data)

            output_file = os.path.join(OUTPUT, f)
            doc.save(output_file)

            zipf.write(output_file, f)

    return zip_path

# ================= FIRST LOAN =================

@app.route("/create-first", methods=["POST"])
def create_first():

    zip_file = generate_and_zip(dict(request.form), [
        "form1.docx",
        "form10.docx"
    ])

    return send_file(zip_file, as_attachment=True)

# ================= CONTINUE LOAN =================

@app.route("/create-continue", methods=["POST"])
def create_continue():

    data = dict(request.form)

    forms = [
        "form1.docx",
        "form3.docx",
        "form4.docx",
        "form5.docx",
        "form6.docx",
        "form7.docx",
        "form8.docx",
        "form10.docx"
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

    zip_file = generate_and_zip(data, forms)

    return send_file(zip_file, as_attachment=True)

# ================= CARD =================

@app.route("/create-card", methods=["POST"])
def create_card():

    zip_file = generate_and_zip(dict(request.form), [
        "form1.docx",
        "form2.docx",
        "form9.docx",
        "form10.docx",
        "form11.docx"
    ])

    return send_file(zip_file, as_attachment=True)

# ================= LOGOUT =================

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ================= RUN =================

if __name__ == "__main__":
    app.run(debug=True)
