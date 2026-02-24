from flask import Flask, render_template, request, redirect, session, flash, send_file
import sqlite3, os, ast
from docxtpl import DocxTemplate
from docx import Document
from copy import deepcopy

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
    name = request.form.get("ClientName_AR", "")
    data = str(dict(request.form))

    with db() as con:
        con.execute("INSERT INTO clients(name,data) VALUES(?,?)", (name, data))

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

# ================= WORD MERGER (FINAL CLEAN) =================

def generate_docs(data, forms):

    if not forms:
        return None

    # أول نموذج كأساس
    base_path = os.path.join(WORD_DIR, forms[0])

    base_tpl = DocxTemplate(base_path)
    base_tpl.render(data)

    temp_base = os.path.join(OUTPUT, "_base.docx")
    base_tpl.save(temp_base)

    final_doc = Document(temp_base)

    for f in forms[1:]:

        src = os.path.join(WORD_DIR, f)
        if not os.path.isfile(src):
            continue

        tpl = DocxTemplate(src)
        tpl.render(data)

        temp = os.path.join(OUTPUT, "_temp.docx")
        tpl.save(temp)

        sub_doc = Document(temp)

        final_doc.add_page_break()

        for element in sub_doc.element.body:
            final_doc.element.body.append(deepcopy(element))

    final_path = os.path.join(OUTPUT, "final_forms.docx")
    final_doc.save(final_path)

    return final_path

# ================= FIRST LOAN =================

@app.route("/create-first", methods=["POST"])
def create_first():
    file = generate_docs(dict(request.form), [
        "form1.docx",
        "form10.docx"
    ])
    return send_file(file, as_attachment=True)

# ================= CONTINUE LOAN =================

@app.route("/create-continue", methods=["POST"])
def create_continue():

    data = dict(request.form)

    forms = [
        "form1.docx","form3.docx","form4.docx",
        "form5.docx","form6.docx","form7.docx",
        "form8.docx","form10.docx"
    ]

    if data.get("debt_card"):
        forms.remove("form5.docx")
    else:
        forms.remove("form6.docx")

    if not data.get("campaign"):
        forms.remove("form7.docx")

    file = generate_docs(data, forms)
    return send_file(file, as_attachment=True)

# ================= CARD =================

@app.route("/create-card", methods=["POST"])
def create_card():

    file = generate_docs(dict(request.form), [
        "form1.docx",
        "form2.docx",
        "form9.docx",
        "form10.docx",
        "form11.docx"
    ])

    return send_file(file, as_attachment=True)

# ================= LOGOUT =================

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ================= RUN =================

if __name__ == "__main__":
    app.run(debug=True)
