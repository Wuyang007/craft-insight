from flask import Flask, render_template, request, redirect
import os
import pandas as pd
from datetime import datetime

# ==============================
# Flask App Setup
# ==============================
app = Flask(__name__)

# ==============================
# Paths (AUTO-DETECTED)
# ==============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LAB_USERS_FOLDER = os.path.join(BASE_DIR, "LabUsers")
UPLOAD_FOLDER = os.path.join(LAB_USERS_FOLDER, "uploads")
MASTER_FILE = os.path.join(LAB_USERS_FOLDER, "LabUserMaster.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ==============================
# Home Route (Form Page)
# ==============================
@app.route("/")
def register():
    return render_template("form.html")


# ==============================
# Form Submission Route
# ==============================
@app.route("/submit", methods=["POST"])
def submit():

    # ---- Get Form Data ----
    name = request.form.get("name")
    department = request.form.get("department")
    degree = request.form.get("degree")
    supervisor = request.form.get("supervisor")
    equipment = request.form.get("equipment")
    purpose = request.form.get("purpose")

    file = request.files.get("id_card")

    if not name:
        return "Name is required"

    # ---- Create User Folder ----
    user_folder = os.path.join(LAB_USERS_FOLDER, name)
    os.makedirs(user_folder, exist_ok=True)

    # ---- Save Uploaded File ----
    if file and file.filename != "":
        filename = f"{name}_ID_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
        file_path = os.path.join(user_folder, filename)
        file.save(file_path)

    # ---- Prepare Record ----
    new_record = {
        "Name": name,
        "Department": department,
        "Degree": degree,
        "Supervisor": supervisor,
        "Equipment": equipment,
        "Purpose": purpose,
        "Submission Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # ---- Update Excel ----
    if os.path.exists(MASTER_FILE):
        df = pd.read_excel(MASTER_FILE, engine="openpyxl")
        df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
    else:
        df = pd.DataFrame([new_record])

    df.to_excel(MASTER_FILE, index=False, engine="openpyxl")

    return redirect("/")


# ==============================
# Run App
# ==============================
if __name__ == "__main__":
    app.run(debug=True)
