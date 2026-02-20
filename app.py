from flask import Flask, render_template, request, redirect
import os
import pandas as pd
from datetime import datetime
from io import BytesIO
from azure.storage.blob import BlobServiceClient
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==============================
# Flask App Setup
# ==============================
app = Flask(__name__)

# ==============================
# Azure Blob Storage Setup
# ==============================
AZURE_CONNECTION_STRING = os.environ.get("AZURE_STORAGE_CONNECTION_STRING")
CONTAINER_NAME = os.environ.get("AZURE_CONTAINER_NAME", "labusers")

if not AZURE_CONNECTION_STRING:
    raise ValueError("AZURE_STORAGE_CONNECTION_STRING environment variable not set.")

blob_service_client = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)
container_client = blob_service_client.get_container_client(CONTAINER_NAME)

# Ensure container exists
try:
    container_client.create_container()
except Exception:
    pass

# ==============================
# Email Sending Function
# ==============================
def send_email(to_email, subject, body):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    gmail_user = os.environ.get("GMAIL_USER")
    gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD")

    if not gmail_user or not gmail_app_password:
        raise ValueError("GMAIL_USER or GMAIL_APP_PASSWORD not set")

    msg = MIMEMultipart()
    msg["From"] = gmail_user
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(gmail_user, gmail_app_password)
    server.send_message(msg)
    server.quit()

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
    email = request.form.get("email")
    department = request.form.get("department")
    degree = request.form.get("degree")
    supervisor = request.form.get("supervisor")
    equipment = request.form.get("equipment")
    purpose = request.form.get("purpose")
    file = request.files.get("id_card")

    if not name or not email:
        return "Name and email are required"

    # ---- Save Uploaded File to Blob ----
    if file and file.filename != "":
        filename = f"{name}_ID_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
        blob_path = f"UserFolder/{name}/{filename}"
        blob_client = container_client.get_blob_client(blob_path)
        blob_client.upload_blob(file.stream, overwrite=True)

    # ---- Prepare Record ----
    new_record = {
        "Name": name,
        "Email": email,
        "Department": department,
        "Degree": degree,
        "Supervisor": supervisor,
        "Equipment": equipment,
        "Purpose": purpose,
        "Submission Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # ---- Update Excel in Blob ----
    blob_name = "LabUserMaster.xlsx"
    blob_client = container_client.get_blob_client(blob_name)

    try:
        stream = BytesIO()
        blob_client.download_blob().readinto(stream)
        stream.seek(0)
        df = pd.read_excel(stream, engine="openpyxl")
        df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
    except Exception:
        df = pd.DataFrame([new_record])

    # Save updated Excel back to Blob
    output_stream = BytesIO()
    df.to_excel(output_stream, index=False, engine="openpyxl")
    output_stream.seek(0)
    blob_client.upload_blob(output_stream, overwrite=True)

    # ---- Send Emails ----
    # User Confirmation
    send_email(
        to_email=email,
        subject="Lab Registration Confirmation",
        body=f"Hello {name},\n\nThank you for registering. Your information has been recorded."
    )

    # Admin Notification
    admin_email = "wuyang.gao007@gmail.com"
    send_email(
        to_email=admin_email,
        subject=f"New Lab Registration: {name}",
        body=f"User {name} has registered.\nEmail: {email}\nDepartment: {department}\nDegree: {degree}\nSupervisor: {supervisor}\nEquipment: {equipment}\nPurpose: {purpose}"
    )

    return redirect("/")

# ==============================
# Run App
# ==============================
if __name__ == "__main__":
    app.run(debug=True)
