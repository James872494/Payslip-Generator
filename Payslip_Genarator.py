'''from dotenv import load_dotenv
import os

load_dotenv()

print("SMTP Server:", os.getenv("SMTP_SERVER"))
print("SMTP Port:", os.getenv("SMTP_PORT"))
print("Sender Email:", os.getenv("SENDER_EMAIL"))
print("Sender Password:", os.getenv("SENDER_PASSWORD"))  # 🔒 Only for testing'''

import os
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from dotenv import load_dotenv

# 🚀 Load environment variables (SMTP config)
load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
GMAIL_USER = os.getenv("GMAIL_USER", "jamesmutembedza87@gmail.com")  # ✅ Your Gmail
GMAIL_PASS = os.getenv("GMAIL_PASS", "pyllmkgxvffvobba")              # ✅ App password

# 📁 Create output folder for payslips
output_dir = "payslips"
os.makedirs(output_dir, exist_ok=True)

# 📄 Read Excel file
try:
    df = pd.read_excel("employees.xlsx")
except FileNotFoundError:
    print("❌ Error: employees.xlsx not found.")
    exit(1)

# 🧾 Generate payslip PDF
def generate_payslip(row):
    net_salary = row["BASIC SALARY"] + row["ALLOWENCES"] - row["DEDUCTIONS"]
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Header
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt="Employee Payslip", ln=True, align='C')
    pdf.ln(10)

    # Info
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Employee ID: {row['EMPLOYEE ID']}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {row['NAME']}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${row['BASIC SALARY']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${row['ALLOWENCES']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${row['DEDUCTIONS']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${net_salary:.2f}", ln=True)

    filename = os.path.join(output_dir, f"{row['EMPLOYEE ID']}.pdf")
    pdf.output(filename)
    return filename

# ✉️ Send email with attachment
def send_email(to_email, name, attachment_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = GMAIL_USER
        msg['To'] = to_email
        msg['Subject'] = "Your Payslip for This Month"

        body = f"Dear {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Department"
        msg.attach(MIMEText(body, 'plain'))

        with open(attachment_path, "rb") as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(attachment_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            msg.attach(part)

        # ✅ Send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.set_debuglevel(1)
            server.starttls()
            server.login(GMAIL_USER, GMAIL_PASS)
            server.send_message(msg)

        print(f"✅ Payslip sent to {name} at {to_email}")

    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")

# 🔁 Main loop
for index, row in df.iterrows():
    if pd.isna(row["EMAIL"]):
        print(f"⚠️ Skipping {row['NAME']} due to missing email.")
        continue

    try:
        payslip_path = generate_payslip(row)
        send_email(row["EMAIL"], row["NAME"], payslip_path)
    except Exception as e:
        print(f"❌ Error processing employee ID {row['EMPLOYEE ID']}: {e}")

print("🏁 All payslips processed.")

