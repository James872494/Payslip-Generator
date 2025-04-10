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


def generate_payslip(row):
    net_salary = row["BASIC SALARY"] + row["ALLOWENCES"] - row["DEDUCTIONS"]
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # 🔷 Header Section
    pdf.set_fill_color(0, 102, 204)  # Blue background
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 15, "  Employee Payslip", ln=True, fill=True)

    # 🧾 Employee Details
    pdf.ln(10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Employee Information", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 8, f"Employee ID: {row['EMPLOYEE ID']}", ln=True)
    pdf.cell(0, 8, f"Name: {row['NAME']}", ln=True)
    pdf.cell(0, 8, f"Email: {row['EMAIL']}", ln=True)

    # 💰 Salary Breakdown
    pdf.ln(10)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Salary Breakdown", ln=True)
    pdf.set_font("Arial", "", 12)

    def draw_row(label, amount, fill=False):
        pdf.set_fill_color(245, 245, 245)  # Light gray
        pdf.cell(90, 10, label, border=1, fill=fill)
        pdf.cell(90, 10, f"${amount:.2f}", border=1, ln=True, fill=fill)

    draw_row("Basic Salary", row["BASIC SALARY"], fill=True)
    draw_row("Allowances", row["ALLOWENCES"])
    draw_row("Deductions", row["DEDUCTIONS"], fill=True)

    # 🧮 Net Salary
    pdf.set_font("Arial", "B", 12)
    pdf.ln(5)
    pdf.set_fill_color(204, 255, 204)  # Green fill
    pdf.cell(90, 10, "Net Salary", border=1, fill=True)
    pdf.cell(90, 10, f"${net_salary:.2f}", border=1, ln=True, fill=True)

    # 📝 Footer
    pdf.ln(15)
    pdf.set_font("Arial", "I", 10)
    pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(0, 8, "Note: This payslip is system-generated and does not require a signature.\nIf you have any questions, please contact HR.")
    pdf.set_text_color(0, 0, 0)

    # 💾 Save PDF
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

