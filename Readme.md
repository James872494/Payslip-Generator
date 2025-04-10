# 📄 Payslip Generator

The Payslip Generator is a Python tool designed to streamline the process of creating and distributing monthly payslips to employees. It reads employee data from an Excel file, computes net salaries, generates PDF payslips, and sends them via email. Below is a comprehensive guide to setting up and using the tool.

# 📌 Key Features
-Import Employee Records: Easily load employee data from an Excel file using the `pandas` library.
-Salary Calculation: Automatically compute each employee’s net salary after deductions.
-PDF Payslip Creation: Utilize fpdf to generate structured PDF documents for payslips.
-Automated Emailing: Send the generated payslips to employees directly via email using smtplib.
-Secure Email Management: Handle email credentials securely using python-dotenv.

# 🛠 Requirements
-Ensure you have the following installed:

✔Python 3.x
✔Required libraries:
✔pandas
✔fpdf
✔python-dotenv
✔Install Dependencies

-You can install the required packages using pip. Run the following command in your terminal:

pip install `pandas`, `fpdf`, `python-dotenv`

