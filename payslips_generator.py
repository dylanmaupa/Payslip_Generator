import pandas as pd
from fpdf import FPDF
import yagmail

df = pd.read_excel("employees.xlsx")  # Make sure the file is in the same directory
df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']



def generate_payslip(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    pdf.cell(200, 10, txt="Employee Payslip", ln=True, align="C")
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Employee ID: {row['Employee ID']}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {row['Name']}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${row['Basic Salary']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${row['Allowances']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${row['Deductions']:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${row['Net Salary']:.2f}", ln=True)

    filename = f"payslips/{row['Name']}.pdf"
    pdf.output(filename)
    return filename

# Create a directory for payslips if it doesn't exist
import os

os.makedirs("payslips", exist_ok=True)

# Generate payslips and send emails

yag = yagmail.SMTP(user="Email", password="password")

def send_email(row, pdf_path):
    subject = "Your Payslip for This Month"
    body = f"Dear {row['Name']},\\n\\nPlease find your payslip attached.\\n\\nRegards, HR"
    
    yag.send(
        to=row["Email"],
        subject=subject,
        contents=body,
        attachments=pdf_path
    )

for _, row in df.iterrows():
    try:
        pdf_path = generate_payslip(row)
        send_email(row, pdf_path)
        print(f"Payslip sent to {row['Name']} at {row['Email']}")
    except Exception as e:
        print(f"Error for {row['Name']}: {e}")

