import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

# Ensure payslips folder exists
os.makedirs("payslips", exist_ok=True)

# Generate a payslip PDF for a single employee
def generate_payslip(employee):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, "Monthly Payslip", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, f"Employee ID: {employee['Employee ID']}", ln=True)
    pdf.cell(200, 10, f"Name: {employee['Name']}", ln=True)
    pdf.ln(5)
    
    pdf.cell(200, 10, f"Basic Salary: ${employee['Basic Salary']:.2f}", ln=True)
    pdf.cell(200, 10, f"Allowances: ${employee['Allowances']:.2f}", ln=True)
    pdf.cell(200, 10, f"Deductions: ${employee['Deductions']:.2f}", ln=True)
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(200, 10, f"Net Salary: ${employee['Net Salary']:.2f}", ln=True)

    filename = f"payslips/{employee['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

# Send email with the payslip attachment
def send_email(to_email, attachment_path):
    try:
        yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASS)
        subject = "Your Payslip for This Month"
        body = "Dear Employee,\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Team"
        yag.send(to=to_email, subject=subject, contents=body, attachments=attachment_path)
        print(f"[✓] Email sent to {to_email}")
    except Exception as e:
        print(f"[✗] Failed to send email to {to_email}: {e}")

# Main function
def main():
    try:
        df = pd.read_excel("employees.xlsx", engine='openpyxl')
        required_cols = ['Employee ID', 'Name', 'Email', 'Basic Salary', 'Allowances', 'Deductions']
        
        if not all(col in df.columns for col in required_cols):
            raise ValueError("Excel file is missing required columns.")
        
        for _, row in df.iterrows():
            try:
                row['Net Salary'] = row['Basic Salary'] + row['Allowances'] - row['Deductions']
                payslip_path = generate_payslip(row)
              #  send_email(row['Email'], payslip_path)
            except Exception as emp_err:
                print(f"[!] Error processing employee {row['Name']}: {emp_err}")
    
    except FileNotFoundError:
        print("[!] employees.xlsx not found. Please make sure it is in the same folder.")
    except Exception as e:
        print(f"[!] An error occurred: {e}")

if __name__ == "__main__":
    main()
