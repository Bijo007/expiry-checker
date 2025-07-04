import openpyxl
import smtplib
from email.message import EmailMessage
from datetime import datetime

# Load your Excel file (replace with your actual file path)
wb = openpyxl.load_workbook('expiry_dates.xlsx')
sheet = wb.active

today = datetime.today().date()

expiring_soon = []

for row in sheet.iter_rows(min_row=2, values_only=True):  # assuming first row headers
    item, expiry = row
    if isinstance(expiry, datetime):
        expiry_date = expiry.date()
        days_left = (expiry_date - today).days
        if days_left <= 30:
            expiring_soon.append(f"{item} expires in {days_left} days")

if expiring_soon:
    msg = EmailMessage()
    msg['Subject'] = 'Expiry Alert'
    msg['From'] = 'your-email@example.com'
    msg['To'] = 'recipient@example.com'
    msg.set_content('\n'.join(expiring_soon))

    # Send email (fill in your SMTP details)
    with smtplib.SMTP('smtp.example.com', 587) as smtp:
        smtp.starttls()
        smtp.login('your-email@example.com', 'your-email-password')
        smtp.send_message(msg)

    print("Expiry alert sent.")
else:
    print("No items expiring soon.")
