# RESE
read excel and send a email
mport necessary modules
import openpyxl
from email.mime.text import MIMEText
import smtplib

Set the file path for the excel file
file_path = "excel_file.xlsx"

Open the workbook and select the active sheet
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

Create an empty list to store the data from the excel file
data = []

Iterate through the rows and columns of the sheet
for row in sheet.iter_rows(values_only=True):
data.append(row)

Create the email letter
letter = "Dear [Recipient],\n\n"
letter += "I am writing to inform you of the following information:\n\n"
letter += "Slot 1: {}\n".format(data[0][0])
letter += "Slot 2: {}\n".format(data[0][1])
letter += "Slot 3: {}\n\n".format(data[0][2])
letter += "Thank you for your attention to this matter.\n\n"
letter += "Sincerely,\n[Your Name]"

Convert the letter to an email message
msg = MIMEText(letter)
msg['Subject'] = 'Information from Excel File'
msg['From'] = '[Your Email Address]'
msg['To'] = '[Recipient Email Address]'

Set up the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login('[Your Email Address]', '[Your Email Password]')

Send the email
server.send_message(msg)

Close the SMTP server
server.quit()

print("Email sent successfully.")
