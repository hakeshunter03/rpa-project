# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = single_input("Please input the directory where the ProductData.xlsx, MasterInvoice.pdf, and Order files are located.")

    if directory[-1] != "/":
        directory = directory + "/"

    filenames= os.listdir(directory)

    if ('ProductData.xlsx' in filenames and
        'MasterInvoice.pdf' in filenames and
        'Order1.xlsx' in filenames and
        'Order2.xlsx' in filenames and
        'Order3.xlsx' in filenames and
        'Order4.xlsx' in filenames):
        break
    else:
        print("Missing files. Please enter the directory again.")
    
      
      # INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
descriptions_df = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Descriptions")

for index,row in descriptions_df.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description


product_prices = {}
prices_df = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Prices")

for index,row in prices_df.iterrows():
    unit_number = row['Unit number']
    price = row['Price per unit']
    product_prices[unit_number] = price


print(product_descriptions)
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory+filename, data_only = True)
    sheet = template['OrderForm']

    name = str(sheet['C14'].value)
    email = str(sheet['C13'].value)
    date = sheet['C9'].value.strftime('%m/%d/%Y')

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = str(sheet['C15'].value)
    order_data['CustomerAddress1'] = str(sheet['C16'].value)
    order_data['CustomerAddress2'] = f"{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}"

    total = 0
    row = 24
    i = 0
    while sheet[f'B{row}'].value is not None:
        unit = sheet[f'B{row}'].value
        qty = sheet[f'C{row}'].value
        price = product_prices[unit]
        amount = price * qty
        total = total + amount

        order_data[f'Description{i}'] = str(product_descriptions[unit])
        order_data[f'Quantity{i}'] = str(qty)
        order_data[f'Price{i}'] = f'${price:,.2f}'
        order_data[f'Amount{i}'] = f'${amount:,.2f}'

        row = row + 1
        i = i + 1

    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date

#order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)


#create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    gmail_user = "hakeshunter03@gmail.com"
  
    # List of recipients
    recipients = [email]

    # Email subject
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email
    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment
    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    # Log in to the Gmail server
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email
    mail_server.send_message(message)
    mail_server.quit()


# send_email(8515, 'Thomas Holt', 'hakeshunter03@gmail.com', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = sorted(os.listdir(directory))
invoice_num = 8515
for file in filenames:
    if "Order" in file and "xlsx" in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num,order_data, directory)
        send_email(invoice_num,name, email, date, directory)
        invoice_num = invoice_num +1
