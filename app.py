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
pd.set_option('display.max_rows', None) # Display option to show all rows# Insert your python code here

# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = single_input('Please provide your directory:')
    if directory[-1] != ('\\'):
        directory = directory + '\\'
    try:
        filenames = os.listdir(directory)
    except:
        filenames = []
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        show_message('Success', 'Directory added.')
        break
    else:
        show_message('Error', 'Directory does not contain required files.')

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index,row in descriptions_df.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description

prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for index,row in prices_df.iterrows():
    unit_number2 = row['Unit number']
    price = row['Price per unit']
    product_prices[unit_number2] = price
  
# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):

    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    date = selected_sheet['C9'].value
    date = date.strftime('%m/%d/%y')
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value

    order_data = {}

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'

    total = 0

    if selected_sheet['B24'].value is not None:
        order_data['Price0'] = f'${product_prices[selected_sheet['B24'].value]:,.2f}'
        order_data['Quantity0'] = selected_sheet['C24'].value
        amount = product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value
        order_data['Amount0'] = f'${amount:,.2f}'
        order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]

        total = total + amount

    if selected_sheet['B25'].value is not None:
        order_data['Price1'] = f'${product_prices[selected_sheet['B25'].value]:,.2f}'
        order_data['Quantity1'] = selected_sheet['C25'].value
        amount = product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value
        order_data['Amount1'] = f'${amount:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]

        total = total + amount

    if selected_sheet['B26'].value is not None:
        order_data['Price2'] = f'${product_prices[selected_sheet['B26'].value]:,.2f}'
        order_data['Quantity2'] = selected_sheet['C26'].value
        amount = product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value
        order_data['Amount2'] = f'${amount:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]

        total = total + amount

    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date


order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    gmail_user = "grantmalakai22@gmail.com"

    # List of recipients

    recipients = email

    # Email subect

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
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', email, '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
invoice_num = 8515

filenames = os.listdir(directory)

for file in filenames:
    try:
        if "Order" and "xlsx" in file:
            dictionary, new_name, new_email, new_date = get_order_data(invoice_num, file, directory)
            create_invoice(invoice_num, dictionary, directory)
            send_email(invoice_num, new_name, new_email, new_date, directory)
            invoice_num = invoice_num + 1
    except:
        print("Not an order workbook.")
