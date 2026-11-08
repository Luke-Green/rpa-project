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


directory = single_input("Please enter the directory.", mask=False, width=400, height=200)+'\\'
print(os.getcwd())
files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']
while True:
    if files[0] in os.listdir(directory) and files[1] in os.listdir(directory) and files[2] in os.listdir(directory) and files[3] in os.listdir(directory) and files[4] in os.listdir(directory) and files[5] in os.listdir(directory):
        break
    else:
        # print(os.listdir(directory))
        directory = single_input("Directory does not contain the required files, please try again.", mask=False, width=400, height=200)+'\\'

  df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}

for r, row in df1.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df2 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices = {}

for r, row in df2.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
print(product_descriptions)
print(product_prices)

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename)
    sheet = template['OrderForm']
    cust_name = sheet['C14'].value
    cust_email = sheet['C13'].value
    date = sheet['C9'].value
    date = date.strftime('%m/%d/%Y')
    unit0 = sheet['B24'].value
    unit1 = sheet['B25'].value
    unit2 = sheet['B26'].value
    unit3 = sheet['B27'].value
    unit4 = sheet['B28'].value
    print(unit0)
    print(product_prices[unit0])
    price0 = int(product_prices.get(unit0,0))
    price1 = int(product_prices.get(unit1,0))
    price2 = int(product_prices.get(unit2,0))
    price3 = int(product_prices.get(unit3,0))
    price4 = int(product_prices.get(unit4,0))
    quantity0 = sheet['C24'].value
    if quantity0 is None:
        quantity0 = 0
    quantity0 = int(quantity0)
    quantity1 = sheet['C25'].value
    if quantity1 is None:
            quantity1 = 0
    quantity1 = int(quantity1)
    quantity2 = sheet['C26'].value
    if quantity2 is None:
            quantity2 = 0
    quantity2 = int(quantity2)
    quantity3 = sheet['C27'].value
    if quantity3 is None:
            quantity3 = 0
    quantity3 = int(quantity3)
    quantity4 = sheet['C28'].value
    if quantity4 is None:
            quantity4 = 0
    quantity4 = int(quantity4)
    update_dict = {}
    update_dict['InvoiceNum'] = f'{invoice_num}'
    update_dict['Date'] = f'{date}'
    update_dict['Total'] = f'{price0*quantity0+price1*quantity1+price2*quantity2+price3*quantity3+price4*quantity4}'
    update_dict['CustomerAddress1'] = f'{sheet['C16'].value}'
    update_dict['CustomerName'] = f'{cust_name}'
    update_dict['CustomerCompany'] = f'{sheet['C15'].value}'
    update_dict['CustomerAddress2'] = ''
    update_dict['Price0'] = f'{price0}'
    update_dict['Quantity0'] = f'{quantity0}'
    update_dict['Amount0'] = f'{price0*quantity0}'
    update_dict['Price1'] = f'{price1}'
    update_dict['Quantity1'] = f'{quantity1}'
    update_dict['Amount1'] = f'{price1*quantity1}'
    update_dict['Price2'] = f'{price2}'
    update_dict['Quantity2'] = f'{quantity2}'
    update_dict['Amount2'] = f'{price2*quantity2}'
    update_dict['Price3'] = f'{price3}'
    update_dict['Quantity3'] = f'{quantity3}'
    update_dict['Amount3'] = f'{price3*quantity3}'
    update_dict['Price4'] = f'{price4}'
    update_dict['Quantity4'] = f'{quantity4}'
    update_dict['Amount4'] = f'{price4*quantity4}'
    update_dict['Description0'] = f'{unit0}'
    update_dict['Description1'] = f'{unit1}'
    update_dict['Description2'] = f'{unit2}'
    update_dict['Description3'] = f'{unit3}'
    update_dict['Description4'] = f'{unit4}'
    return update_dict, cust_name, cust_email, date

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
        
    # Enter your gmail user name and password

    gmail_user = "lukegreen225@gmail.com"
    

    # List of recipients

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is the your invoice dated {date}.

    Thank you for your business.   
    
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com """

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    
    message = add_attachment(directory + f'Invoice {invoice_num}.pdf', message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:

        order_data, name, email, date = get_order_data(invoice_num, filename, directory)

        create_invoice(invoice_num, order_data, directory)

        send_email(invoice_num, name, email, date, directory)
        
        invoice_num = invoice_num + 1
