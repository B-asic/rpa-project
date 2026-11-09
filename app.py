
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

#INSERT ANSWER FOR QUESTION 1 HERE

required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx','Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input('Enter directory for project files')
    directory = directory.strip()

    if not directory.endswith('\\'):
        directory = directory + '\\'
    try:
        filenames= os.listdir(directory) #gets the names of files in the folder
    except Exception as e:
        filenames = [] #empty list if missing files
    missed_files =[]

    for filename in required_files:
        if filename not in filenames:
            missed_files.append(filename)# adds missing file to the previously made list
    if len(missed_files) == 0:
        break
    else: show_message('Error', 'Folder is incorrect or files are missing, please try again.' )

#C:\Users\likin\OneDrive\Desktop\Fall 2026 Course Planner - 2026-09-03\Fall_2026_All_Downloads\Fall_2026_Courses\ACC_525\Advanced_Analytics\17 - Project 3 - Robotic Process Automation# Insert your python code here
# INSERT ANSWER FOR QUESTION 2 HERE
descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')

product_descriptions= {} #dictionary for products lol

for index, row in descriptions.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description
    
prices = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')

product_prices= {} #dictionary for prices lol

for index, row in prices.iterrows():
    unit_number = row['Unit number']
    price = row['Price per unit']
    product_prices[unit_number] = price
    
# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename,directory):

    workbook = openpyxl.load_workbook(directory + filename,data_only= True)
    selected_sheet = workbook['OrderForm']
    
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    company = selected_sheet['C15'].value
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value

    date = selected_sheet['C9'].value
    date = date.strftime('%m/%d/%Y')

    order_data = {}

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = str(name)
    order_data['CustomerCompany'] = str(company)
    order_data['CustomerAddress1'] = str(address)
    order_data['CustomerAddress2'] = f'{city},{state} {zipcode}'

    total = 0 #starts order total

    for index, row_number in enumerate([24,25,26]):

        unit_number = selected_sheet[f'B{row_number}'].value

        if unit_number is not None:
            quantity = selected_sheet[f'C{row_number}'].value
            description = product_descriptions[unit_number]
            price = product_prices[unit_number]

            amount = price * quantity
            total = total + amount

            order_data[f'Description{index}'] = str(description)
            order_data[f'Price{index}']= f'${price:,.2f}'
            order_data[f'Quantity{index}']= str(quantity)
            order_data[f'Amount{index}']= f'${amount:,.2f}'
    order_data['Total'] = f'${total:,.2f}'

    workbook.close()

    return order_data, name ,email, date

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num,order_data, directory):
    template = directory + 'MasterInvoice.pdf'
    invoice_file = directory + f'Invoice {invoice_num}.pdf'

    create_pdf(template,invoice_file, order_data)
 #it worked but I have no idea what the font error is about 

# INSERT ANSWER FOR QUESTION 5 HERE
gmail_user = "jay.mvasquez7@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email,date, directory):

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email


    body = f"""Dear {name},
    
Attached is your invoice dated {date}

Thank you for your business.

PythonTech Co
10 Python Drive
Pleasant Grove, UT 84062
Phone: 801-555-2372
Email: sales@pythontechco.com

"""

# Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

# Add attachment

    path_to_file = directory+ f'Invoice {invoice_num}.pdf'


    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package

        
# Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

# Send the email

    mail_server.send_message(message)
    mail_server.quit()

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory) #filenames from the folder
filenames.sort()
invoice_num = 8515 #beg. invoice number

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        #get order info
        order_data,name,email, date = get_order_data(invoice_num, filename, directory)
        #creates invoice
        create_invoice(invoice_num,order_data, directory)
        #sends email (wicked coolbtw)
        send_email(invoice_num,name,email,date,directory)

        print(f'Invoice {invoice_num} sent to {email}')

        invoice_num =invoice_num + 1



    
    
