# Insert your python code here
while True:
    directory = single_input("Please input the directory where the ProductData.xlsx, MasterInvoice.pdf, and Order files are located.")
    filenames = os.listdir(directory)

    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        if directory[-1] == '\\':
            break
        else:
            directory = directory + '\\'
            break
    else:
        show_message('Error', 'Please enter the correct files')

product_prices = {}
product_descriptions = {}

df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name= 'Descriptions')
for index,row in df.iterrows():
    
    template = openpyxl.load_workbook('ProductData.xlsx')
    selected_sheet = template['Descriptions']
    product_descriptions[row['Unit number']] = row['Description']

    
df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name= 'Prices')
for index,row in df1.iterrows():
    template = openpyxl.load_workbook('ProductData.xlsx')
    selected_sheet = template['Prices']
    product_prices[row['Unit number']] = row['Price per unit']


def get_order_data(invoice_num, filename, directory):
    order_data = {}
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    row = 24
    count = 0
    total = 0
    while selected_sheet[f'B{row}'].value is not None:
        order_data[f'Price{count}'] = f'${selected_sheet[f'C{row}'].value}'
        order_data[f'Quantity{count}'] = selected_sheet[f'B{row}'].value
        order_data[f'Amount{count}'] = f'${selected_sheet[f'C{row}'].value * product_prices[selected_sheet[f'B{row}'].value]:,.2f}'
        order_data[f'Description{count}'] = product_descriptions[selected_sheet[f'B{row}'].value]
        total += selected_sheet[f'C{row}'].value * product_prices[selected_sheet[f'B{row}'].value]
        row += 1
        count += 1

    order_data['Total'] = total
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress1'] = (str(selected_sheet['C16'].value))
    order_data['CustomerAddress2'] = (f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}')
    order_data['InvoiceNum'] = invoice_num
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['CustomerName'] = name
    order_data['Email'] = email
    order_data['Date'] = date
    return [order_data, name, email, date]


def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)


def send_email(invoice_num, name, email, date, directory):
    gmail_user = "Samuel.keehan@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = name

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

send_email(8515, 'Thomas Holt', 'samuel.keehan@gmail.com', '11/13/2023', directory)

filenames = os.listdir(directory)
invoice_num = 8515

for filename in filenames:
    
    if 'Order' in filename and 'xlsx' in filename:
        print(invoice_num)
        get_order_data(invoice_num,filename,directory)
        order_data = get_order_data(invoice_num,filename,directory)[0]
        name = get_order_data(invoice_num, filename, directory)[1]
        email = get_order_data(invoice_num, filename, directory) [2]
        date = get_order_data(invoice_num, filename, directory)[3]
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name, email, date, directory)
        invoice_num += 1

        