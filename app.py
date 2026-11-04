while True:
    directory = single_input('Please enter the directory where the Product Data, Master invoice, and Orders are located, ending with \\')
    directory = directory.strip('"')
    if directory[-1:] != "\\":
        directory = directory + "\\"
    try:
        files = os.listdir(directory)
        if 'ProductData.xlsx' in files and 'MasterInvoice.pdf' in files and 'Order1.xlsx' in files and 'Order2.xlsx' in files and 'Order3.xlsx' in files and 'Order4.xlsx' in files:
            show_message('Success','This is the correct filepath, saved.')
            break
        else:
            show_message('Error--wrong directory','Please enter the directory for the invoices, orders, and product data.')
    except:
        show_message('Error--not a filepath','Please enter a valid filepath.')
product_descriptions = {}
descriptions = pd.read_excel(directory + 'ProductData.xlsx','Descriptions')
for i, row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices = pd.read_excel(directory + 'ProductData.xlsx','Prices')
prices.head()
product_prices = {}
for i, row in prices.iterrows():
    product_prices[row['Unit number']] = row["Price per unit"]
def get_order_data(invoice_num,filename,directory):
    working_file = openpyxl.load_workbook(directory+filename)
    selected_sheet = working_file['OrderForm']
    order_data = {}
    form = get_pdf('MasterInvoice.pdf')
    for i, row in form.iterrows():
        order_data[row['name']] = ""

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['CustomerAddress1'] = str(selected_sheet['C16'].value)
    order_data['CustomerName'] = str(selected_sheet['C14'].value)
    order_data['CustomerCompany'] = str(selected_sheet['C15'].value)
    order_data['CustomerAddress2'] = str(selected_sheet['C17'].value) + ", " + str(selected_sheet['C18'].value) + " " + str(selected_sheet['C19'].value)
    total_count = 0
    while True:
        if pd.isnull([selected_sheet['B24'].value])== False:
            order_data['Description0'] =  product_descriptions[selected_sheet['B24'].value]
            order_data['Price0'] = product_prices[selected_sheet['B24'].value]
            order_data['Quantity0'] = selected_sheet['C24'].value
            order_data['Amount0'] = order_data['Price0'] * order_data['Quantity0']
            total_count = total_count + float(order_data['Amount0'])
            order_data['Amount0'] = f'${(order_data['Price0'] * order_data['Quantity0']):,.2f}'

        else:
            break
        if pd.isnull([selected_sheet['B25'].value])== False:
            order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
            order_data['Price1'] = product_prices[selected_sheet['B25'].value]
            order_data['Quantity1'] = selected_sheet['C25'].value
            order_data['Amount1'] = order_data['Price1'] * order_data['Quantity1']
            total_count = total_count + float(order_data['Amount1'])
            order_data['Amount1'] = f'${(order_data['Price1'] * order_data['Quantity1']):,.2f}'
        else:
            break
        if pd.isnull([selected_sheet['B26'].value])== False:
            order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]
            order_data['Price2'] = product_prices[selected_sheet['B26'].value]
            order_data['Quantity2'] = selected_sheet['C26'].value
            order_data['Amount2'] = order_data['Price2'] * order_data['Quantity2']
            total_count = total_count + float(order_data['Amount2'])
            order_data['Amount2'] = f'${(order_data['Price2'] * order_data['Quantity2']):,.2f}'
        else:
            break
        if pd.isnull([selected_sheet['B27'].value])== False:
            order_data['Description3'] = product_descriptions[selected_sheet['B27'].value]
            order_data['Price3'] = product_prices[selected_sheet['B27'].value]
            order_data['Quantity3'] = selected_sheet['C27'].value
            order_data['Amount3'] = order_data['Price3'] * order_data['Quantity3']
            total_count = total_count + float(order_data['Amount3'])
            order_data['Amount3'] = f'${(order_data['Price3'] * order_data['Quantity3']):,.2f}'
        else:
            break
        if pd.isnull([selected_sheet['B28'].value])== False:
            order_data['Description4'] = product_descriptions[selected_sheet['B28'].value]
            order_data['Price4'] = product_prices[selected_sheet['B28'].value]
            order_data['Quantity4'] = selected_sheet['C28'].value
            order_data['Amount4'] = order_data['Price4'] * order_data['Quantity4']
            total_count = total_count + float(order_data['Amount4'])
            order_data['Amount4'] = f'${(order_data['Price4'] * order_data['Quantity4']):,.2f}'
            break
        else:
            break
    order_data['Total'] = f'${total_count:,.2f}'
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = str(selected_sheet['C9'].value).split(" ")[0].split('-')[1] + "/" + str(selected_sheet['C9'].value).split(" ")[0].split('-')[2] + "/" + str(selected_sheet['C9'].value).split(" ")[0].split('-')[0]
    order_data['Date'] = date
    return order_data,name,email,date
def create_invoice(invoice_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf',directory+f'Invoice {invoice_num}.pdf',order_data)
order_dict = get_order_data(8515,'Order1.xlsx',directory)[0]
create_invoice(8515,order_dict,directory)
def send_email(invoice_num,name,email,date,directory):
    gmail_user = "benlpettit@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()


    recipients = email


    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'


    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    E-mail: sales@pythontechco.com"""
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)
    path_to_file = directory + f"\\Invoice {invoice_num}.pdf"
    message = add_attachment(f"Invoice {invoice_num}.pdf", message) # This function is available from the analyticskit package
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()
name = get_order_data(8515,'Order1.xlsx',directory)[1]
email = get_order_data(8515,'Order1.xlsx',directory)[2]
date = get_order_data(8515,'Order1.xlsx',directory)[3]
send_email(8515,name,email,date,directory)
filenames = os.listdir(directory)
for f, filename in enumerate(filenames):
    invoicenum = 8515
    if 'Order' in filename and 'xlsx' in filename:
        create_invoice(invoicenum,get_order_data(invoicenum,filename,directory)[0],directory)
        sentorderdata = get_order_data(invoicenum,filename,directory)[0]
        create_invoice(invoicenum,sentorderdata,directory)
        sentname = get_order_data(invoicenum,filename,directory)[1]
        sentemail = get_order_data(invoicenum,filename,directory)[2]
        sentdate = get_order_data(invoicenum,filename,directory)[3]
        send_email(invoicenum,sentname,sentemail,sentdate,directory)
        invoicenum +=1