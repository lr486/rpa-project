#QUESTION 1
while True:
    directory = single_input("Please enter the filepath to the directory", mask=False, width=400, height=200)
    directory = directory + '/'
    directory_files = os.listdir(directory)
    if 'ProductData.xlsx' not in directory_files:
        show_message("Error", "The directory you entered is invalid. Please try again.")
    elif 'Order1.xlsx' not in directory_files:
        show_message("Error", "The directory you entered is invalid. Please try again.")
    elif 'Order2.xlsx' not in directory_files:
        show_message("Error", "The directory you entered is invalid. Please try again.")
    elif 'Order3.xlsx' not in directory_files:
        show_message("Error", "The directory you entered is invalid. Please try again.")
    elif 'Order4.xlsx' not in directory_files:
        show_message("Error", "The directory you entered is invalid. Please try again.")
    elif 'MasterInvoice.pdf' not in directory_files:
        show_message("Error", "The directory you entered is invalid. Please try again.")
    else:
          break


#QUESTION 2
product_prices = {}
p_prices = pd.read_excel('ProductData.xlsx', sheet_name='Prices')
for i, row in p_prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

product_descriptions = {}
p_descriptions = pd.read_excel('ProductData.xlsx', sheet_name='Descriptions')
for i, row in p_descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']# Insert your python code here



#QUESTION 3:
def get_order_data(invoice_num, filename, directory):
    file = filename
    order_excel = openpyxl.load_workbook(file, data_only=True)
    
    selected_sheet = order_excel['OrderForm']
    name = selected_sheet['C14'].value
    raw_date = str(selected_sheet['C9'].value)

    month = str(raw_date).split('-')[0]
    day = str(raw_date).split('-')[1]
    short_year = str(raw_date).split('-')[2]
    long_year = '20'+ short_year
    date = month + '/' + day + '/' + long_year
    email = selected_sheet['C13'].value


    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    
    order_data['Date'] = selected_sheet['C14'].value
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    # order_data['CustomerAddress2'] = selected_sheet[''].value



    if selected_sheet['C25'].value is None and selected_sheet['C26'].value is None:
        order_data['Quantity0'] = str(selected_sheet['C24'].value)
        order_data['Quantity1'] = ''
        order_data['Quantity2'] = ''
        quant1 = int(selected_sheet['C24'].value)


        unit_num1 = selected_sheet['B24'].value

      
        order_data['Description0'] = str(product_descriptions[f'{unit_num1}'])
        order_data['Description1'] = ''
        order_data['Description2'] = ''
        price1 = product_prices[f'{unit_num1}']


        order_data['Price0'] = str(f'${price1:,.2f}')
        order_data['Price1'] = ''
        order_data['Price2'] = ''
        amount1 = price1 * quant1

 
        order_data['Amount0'] = str(f'${amount1:,.2f}')
        order_data['Amount1'] = ''
        order_data['Amount2'] = ''

        total = (price1 * quant1)
        order_data['Total'] = f'${total:,.2f}'


    elif selected_sheet['C26'].value is None:
        order_data['Quantity0'] = str(selected_sheet['C24'].value)
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        order_data['Quantity2'] = ''
        quant1 = int(selected_sheet['C24'].value)
        quant2 = int(selected_sheet['C25'].value)

        unit_num1 = selected_sheet['B24'].value
        unit_num2 = selected_sheet['B25'].value
      
        order_data['Description0'] = str(product_descriptions[f'{unit_num1}'])
        order_data['Description1'] = str(product_descriptions[f'{unit_num2}'])
        order_data['Description2'] = ''
        price1 = product_prices[f'{unit_num1}']
        price2 = product_prices[f'{unit_num2}']

        order_data['Price0'] = f'${price1:,.2f}'
        order_data['Price1'] = f'${price2:,.2f}'
        order_data['Price2'] = ''
        amount1 = price1 * quant1
        amount2 = price2 * quant2
 
        order_data['Amount0'] = f'${amount1:,.2f}'
        order_data['Amount1'] = f'${amount2:,.2f}'
        order_data['Amount2'] = ''

        total = (price1 * quant1) + (price2 * quant2)
        order_data['Total'] = f'${total:,.2f}'
    elif selected_sheet['C26'].value is not None:
        order_data['Quantity0'] = str(selected_sheet['C24'].value)
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        order_data['Quantity2'] = str(selected_sheet['C25'].value)
        quant1 = int(selected_sheet['C24'].value)
        quant2 = int(selected_sheet['C25'].value)
        quant3 = int(selected_sheet['C26'].value)
        unit_num1 = selected_sheet['B24'].value
        unit_num2 = selected_sheet['B25'].value
        unit_num3 = selected_sheet['B26'].value
        order_data['Description0'] = str(product_descriptions[f'{unit_num1}'])
        order_data['Description1'] = str(product_descriptions[f'{unit_num2}'])
        order_data['Description2'] = str(product_descriptions[f'{unit_num3}'])
        price1 = product_prices[f'{unit_num1}']
        price2 = product_prices[f'{unit_num2}']
        price3 = product_prices[f'{unit_num3}']
        order_data['Price0'] = f'${price1:,.2f}'
        order_data['Price1'] = f'${price2:,.2f}'
        order_data['Price2'] = f'${price3:,.2f}'
        amount1 = price1 * quant1
        amount2 = price2 * quant2
        amount3 = price3 * quant3
        order_data['Amount0'] = f'${amount1:,.2f}'
        order_data['Amount1'] = f'${amount2:,.2f}'
        order_data['Amount2'] = f'${amount3:,.2f}'

        total = (price1 * quant1) + (price2 * quant2) + (price3 * quant3)
        order_data['Total'] = f'${total:,.2f}'
  
    return order_data, name, email, date 

order_data_output = get_order_data(8515, 'Order1.xlsx', directory)


#QUESTION 4
def create_invoice (invoice_num, order_data, directory):
    create_pdf(f'{directory}MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)
 
create_invoice(8515, order_data, directory)





#QUESTION 5
def send_email (invoice_num, name, email, date, directory):  
# Enter your gmail user name and password
    gmail_user = 'westgluke@gmail.com'
    gmail_password = 'nzuf jczx esek hzgl'
# List of recipients
    recipients = email
# Email subect
    subject = 'Invoice Files'
# Body of the email
    body = f"""Dear {name},
    
    Attached is the invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phont: (801 555-2372)
    Email: sales@pythontechco.com"""
# Initialize the email message
    message = EmailMessage()
    message['From'] = 'sales@pythontechco.com'
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)
# Add attachment
    path_to_file1 = f'{directory}Invoice {invoice_num}.pdf'
    messsage = add_attachment(path_to_file1, message)     
# Log in to the Gmail server
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
# Send the email
    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'westgluke@gmail.com', '11/13/2023', directory)



#QUESTION 6
filenames = os.listdir()
inv_num = 8515

for file in filenames:
    if 'Order' in str(file) and  '.xlsx' in str(file):
        order_data = get_order_data(inv_num, file, directory)[0]
        order_data_name = get_order_data(inv_num, file, directory)[1]
        order_data_email = get_order_data(inv_num, file, directory)[2]
        order_data_date = get_order_data(inv_num, file, directory)[3]
        create_invoice(inv_num, order_data, directory)
        send_email(inv_num, order_data_name, order_data_email, order_data_date, directory)
        inv_num = inv_num + 1
    else:  
        break





