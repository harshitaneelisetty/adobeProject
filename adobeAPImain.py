from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType

import os.path
import zipfile
import json
import pandas
import re


master_list = []


def readPDF():
    for i in range(100):
        zip_file = "./ExtractTextInfoFromPDF.zip"
        if os.path.isfile(zip_file):
            os.remove(zip_file)
        input_pdf = f'output{i}.pdf'
        print(input_pdf)

        # Initial setup, create credentials instance.
        credentials = Credentials.service_account_credentials_builder() \
            .from_file("./pdfservices-api-credentials.json") \
            .build()

        # Create an ExecutionContext using credentials and create a new operation instance.
        execution_context = ExecutionContext.create(credentials)

        extract_pdf_operation = ExtractPDFOperation.create_new()

        # Set operation input from a source file.
        source = FileRef.create_from_local_file(input_pdf)
        extract_pdf_operation.set_input(source)

        # Build ExtractPDF options and set them into the operation
        extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
            .with_element_to_extract(ExtractElementType.TEXT) \
            .build()
        extract_pdf_operation.set_options(extract_pdf_options)

        # Execute the operation.
        result: FileRef = extract_pdf_operation.execute(execution_context)

        # Save the result to the specified location.
        result.save_as(zip_file)

        with zipfile.ZipFile(zip_file, 'r') as archive:
            jsonentry = archive.open('structuredData.json')
            jsondata = jsonentry.read()
        data = json.loads(jsondata)
        data = data['elements']
        jsonentry.close()
        extractData(data, i)

def parseText(texts):
    return '|'.join(texts).replace('|m', 'm').replace(' m|', 'm |').replace(' m |', 'm |').replace('|om', 'om').replace(
        ' om|', 'om |').replace(' om |', 'om |')

def parseEmail(text):
    try:
        email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]{2,})', text)[0]
    except:
        email = ''
        partemail = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+)', text)[0]
        emailInd = text.index(partemail)
        text2 = text[emailInd:].replace('|', '', 1).replace(' ', '', 1)

    if email == '':
        try:
            email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)',
                                                    text2)[0].replace('co ', 'com').replace('.co', '.com').replace(
                '.comm', '.com')
        except:
            email = ''

    if email == '':
        dueDateIndex = text2.index('Due date')
        text3 = text2.replace(text2[dueDateIndex:dueDateIndex + 22], '')
        try:
            email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)',
                                                    text3)[0].replace('co ', 'com').replace('.co', '.com').replace(
                '.comm', '.com')
        except:
            email = ''
    return email


def parseName(text, email):
    cust_name = text.split('BILL')[1].split('|')[1]
    if 'DETAILS' in cust_name:
        cust_name = text.split('BILL')[1].split('|')[3]
    oldname = cust_name
    cust_name = re.sub(email + r'.*', '', cust_name)
    if cust_name == '':
        cust_name = oldname.split(' ')[0] + ' ' + oldname.split(' ')[1]
    if len(cust_name) > 25:
        cust_name = cust_name.split(' ')[0] + ' ' + cust_name.split(' ')[1]
    print(cust_name)
    return cust_name


def parsePhone(text):
    cust_phone = re.findall(r'\d{2,}\-\d{3,}\-\d+', text)[0]
    cust_phone = re.findall(r'\d{2,}\-\d{3,}\-\d+', text)[0]
    print('-------')
    print(text.split(cust_phone)[1])
    return cust_phone

def parseAddress(text, cust_phone):
    cust_add1 = text.split(cust_phone)[1].split('|')[1]
    cust_add2 = text.split(cust_phone)[1].split('|')[2]
    city = cust_add2.removeprefix(' ')
    descrip = ''
    city = cust_add2.removeprefix(' ')
    if text.split(cust_phone)[1].split('|')[3].removeprefix(' ').startswith(
            'Due') or city.startswith('ITEM') or city.startswith('Due') or city.startswith('$'):
        if text.split(cust_phone)[1].split('|')[3][0].isdigit():
            cust_add1 = text.split(cust_phone)[1].split('|')[3].removesuffix(
                ' ').removeprefix(' ')
            cust_add2 = text.split(cust_phone)[1].split('|')[
                5].removesuffix(' ').removeprefix(' ')
            descrip = text.split(cust_phone)[1].split('|')[1] + ' ' + text.split(
                cust_phone)[1].split('|')[4]
        else:
            address1 = text.split(cust_phone)[1].split('|')[0].removesuffix(' ').removeprefix(' ')
            print('aadd')
            print(address1)
            if address1 == '' or address1.isspace():
                address1 = text.split(cust_phone)[1].split('|')[1].removesuffix(' ').removeprefix(
                    ' ')
            arrstring = address1.split(' ')
            cust_add2 = arrstring[len(arrstring) - 1]
            cust_add1 = address1.replace(cust_add2, '')
        if 'ITEM' in cust_add2:
            cust_add2 = cust_add1[len(cust_add1) - 9 :]
            cust_add1 = cust_add1[:len(cust_add1) - 9]
            
    return cust_add1, cust_add2, descrip, city


def parseInvDescription(text, city, descrip):
    inv_description = re.sub(r'.*DETAILS | PAYMENT.*', '', text.replace('|', ''))
    if descrip != '':
        inv_description = descrip

    if 'PAYMENT' in inv_description:
        try:
            inv_description = text.replace('|', '').split('Due')[0].split(city)[1]
        except:
            inv_description = ''

    if inv_description == '':
        print('duedate desc')
        print(text.split('Due date')[0])
        arrdescription = text.split('Due date')[0].split('|')
        print(len(arrdescription))
        inv_description = arrdescription[len(arrdescription) - 2]

    inv_description = inv_description.replace('DETAILS PAYMENT', '').removeprefix(
        ' ').removesuffix(' ')
    if 'Due date' in inv_description:
        idx = inv_description.index('Due date')
        inv_description = inv_description[:idx - 7]
    if '@' in inv_description:
        inv_description = ''
    return inv_description


def parseDueDate(text):
    return re.findall(r'\d+\-\d+\-\d+', re.sub(r'.*Due date:', '', text))[0]


def parseIssueDate(text):
    return re.findall(r'\d+\-\d+\-\d+', re.sub(r'.*Issue', '', text))[0]


def parseInvNum(text):
    return re.sub(r'.*Invoice# ', '', text).split()[0].replace('|', '')


def extractData(data, pdfno):
    texts = []
    for card in data:
        try:
            texts.append(card['Text']).strip()
        except:
            continue

    text = parseText(texts)

    count = '|'.join(texts).split('AMOUNT')[1].split('|')

    y = 1
    print(text)
    for _ in range(15):
        email = parseEmail(text)
        cust_name = parseName(text, email)
        cust_phone = parsePhone(text)
        cust_add1, cust_add2, descrip, city = parseAddress(text, cust_phone)
        city = cust_add2
        if 'Subtotal' in count[y]:
            break
        bill_name = count[y]
        bill_qty = count[y + 1]
        bill_rate = count[y + 2]
        inv_description = parseInvDescription(text, city, descrip)
        inv_due_date = parseDueDate(text)
        inv_issue_date = parseIssueDate(text)
        inv_number = parseInvNum(text)
        file = pdfno
        y += 4

        saveData(email, cust_name, cust_phone, cust_add1, cust_add2, bill_name, bill_qty, bill_rate, inv_description, inv_due_date,inv_issue_date, inv_number, file)


def saveData(email, cust_name, cust_phone, cust_add1, cust_add2, bill_name, bill_qty, bill_rate, inv_description, inv_due_date,inv_issue_date, inv_number, file):
    dat = {}
    dat['Bussiness__City'] = 'Jamestown'
    dat['Bussiness__Country'] = 'Tennessee, USA'
    dat['Bussiness__Description'] = 'We are here to serve you better. Reach out to us in case of any concern or feedbacks.'
    dat['Bussiness__Name'] = 'NearBy Electronics'
    dat['Bussiness__StreetAddress'] = '3741 Glory Road'
    dat['Bussiness__Zipcode'] = '38556'
    dat['Customer__Email'] = email
    dat['Customer__Name'] = cust_name
    dat['Customer__PhoneNumber'] = cust_phone
    dat['Customer__Address__line1'] = cust_add1
    dat['Customer__Address__line2'] = cust_add2
    dat['Invoice__BillDetails__Name'] = bill_name
    dat['Invoice__BillDetails__Quantity'] = bill_qty
    dat['Invoice__BillDetails__Rate'] = bill_rate
    dat['Invoice__Description'] = inv_description
    dat['Invoice__DueDate'] = inv_due_date
    dat['Invoice__IssueDate'] = inv_issue_date
    dat['Invoice__Number'] =inv_number
    dat['Invoice__Tax'] = 10
    dat['Pdf__No'] = file

    master_list.append(dat)
    df = pandas.DataFrame(master_list)
    df.to_csv('ExtractData.csv', index = False)


readPDF()
