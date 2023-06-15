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


def extractData(data, pdfno):
    texts = []
    for card in data:
        try:
            texts.append(card['Text']).strip()
        except:
            continue
    text = '|'.join(texts).replace('|m', 'm').replace(' m|', 'm |').replace(' m |', 'm |').replace('|om', 'om').replace(
        ' om|', 'om |').replace(' om |', 'om |')

    count = '|'.join(texts).split('AMOUNT')[1].split('|')

    y = 1
    print(text)
    for _ in range(20):
        data = {}
        data['Bussiness__City'] = 'Jamestown'
        data['Bussiness__Country'] = 'Tennessee, USA'
        data['Bussiness__Description'] = 'We are here to serve you better. Reach out to us in case of any concern or feedbacks.'
        data['Bussiness__Name'] = 'NearBy Electronics'
        data['Bussiness__StreetAddress'] = '3741 Glory Road'
        data['Bussiness__Zipcode'] = '38556'
        try:
            data['Customer__Email'] = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]{2,})', text)[0]
        except:
            data['Customer__Email'] = ''
            partemail = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+)', text)[0]
            emailInd = text.index(partemail)
            text2 = text[emailInd:].replace('|', '', 1).replace(' ', '', 1)

        email = data['Customer__Email']
        if email == '':
            try:
                data['Customer__Email'] = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)',
                                                     text2)[0].replace('co ', 'com').replace('.co', '.com').replace(
                    '.comm', '.com')
            except:
                data['Customer__Email'] = ''

        email = data['Customer__Email']
        if email == '':
            dueDateIndex = text2.index('Due date')

            text3 = text2.replace(text2[dueDateIndex:dueDateIndex + 22], '')
            try:
                data['Customer__Email'] = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)',
                                                     text3)[0].replace('co ', 'com').replace('.co', '.com').replace(
                    '.comm', '.com')
            except:
                data['Customer__Email'] = ''

        data['Customer__Name'] = text.split('BILL')[1].split('|')[1]
        if 'DETAILS' in data['Customer__Name']:
            data['Customer__Name'] = text.split('BILL')[1].split('|')[3]
        oldname = data['Customer__Name']
        data['Customer__Name'] = re.sub(email + r'.*', '', data['Customer__Name'])
        if data['Customer__Name'] == '':
            data['Customer__Name'] = oldname.split(' ')[0] + ' ' + oldname.split(' ')[1]
        if len(data['Customer__Name']) > 60:
            data['Customer__Name'] = data['Customer__Name'].split(' ')[0] + ' ' + data['Customer__Name'].split(' ')[1]

        data['Customer__PhoneNumber'] = re.findall(r'\d{2,}\-\d{3,}\-\d+', text)[0]
        data['Customer__PhoneNumber'] = re.findall(r'\d{2,}\-\d{3,}\-\d+', text)[0]
        print('-------')
        print(text.split(data['Customer__PhoneNumber'])[1])
        data['Customer__Address__line1'] = text.split(data['Customer__PhoneNumber'])[1].split('|')[1]
        data['Customer__Address__line2'] = text.split(data['Customer__PhoneNumber'])[1].split('|')[2]
        city = data['Customer__Address__line2'].removeprefix(' ')
        descrip = ''
        if text.split(data['Customer__PhoneNumber'])[1].split('|')[3].removeprefix(' ').startswith(
                'Due') or city.startswith('ITEM') or city.startswith('Due') or city.startswith('$'):
            if text.split(data['Customer__PhoneNumber'])[1].split('|')[3][0].isdigit():
                data['Customer__Address__line1'] = text.split(data['Customer__PhoneNumber'])[1].split('|')[3].removesuffix(
                    ' ').removeprefix(' ')
                data['Customer__Address__line2'] = text.split(data['Customer__PhoneNumber'])[1].split('|')[
                    5].removesuffix(' ').removeprefix(' ')
                descrip = text.split(data['Customer__PhoneNumber'])[1].split('|')[1] + ' ' + text.split(
                    data['Customer__PhoneNumber'])[1].split('|')[4]
            else:
                address1 = text.split(data['Customer__PhoneNumber'])[1].split('|')[0].removesuffix(' ').removeprefix(' ')
                print('aadd')
                print(address1)
                if address1 == '' or address1.isspace():
                    address1 = text.split(data['Customer__PhoneNumber'])[1].split('|')[1].removesuffix(' ').removeprefix(
                        ' ')
                arrstring = address1.split(' ')
                data['Customer__Address__line2'] = arrstring[len(arrstring) - 1]
                data['Customer__Address__line1'] = address1.replace(data['Customer__Address__line2'], '')

        city = data['Customer__Address__line2']
        if 'Subtotal' in count[y]:
            break
        data['Invoice__BillDetails__Name'] = count[y]
        data['Invoice__BillDetails__Quantity'] = count[y + 1]
        data['Invoice__BillDetails__Rate'] = count[y + 2]
        data['Invoice__Description'] = re.sub(r'.*DETAILS | PAYMENT.*', '', text.replace('|', ''))
        if descrip != '':
            data['Invoice__Description'] = descrip
        if 'PAYMENT' in data['Invoice__Description']:
            try:
                data['Invoice__Description'] = text.replace('|', '').split('Due')[0].split(city)[1]
            except:
                data['Invoice__Description'] = ''

        if data['Invoice__Description'] == '':
            print('duedate desc')
            print(text.split('Due date')[0])
            arrdescription = text.split('Due date')[0].split('|')
            print(len(arrdescription))
            data['Invoice__Description'] = arrdescription[len(arrdescription) - 2]

        data['Invoice__Description'] = data['Invoice__Description'].replace('DETAILS PAYMENT', '').removeprefix(
            ' ').removesuffix(' ')
        data['Invoice__DueDate'] = re.findall(r'\d+\-\d+\-\d+', re.sub(r'.*Due date:', '', text))[0]
        data['Invoice__IssueDate'] = re.findall(r'\d+\-\d+\-\d+', re.sub(r'.*Issue', '', text))[0]
        data['Invoice__Number'] = re.sub(r'.*Invoice# ', '', text).split()[0].replace('|', '')
        data['Invoice__Tax'] = 10
        data['Pdf__NO'] = pdfno
        y += 4

        master_list.append(data)
        df = pandas.DataFrame(master_list)
        df.to_csv('ExtractData.csv', index=False)


readPDF()
