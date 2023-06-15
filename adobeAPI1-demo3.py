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
dat = [{} for _ in range(100)]
fileIndex = 0


def readPDF():
    global fileIndex
    for i in range(100):
        
        zip_file = "./ExtractTextInfoFromPDF.zip"
        if os.path.isfile(zip_file):
            os.remove(zip_file)
        input_pdf = f'output{i}.pdf'
        print(input_pdf)
        #Initial setup, create credentials instance.
        credentials = Credentials.service_account_credentials_builder()\
            .from_file("./pdfservices-api-credentials.json") \
            .build()

        #Create an ExecutionContext using credentials and create a new operation instance.
        execution_context = ExecutionContext.create(credentials)

        extract_pdf_operation = ExtractPDFOperation.create_new()

        #Set operation input from a source file.
        source = FileRef.create_from_local_file(input_pdf)
        extract_pdf_operation.set_input(source)

        #Build ExtractPDF options and set them into the operation
        extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
            .with_element_to_extract(ExtractElementType.TEXT) \
            .build()
        extract_pdf_operation.set_options(extract_pdf_options)


        #Execute the operation.
        result: FileRef = extract_pdf_operation.execute(execution_context)

        #Save the result to the specified location.
        result.save_as(zip_file)


        with zipfile.ZipFile(zip_file, 'r') as archive:
            jsonentry = archive.open('structuredData.json')
            jsondata = jsonentry.read()
        data = json.loads(jsondata)
        data = data['elements']
        jsonentry.close()
        
        extractData(data)
        fileIndex += 1


def extractData(data):
    global dat
    texts = []
    for card in data:
        try:
            texts.append(card['Text']).strip()
        except:
            continue
    text = '|'.join(texts).replace('hotm |ail','hotmail').replace('otm ail','hotmail')
    count = '|'.join(texts).split('AMOUNT')[1].split('|')

    y = 1
    for index, key in enumerate(dat):
        print(f"Index: {index}, Key: {key}")
    for _ in range(100):
        
        print(data[_])
        dat[_]['Bussiness__City'] = 'Jamestown'
        dat[_]['Bussiness__Country'] = 'Tennessee, USA'
        dat[_]['Bussiness__Description'] = 'We are here to serve you better. Reach out to us in case of any concern or feedbacks.'
        dat[_]['Bussiness__Name'] = 'NearBy Electronics'
        dat[_]['Bussiness__StreetAddress'] = '3741 Glory Road'
        dat[_]['Bussiness__Zipcode'] = '38556'
        try:
            dat[_]['Customer__Email'] = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)',text)[0].replace('co ','com').replace('.co','.com').replace('.comm','.com')
        except:
            dat[_]['Customer__Email'] = ''

        email = dat[_]['Customer__Email']
        dat[_]['Customer__Name'] = text.split('BILL')[1].split('|')[1]
        if 'DETAILS' in dat[_]['Customer__Name']:
            dat[_]['Customer__Name'] = text.split('BILL')[1].split('|')[3]
        dat[_]['Customer__Name'] = re.sub(email + r'.*','',dat[_]['Customer__Name'])
        dat[_]['Customer__PhoneNumber'] = re.findall(r'\d{2,}\-\d{3,}\-\d+',text)[0]
        dat[_]['Customer__Address__line1'] = text.split(dat[_]['Customer__PhoneNumber'])[1].split('|')[1]
        dat[_]['Customer__Address__line2'] = text.split(dat[_]['Customer__PhoneNumber'])[1].split('|')[2]
        city = dat[_]['Customer__Address__line2']

        if 'Subtotal' in count[y]:            
            break
        dat[_]['Invoice__BillDetails__Name'] = count[y]
        dat[_]['Invoice__BillDetails__Quantity'] = count[y+1]
        dat[_]['Invoice__BillDetails__Rate'] = count[y+2]
        dat[_]['Invoice__Description'] = re.sub(r'.*DETAILS | PAYMENT.*','',text.replace('|',''))
        if 'PAYMENT' in dat[_]['Invoice__Description']:
            try:
                dat[_]['Invoice__Description'] = text.replace('|','').split('Due')[0].split(city)[1]
            except:
                dat[_]['Invoice__Description'] = ''
        dat[_]['Invoice__DueDate'] = re.findall(r'\d+\-\d+\-\d+',re.sub(r'.*Due date:','',text))[0]
        dat[_]['Invoice__IssueDate'] = re.findall(r'\d+\-\d+\-\d+',re.sub(r'.*Issue','',text))[0]
        dat[_]['Invoice__Number'] =re.sub(r'.*Invoice# ','',text).split()[0].replace('|','')
        dat[_]['Invoice__Tax'] = 10
        dat[_]['Index'] = fileIndex
        y += 4

        master_list.append(dat[_])
    
    df = pandas.DataFrame(master_list)
    df.to_csv('ExtractData.csv', mode='a', index=False, header=False)
        

readPDF()