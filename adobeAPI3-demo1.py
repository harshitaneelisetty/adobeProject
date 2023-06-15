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


def read_pdf():
    for i in range(100):
        zip_file = "./ExtractTextInfoFromPDF.zip"
        if os.path.isfile(zip_file):
            os.remove(zip_file)
        input_pdf = f'output{i}.pdf'
        print(input_pdf)

        credentials = create_credentials()
        execution_context = create_execution_context(credentials)
        extract_pdf_operation = create_extract_pdf_operation()
        set_operation_input(extract_pdf_operation, input_pdf)
        extract_pdf_options = create_extract_pdf_options()
        set_operation_options(extract_pdf_operation, extract_pdf_options)
        result = execute_operation(extract_pdf_operation, execution_context)
        save_result_as(result, zip_file)
        extracted_data = extract_data_from_zip(zip_file)
        process_data(extracted_data, i)


def create_credentials():
    credentials = Credentials.service_account_credentials_builder() \
        .from_file("./pdfservices-api-credentials.json") \
        .build()
    return credentials


def create_execution_context(credentials):
    execution_context = ExecutionContext.create(credentials)
    return execution_context


def create_extract_pdf_operation():
    extract_pdf_operation = ExtractPDFOperation.create_new()
    return extract_pdf_operation


def set_operation_input(operation, input_pdf):
    source = FileRef.create_from_local_file(input_pdf)
    operation.set_input(source)


def create_extract_pdf_options():
    extract_pdf_options = ExtractPDFOptions.builder() \
        .with_element_to_extract(ExtractElementType.TEXT) \
        .build()
    return extract_pdf_options


def set_operation_options(operation, options):
    operation.set_options(options)


def execute_operation(operation, execution_context):
    result = operation.execute(execution_context)
    return result


def save_result_as(result, zip_file):
    result.save_as(zip_file)


def extract_data_from_zip(zip_file):
    with zipfile.ZipFile(zip_file, 'r') as archive:
        json_entry = archive.open('structuredData.json')
        json_data = json_entry.read()
    data = json.loads(json_data)
    data = data['elements']
    json_entry.close()
    return data


def process_data(data, pdf_no):
    master_list = []
    for card in data:
        extracted_data = extract_text(card, pdf_no)
        if extracted_data is not None:
            master_list.append(extracted_data)
        print(master_list)
    df = pandas.DataFrame(master_list)
    df.to_csv('ExtractData.csv', index=False)


def extract_text(card, pdf_no):
        text = card['Text'].strip()
        extracted_data = {}
        extracted_data['Business_City'] = 'Jamestown'
        extracted_data['Business_Country'] = 'Tennessee, USA'
        extracted_data['Business_Description'] = 'We are here to serve you better. Reach out to us in case of any concern or feedbacks.'
        extracted_data['Business_Name'] = 'NearBy Electronics'
        extracted_data['Business_StreetAddress'] = '3741 Glory Road'
        extracted_data['Business_Zipcode'] = '38556'
        extracted_data['Customer_Email'] = extract_email(text)
        if extracted_data['Customer_Email'] == '':
            extracted_data['Customer_Email'] = extract_email_from_text(text)
        extracted_data['Customer_Name'] = extract_customer_name(text)
        extracted_data['Customer_PhoneNumber'] = extract_phone_number(text)
        extracted_data['Customer_Address_line1'] = extract_address_line1(text)
        extracted_data['Customer_Address_line2'] = extract_address_line2(text)
        extracted_data['Invoice_BillDetails_Name'] = extract_bill_details_name(text)
        extracted_data['Invoice_BillDetails_Quantity'] = extract_bill_details_quantity(text)
        extracted_data['Invoice_BillDetails_Rate'] = extract_bill_details_rate(text)
        extracted_data['Invoice_Description'] = extract_invoice_description(text)
        extracted_data['Invoice_DueDate'] = extract_due_date(text)
        extracted_data['Invoice_IssueDate'] = extract_issue_date(text)
        extracted_data['Invoice_Number'] = extract_invoice_number(text)
        extracted_data['Invoice_Tax'] = 10
        extracted_data['Pdf_NO'] = pdf_no

        return extracted_data


def extract_email(text):
    try:
        email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]{2,})', text)[0]
        return email
    except:
        return ''


def extract_email_from_text(text):
    try:
        part_email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+)', text)[0]
        email_ind = text.index(part_email)
        text2 = text[email_ind:].replace('|', '', 1).replace(' ', '', 1)
        extracted_email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)', text2)[0]
        extracted_email = extracted_email.replace('co ', 'com').replace('.co', '.com').replace('.comm', '.com')
        return extracted_email
    except:
        return ''


def extract_customer_name(text):
    customer_name = text.split('BILL')[1].split('|')[1]
    if 'DETAILS' in customer_name:
        customer_name = text.split('BILL')[1].split('|')[3]
    customer_name = re.sub(extract_email(text) + r'.*', '', customer_name)
    if customer_name == '':
        customer_name = customer_name.split(' ')[0] + ' ' + customer_name.split(' ')[1]
    if len(customer_name) > 60:
        customer_name = customer_name.split(' ')[0] + ' ' + customer_name.split(' ')[1]
    return customer_name


def extract_phone_number(text):
    phone_number = re.findall(r'\d{2,}\-\d{3,}\-\d+', text)[0]
    return phone_number


def extract_address_line1(text):
    address_line1 = text.split(extract_phone_number(text))[1].split('|')[1]
    return address_line1


def extract_address_line2(text):
    address_line2 = text.split(extract_phone_number(text))[1].split('|')[2]
    return address_line2


def extract_bill_details_name(text):
    count = '|'.join(text.split('AMOUNT')[1].split('|'))
    name = count.split('|')[1]
    return name


def extract_bill_details_quantity(text):
    count = '|'.join(text.split('AMOUNT')[1].split('|'))
    quantity = count.split('|')[2]
    return quantity


def extract_bill_details_rate(text):
    count = '|'.join(text.split('AMOUNT')[1].split('|'))
    rate = count.split('|')[3]
    return rate


def extract_invoice_description(text):
    description = re.sub(r'.*DETAILS | PAYMENT.*', '', text.replace('|', ''))
    if 'PAYMENT' in description:
        try:
            description = text.replace('|', '').split('Due')[0].split(city)[1]
        except:
            description = ''
    return description


def extract_due_date(text):
    due_date = re.findall(r'\d+\-\d+\-\d+', re.sub(r'.*Due date:', '', text))[0]
    return due_date


def extract_issue_date(text):
    issue_date = re.findall(r'\d+\-\d+\-\d+', re.sub(r'.*Issue', '', text))[0]
    return issue_date


def extract_invoice_number(text):
    invoice_number = re.sub(r'.*Invoice# ', '', text).split()[0].replace('|', '')
    return invoice_number


read_pdf()
