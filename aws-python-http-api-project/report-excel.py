import pandas as pd
import boto3
from boto3.dynamodb.conditions import Attr
import json
import pyzipper
from datetime import datetime
from collections import defaultdict, Counter
from utils.helpers import ArrayHolder, SheetHolder
from utils.get_data import *
from exports.export_pdf import generate_pdf_from_dataframe
from exports.export_xlsx import gen_excel

from asposecells.api import *
import sys

import os
from fpdf import FPDF
from fpdf.fonts import FontFace

config = {
    "dynamodb_endpoint": "http://localhost:8000",
}
dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])

array_holder = ArrayHolder()
name_holder = SheetHolder()
holder_categories = ArrayHolder()
values_categories = ArrayHolder()

def create_password_protected_zip(files):
    current_dateTime = datetime.now()
    timestamp = convert_time(str(current_dateTime))
    zip_name = f'file/export_RPCL_{timestamp}.zip'
    password = f'export_RPCL_{timestamp}'
    with pyzipper.AESZipFile(zip_name, 'w', compression=pyzipper.ZIP_DEFLATED,
                             encryption=pyzipper.WZ_AES) as zf:
        zf.setpassword(password.encode('utf-8'))
        for file in files:
            zf.write(file)

def scan_data(file_name):
    if file_name == 'RPCL001':
        table = dynamodb.Table('sales_premium_transaction')

        date_to_year = input("Enter year:").strip()
        date_to_month = input("Enter month:").strip().zfill(2)
        
        if date_to_year == '' and date_to_month == '00':
            response = table.scan()
        else:
            response = table.scan(
                FilterExpression=Attr('export_date').gte(f'{date_to_year}-{date_to_month}-01') & Attr('export_date').lte(f'{date_to_year}-{date_to_month}-31')
            )

        items = response.get('Items', [])
        
        if not items:
            sys.exit("No data found for the specified date range.")
        else:
            return items
    elif file_name == 'RPCL002':
        table = dynamodb.Table('insurance_company_report')
        response = table.scan()
        items = response['Items']
        data = object_RPCL002(items)
        return data
def select_mode():
    file = ['RPCL001', 'RPCL002', 'RPCL003']
    print('Export File : \nRPCL001\nRPCL002\nRPCL003')
    while True:
        file_name = input('File : ')
        if file_name in file:
            while True:
                print('Export type : xlsx/xls , pdf')
                type_export = input('Type : ')
                if type_export in ['xlsx','xls','xlsx/xls','xlsx xls']:
                    filter_items = input('insurance , product , None \nFilter : ')
                    gen_excel(pd_dataframe, file_name, filter_items)
                    break
                elif type_export == 'pdf':
                    pd_dataframe = scan_data(file_name)
                    if file_name == 'RPCL001':
                        for i, chunk in enumerate([pd_dataframe]):
                            data_transaction = get_customer_report(chunk)
                            pd_dataframe = pd.DataFrame(data_transaction)
                            generate_pdf_from_dataframe(pd_dataframe, file_name)
                        break
                    else:
                        for i, chunk in enumerate([pd_dataframe]):
                            data_transaction = get_insurance_company_report(chunk)
                            pd_dataframe = pd.DataFrame(data_transaction)
                            generate_pdf_from_dataframe(pd_dataframe, file_name)
                        break
def main():
    select_mode()
    create_password_protected_zip(array_holder.values)

if __name__ == "__main__":
    main()