import pandas as pd
import boto3
from boto3.dynamodb.conditions import Attr
import pyzipper
from datetime import datetime
from utils.helpers import ArrayHolder
from utils.get_data import *
from exports.export_pdf import generate_pdf_from_dataframe
from exports.export_xlsx import gen_excel
from asposecells.api import *
import sys
import os, shutil

config = {
    "dynamodb_endpoint": "http://localhost:8000",
}
dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])

file_holder = ArrayHolder()

def create_folder():
    """สร้าง folder เพื่อจัดเก็บไฟล์ PDF และ Excel ที่สร้าง"""
    folder_path = './pdf_xlsx'
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def delete_folder():
    """ลบ folder ที่เก็บไฟล์ PDF และ Excel"""
    folder_path = './pdf_xlsx'
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)

def create_password_protected_zip(folder_path, zip_file_path, password):
    """สร้างไฟล์ ZIP และ set รหัสผ่าน"""
    if not os.path.isdir(folder_path):
        raise FileNotFoundError(f"The folder {folder_path} does not exist")

    with pyzipper.AESZipFile(f"./file/{zip_file_path}", 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zip_file:
        zip_file.setpassword(password.encode('utf-8'))

        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zip_file.write(file_path, arcname)
    
    print(f"Folder '{folder_path}' has been zipped into '{zip_file_path}' with a password")

def scan_data(file_name):
    """สแกนข้อมูลจาก DynamoDB ตามชื่อไฟล์และช่วงวันที่ที่ระบุ"""
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

    elif file_name == 'RPCL003':
        table = dynamodb.Table('claims_cause_analysis_report')
        response = table.scan()
        items = response['Items']
        data = object_RPCL003(items)
        return data

def select_mode():
    """แจ้งให้ผู้ใช้เลือกไฟล์และประเภทการส่งออก และประมวลผลข้อมูลตามนั้น"""
    create_folder()
    file = ['RPCL001', 'RPCL002', 'RPCL003']
    print('Export File : \nRPCL001\nRPCL002\nRPCL003')
    while True:
        file_name = input('File : ')
        if file_name in file:
            while True:
                print('Export type : xlsx/xls , pdf')
                type_export = input('Type : ')
                if type_export in ['xlsx','xls','xlsx/xls','xlsx xls']:
                    if file_name == 'RPCL001':
                        filter_items = input('insurance , product , None \nFilter : ')
                        pd_dataframe = scan_data(file_name)
                        if filter_items == '' or filter_items == 'None':
                            filter_items = None
                        result = gen_excel(pd_dataframe, file_name, filter_items)
                        file_holder.add_value(result)
                        break
                    else:
                        pd_dataframe = scan_data(file_name)
                        result = gen_excel(pd_dataframe, file_name)
                        file_holder.add_value(result)
                        break
                elif type_export == 'pdf':
                    pd_dataframe = scan_data(file_name)
                    if file_name == 'RPCL001':
                        for i, chunk in enumerate([pd_dataframe]):
                            data_transaction = get_customer_report(chunk)
                            pd_dataframe = pd.DataFrame(data_transaction)
                            result = generate_pdf_from_dataframe(pd_dataframe, file_name)
                            file_holder.add_value(result)
                        break
                    else:
                        for i, chunk in enumerate([pd_dataframe]):
                            if file_name == 'RPCL002':
                                data_transaction = get_insurance_company_report(chunk)
                                pd_dataframe = pd.DataFrame(data_transaction)
                                result = generate_pdf_from_dataframe(pd_dataframe, file_name)
                                file_holder.add_value(result)
                            elif file_name == 'RPCL003':
                                data_transaction = get_claims_cause_analysis_report(chunk)
                                pd_dataframe = pd.DataFrame(data_transaction)
                                result = generate_pdf_from_dataframe(pd_dataframe, file_name)
                                file_holder.add_value(result)
                        break
        elif file_name == 'back':
            folder_to_zip = './pdf_xlsx'
            current_dateTime = datetime.now()
            timestamp = convert_time(str(current_dateTime))
            zip_file_name = f'./export_RPCL_{timestamp}.zip'
            password = f'export_RPCL_{timestamp}'
            create_password_protected_zip(folder_to_zip, zip_file_name, password)
            delete_folder()
            break
def main():
    select_mode()

if __name__ == "__main__":
    main()