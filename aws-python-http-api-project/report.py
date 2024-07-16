import os
import boto3
import json
import xlsxwriter
import pandas as pd
from sheet import sheet1, sheet2, sheet3

from datetime import datetime, timezone, timedelta
from decimal import Decimal

with open('config.json', 'r') as config_file:
    config = json.load(config_file)

dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])

directory = config['directory']
timestamp = datetime.now().strftime(config['timestamp_format'])


excel_file = os.path.join(
    directory, config['file_name_daily']+f'{timestamp}.xlsx')
csv_file = os.path.join(
    directory, config['file_name_daily']+f'{timestamp}.csv')
excel_weekly_file = os.path.join(
    directory, config['file_name_weekly']+f'{timestamp}.xlsx')
csv_weekly_file = os.path.join(
    directory, config['file_name_weekly']+f'{timestamp}.csv')

summary_file_csv = os.path.join(
    directory, config['sum_name_daily']+f'{timestamp}.csv')
summary_file_excel = os.path.join(
    directory, config['sum_name_daily']+f'{timestamp}.xlsx')
summary_file_weekly_csv = os.path.join(
    directory, config['sum_name_weekly']+f'{timestamp}.csv')
summary_file_weekly_excel = os.path.join(
    directory, config['sum_name_weekly']+f'{timestamp}.xlsx')

quarterly_file = os.path.join(directory, 'quarterly_file.xlsx')
os.makedirs(directory, exist_ok=True)


def calculate_grade(score):
    if score is not None:
        if score >= 90:
            return "A"
        elif score >= 80:
            return "B"
        elif score >= 70:
            return "C"
        elif score >= 60:
            return "D"
        else:
            return "F"
    return None


def cut_score(score):
    if score is not None:
        if score > 100:
            return 100
        elif score < 0:
            return 0
        return score
    return None


def convert_decimal_to_float(item):
    for key, value in item.items():
        if isinstance(value, Decimal):
            item[key] = float(value)
        elif isinstance(value, dict):
            item[key] = convert_decimal_to_float(value)
        elif isinstance(value, list):
            item[key] = [
                convert_decimal_to_float(i) if isinstance(i, dict)
                else float(i) if isinstance(i, Decimal)
                else i for i in value
            ]
    return item


def extract_university_info(df):
    df['university'] = df['university'].apply(json.loads)
    df['university_id'] = df['university'].apply(
        lambda x: x.get('id') if isinstance(x, dict) else None)
    df['university_name'] = df['university'].apply(
        lambda x: x.get('name') if isinstance(x, dict) else None)
    return df


def write_to_excel_and_csv(students, mode):
    for student in students:
        student = convert_decimal_to_float(student)
        student['score'] = cut_score(student.get('score'))
        student['grade'] = calculate_grade(student.get('score'))
        student['university'] = json.dumps(student.get('university'))
        student['enroll_subject'] = json.dumps(student.get('enroll_subject'))

    df = pd.DataFrame(students, columns=config['data_frame'])

    if mode == 'daily':
        df.to_excel(excel_file, index=False)
        df.to_csv(csv_file, index=False)

        df = extract_university_info(df)
        summary = df.groupby(['university_id', 'university_name', 'grade']).size(
        ).reset_index(name='student_count')

        summary.to_csv(summary_file_csv, index=False)
        summary.to_excel(summary_file_excel, index=False)

    if mode == 'weekly':
        df.to_excel(excel_weekly_file, index=False)
        df.to_csv(csv_weekly_file, index=False)

        df = extract_university_info(df)
        summary = df.groupby(['university_id', 'university_name', 'grade']).size(
        ).reset_index(name='student_count')

        summary.to_csv(summary_file_weekly_csv, index=False)
        summary.to_excel(summary_file_weekly_excel, index=False)


def no_data_to_day():
    df = pd.DataFrame(columns=config['data_frame'])
    df.to_excel(excel_file, index=False)
    df.to_csv(csv_file, index=False)
    summary = pd.DataFrame(columns=config['data_frame_summary'])
    summary.to_excel(summary_file_excel, index=False)
    summary.to_csv(summary_file_csv, index=False)


def report_daily(event=None, context=None):
    mode = 'daily'
    table = dynamodb.Table(config['table_name'])

    response = table.scan()
    items = response['Items']

    today_date = datetime.now(timezone.utc).date()
    filtered_items = [
        item for item in items
        if 'date_update' in item and datetime.fromisoformat(item['date_update']).date() == today_date
    ]
    if filtered_items != []:
        write_to_excel_and_csv(filtered_items, mode)
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "Report success"})
        }
    else:
        no_data_to_day()
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "No report to day"})
        }


def report_weekly(event=None, context=None):
    mode = 'weekly'
    table = dynamodb.Table(config['table_name'])

    response = table.scan()
    items = response['Items']
    today_date = datetime.now(timezone.utc).date()
    one_week_ago_date = today_date - timedelta(days=7)
    filtered_items = [
        item for item in items
        if 'date_update' in item and one_week_ago_date <= datetime.fromisoformat(item['date_update']).date() <= today_date
    ]
    if filtered_items:
        write_to_excel_and_csv(filtered_items, mode)
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "Report success"})
        }
    else:
        no_data_to_day()
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "No report this week"})
        }

# def insert_student(event=None, context=None):
#     # Assume config contains your AWS credentials and table name
#     config = {
#         'table_name': 'your_table_name_here'
#     }
    
#     # Creating the DataFrame with appropriate MultiIndex columns
#     columns = pd.MultiIndex.from_product([[' ', 'เบี้ยประกันชีวิต', 'ค้าจ้างหรือค่าบำเหน็จในเดือนนี้'],
#                                           [' ', 'ประเภทสามัญ', 'ประเภทอุตสาหกรรม', 'ประเภทกลุ่ม', 'การรับประกันภัย']],
#                                          names=[' ', 'ประเภท', 'ปี'])
#     df = pd.DataFrame(columns=columns)
    
#     # Fetch data from DynamoDB
#     dynamodb = boto3.resource('dynamodb')
#     table = dynamodb.Table(config['table_name'])
#     response = table.scan()
#     items = response['Items']
    
#     # Populate DataFrame with DynamoDB items
#     for i, item in enumerate(items):
#         df.loc[i+1, (' ', ' ', 'บริษัทประกันชีวิต')] = item.get('name')
#         df.loc[i+1, ('เบี้ยประกันชีวิต', 'ประเภทสามัญ', 'ปีต่อไป')] = int(item.get('score'))
#         df.loc[i+1, ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้', 'ประเภทสามัญ', 'ปีต่อไป')] = item.get('enroll_subject')
#         df.loc[i+1, ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้', 'ประเภทอุตสาหกรรม', 'ปีต่อไป')] = item.get('field')
#         df.loc[i+1, ('เบี้ยประกันชีวิต', 'การรับประกันภัย', 'อุบัติเหตุส่วนบุคคล')] = item.get('year')
    
#     # Calculate total sum for 'ประเภทสามัญ', 'ปีต่อไป'
#     sum_value = df[('เบี้ยประกันชีวิต', 'ประเภทสามัญ', 'ปีต่อไป')].sum()
#     df.loc[len(items)+1, ('เบี้ยประกันชีวิต', 'ประเภทสามัญ', 'ปีต่อไป')] = sum_value
#     df.loc[len(items)+1, (' ', ' ', 'บริษัทประกันชีวิต')] = "รวม"
    
#     # Output to Excel file
#     output_file_path = 'output.xlsx'
#     with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
#         df.to_excel(writer, sheet_name='Sheet1', index=True)
#         workbook = writer.book
#         worksheet = writer.sheets['Sheet1']
        
#         # Define formats for different cell types
#         header_format = workbook.add_format({
#             'align': 'center',
#             'border': 1,
#             'fg_color': '#888888',
#             'font_color': '#ffffff',
#             'valign': 'vcenter'
#         })
#         type_format = workbook.add_format({
#             'align': 'center',
#             'border': 1,
#             'valign': 'vcenter'
#         })
        
#         # Apply formatting to headers and data cells
#         header_ranges = ['D1:J1', 'K1:Q1', 'C3:T3']
#         header_texts = ['เบี้ยประกันชีวิต', 'ค้าจ้างหรือค่าบำเหน็จในเดือนนี้', 'บริษัทประกันชีวิต']
#         for header, text in zip(header_ranges, header_texts):
#             worksheet.write(header, text, header_format)
        
#         type_ranges = ['D2:G2', 'K2:N2', 'C4:T4']
#         type_texts = ['ประเภทสามัญ', 'ประเภทอุตสาหกรรม', 'ประเภทกลุ่ม', 'การรับประกันภัย']
#         for type_range, type_text in zip(type_ranges, type_texts):
#             worksheet.write(type_range, type_text, type_format)
    
#     print(f"Excel file '{output_file_path}' created successfully.")

def insert_student(event=None, context=None):
    sheet_1 = sheet1()
    sheet_2 = sheet2()
    sheet_3 = sheet3()
    df = pd.DataFrame(columns=sheet_1)
    df2 = pd.DataFrame(columns=sheet_2)
    df3 = pd.DataFrame(columns=sheet_3)

    table = dynamodb.Table(config['life_insurance'])
    response = table.scan()
    items = response['Items']
    first_year = 0
    next_year = 0

    for i in range(len(items)):
        df.loc[i+1, ([' '],[' '],['บริษัทประกันชีวิต'])] = items[i].get('name_insurance')
        df.loc[i+1, (['เบี้ยประกันชีวิต'],['ประเภทสามัญ'],['ปีแรก'])] = int(items[i].get('premium'))
    
    sorted_data = sorted(items, key=lambda x: x['type_insurance'])
    for i in range(len(items)):
        df2.loc[i+1, (['แผนความคุ้มครอง'])] = sorted_data[i].get('name_insurance')
        df2.loc[i+1, (['ชื่อ'])] = sorted_data[i].get('insurance')
        df2.loc[i+1, (['ประเภท'])] = sorted_data[i].get('type_insurance')
        df2.loc[i+1, (['ความคุ้มครอง'])] = int(sorted_data[i].get('premium'))
        df2.loc[i+1, (['เบี้ยประกันเพิ่มต่อเนื่อง'])] = int(sorted_data[i].get('rtu'))
        df2.loc[i+1, (['จำนวนเงินเอาประกันภัย'])] = int(sorted_data[i].get('amount_insured'))
        df2.loc[i+1, (['จำนวนประกันภัยรวมที่ต้องชำระต่องวด'])] = int(sorted_data[i].get('rtu')+sorted_data[i].get('premium'))

    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    df2.to_excel(writer, sheet_name='Sheet2')
    df3.to_excel(writer, sheet_name='Sheet3')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet2 = writer.sheets['Sheet2']
    worksheet2.autofit()

    header_color = workbook.add_format({
        'align': 'center',
        'border': 1,
        'fg_color': '#888888',
        'font_color': '#ffffff',
        'valign': 'vcenter'
    })
    industry_type = workbook.add_format({
        'align': 'center',
        'border': 1,
        'fg_color': '#ffcc99',
        'valign': 'vcenter'
    })
    tripod_types = workbook.add_format({
        'align': 'center',
        'border': 1,
        'fg_color': '#ffff99',
        'valign': 'vcenter'
    })
    group_type = workbook.add_format({
        'align': 'center',
        'border': 1,
        'fg_color': '#ccffcc',
        'valign': 'vcenter'
    })
    insurance = workbook.add_format({
        'align': 'center',
        'border': 1,
        'fg_color': '#cc99ff',
        'valign': 'vcenter'
    })
    headers = ['D1:J1', 'K1:Q1', 'C3', 'R3', 'S3', 'T3']
    header_texts = ['เบี้ยประกันชีวิต', 'ค้าจ้างหรือค่าบำเหน็จในเดือนนี้', 'บริษัทประกันชีวิต',
                    'ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้', 'ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน', 'เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน']
    for header, text in zip(headers, header_texts):
        worksheet.write(header, text, header_color)

    commonType = ['D2:E2', 'D3', 'E3', 'K2:L2', 'K3', 'L3']
    commonType_text = ['ประเภทสามัญ', 'ปีแรก',
                       'ปีต่อไป', 'ประเภทสามัญ', 'ปีแรก', 'ปีต่อไป']
    for common, text in zip(commonType, commonType_text):
        worksheet.write(common, text, industry_type)

    tripodTypes = ['F2:G2', 'F3', 'G3', 'M2:N2', 'M3', 'N3']
    tripodTypes_text = ['ประเภทอุตสาหกรรม', 'ปีแรก',
                       'ปีต่อไป', 'ประเภทอุตสาหกรรม', 'ปีแรก', 'ปีต่อไป']
    for tripod, text in zip(tripodTypes, tripodTypes_text):
        worksheet.write(tripod, text, tripod_types)

    groupType = ['H2:I2', 'H3', 'I3', 'O2:P2', 'O3', 'P3']
    groupType_text = ['ประเภทกลุ่ม', 'ปีแรก',
                       'ปีต่อไป', 'ประเภทกลุ่ม', 'ปีแรก', 'ปีต่อไป']
    for group, text in zip(groupType, groupType_text):
        worksheet.write(group, text, group_type)
    
    insuranceType = ['J2', 'J3', 'Q2', 'Q3']
    insuranceType_test = ['การรับประกันภัย', 'อุบัติเหตุส่วนบุคคล', 'การรับประกันภัย', 'อุบัติเหตุส่วนบุคคล']
    for insure, text in zip(insuranceType, insuranceType_test):
        worksheet.write(insure, text, insurance)
    writer._save()
