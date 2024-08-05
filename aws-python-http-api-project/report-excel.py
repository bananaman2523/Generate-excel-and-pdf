import pandas as pd
import boto3
import json
from datetime import datetime
from dataclasses import dataclass, field
from collections import defaultdict
import zipfile
import pyzipper
from typing import List

config = {
    "dynamodb_endpoint": "http://localhost:8000",
}
dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])

def set_paper(sheet_name: str, writer: pd.ExcelWriter, user_name, chunks):
    worksheet = writer.sheets[sheet_name]
    now = datetime.now()
    thai_year = now.year + 543
    thai_date = now.strftime(f'%d/%m/{thai_year}')
    footer_left = f'&L Users : {user_name}'
    footer_right = f'&R Create At : {thai_date}'
    footer_center = f'&C หน้าที่ &P / {chunks}'

    worksheet.set_footer(footer_left + '    ' + footer_center + '    ' + footer_right)

def export_xlsx(dataframe: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str, xlsx_index: bool = False):
    if len(dataframe.index) != 0:
        dataframe.to_excel(writer, sheet_name=sheet_name, index=xlsx_index)

def add_style(dataframe: pd.DataFrame, styles: dict, writer: pd.ExcelWriter, len_template, custom_header, header_style, header_titles, file, sheet_name: str):
    if styles:
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        custom_styles = {}
        for style_key, style_value in styles.items():
            custom_styles[style_key] = workbook.add_format(style_value)

        columns_to_border = dataframe.columns
        columns_format = workbook.add_format({'border': 0, 'font_color': 'white', 'align' : 'right'})

        for col in columns_to_border:
            worksheet.write(0, dataframe.columns.get_loc(col), col, columns_format)

        for style_key, style_value in styles.items():
            for col_num, column in enumerate(dataframe.columns):
                amount_row = dataframe[column].__len__()
                if column == style_key:
                    for row_num, value in enumerate(dataframe[column], start=1):
                        if row_num >= len_template + 1:
                            style = get_style_in_list(column+"_last", custom_styles) if row_num == amount_row else get_style_in_list(column, custom_styles)
                            worksheet.write(row_num, col_num, dataframe[column][row_num-1], style)
                        if file == 'RPCL001':
                            claim_amount = dataframe.get('claim_amount')
                            payout_amount = dataframe.get('payout_amount')

                            if claim_amount is not None and payout_amount is not None:
                                if dataframe.at[row_num-1, 'claim_amount'] != dataframe.at[row_num-1, 'payout_amount'] and dataframe.at[row_num-1, 'payout_amount'] != 0:
                                    red_font_style = workbook.add_format({'font_color': 'red','font_size': 9, 'num_format': '#,##0.00', 'border': 1,'text_wrap' : True, 'valign': 'top'})
                                    worksheet.write(row_num, dataframe.columns.get_loc('payout_amount'), dataframe.at[row_num-1, 'payout_amount'], red_font_style)

        add_style_column(worksheet, header_style, header_titles, writer, dataframe)
        merge_cells(worksheet, dataframe, len(dataframe.columns), custom_header, writer)

def add_style_column(worksheet, header_style, header_titles, writer, dataframe):
    workbook = writer.book
    column_style_map = {}
    for style_entry in header_style:
        columns = style_entry.get('column', [])
        style = workbook.add_format(style_entry.get('style', {}))
        for column in columns:
            column_style_map[column] = style

    for col_num, column in enumerate(dataframe.columns):
        for row_num, value in enumerate(dataframe[column], start=1):
            header_format = column_style_map.get(column)
            if value in header_titles:
                worksheet.write(row_num, col_num, value, header_format)

def merge_cells(worksheet, dataframe, min_consecutive, custom_header, writer):
    workbook = writer.book
    add_header_values = [item['add_header'] for item in custom_header if isinstance(item, dict) and 'add_header' in item]
    add_header_styles = [workbook.add_format(item['style']) for item in custom_header if isinstance(item, dict) and 'style' in item]
    
    num_rows, num_cols = dataframe.shape
    
    for row in range(num_rows):
        start_col = 0
        while start_col < num_cols:
            current_value = dataframe.iloc[row, start_col]

            if pd.isna(current_value):
                start_col += 1
                continue

            end_col = start_col
            while end_col + 1 < num_cols and dataframe.iloc[row, end_col + 1] == current_value:
                end_col += 1
            if end_col > start_col and (end_col - start_col + 1) >= min_consecutive:
                style_header = None
                for header, style in zip(add_header_values, add_header_styles):
                    if current_value == header and pd.notna(current_value):
                        style_header = style
                        break
                
                if style_header:
                    worksheet.merge_range(row + 1, start_col, row + 1, end_col, current_value, style_header)
                else:
                    worksheet.merge_range(row + 1, start_col, row + 1, end_col, current_value)
                
            start_col = end_col + 1

def custom_style(style, style_format):
    styles_mapping = {}
    for header, style_name in style_format.items():
        if style_name in style:
            styles_mapping[header] = style[style_name]
    return styles_mapping

def get_style_in_list(column_name: str, styles: dict):
    style = styles.get(column_name)
    if not style:
        style = styles.get(column_name.split('_last')[0])
    return style

def add_space(pd_dataframe):
    num_rows = 1
    if num_rows <= 0:
        return pd_dataframe
    empty_rows = pd.DataFrame([[None] * len(pd_dataframe.columns)] * num_rows, columns=pd_dataframe.columns)
    
    pd_dataframe_with_space = pd.concat([empty_rows, pd_dataframe], ignore_index=True)
    
    return pd_dataframe_with_space

def convert_time(date):
    if date:
        # datetime_str = date
        datetime_obj = datetime.fromisoformat(date)
        formatted_date = datetime_obj.strftime('%Y-%m-%d')
        return formatted_date

def add_header(pd_dataframe, header):
    header_df = pd.DataFrame([[header]* (len(pd_dataframe.columns))], columns=pd_dataframe.columns)
    concatenated_df = pd.concat([header_df, pd_dataframe], ignore_index=True)
    return concatenated_df

def report_excel(data): 
    # data = json.loads(event['body'])
    if data == 'RPCL001':
        filter = ['insurance', 'product', None]
        for i in filter:
            template_input = {
                'file_name' : 'RPCL001',
                'data_per_page' : 20,
                'filter' : i,
                'language' : 'th',
                'table' : 'sales_premium_transaction',
                'user_name' : 'test test',
                'set_column' : 11,
                'output_header': [
                    {'title': 'number', 'variable': ' '},
                    {'title': 'customer_name', 'variable': 'Customer Name'},
                    {'title': 'loan_id', 'variable': 'Loan ID'},
                    {'title': 'insurance_company', 'variable': 'Insurance Company'},
                    {'title': 'claim_amount', 'variable': 'Claim Amount'},
                    {'title': 'claim_status', 'variable': 'Claim Status'},
                    {'title': 'submission_date', 'variable': 'Submission Date'},
                    {'title': 'approval_date', 'variable': 'Approval Date'},
                    {'title': 'payout_amount', 'variable': 'Payout Amount'}
                ],
                'header_style' : [
                    {
                        'column' : ['number','claim_amount','payout_amount','submission_date','approval_date','customer_name','loan_id','insurance_company','claim_status'], 
                        'style' : {'align': 'center', 'valign': 'top', 'border': 1, 'font_size': 9, 'bold' : 1,'fg_color' :'#E6F7FF','text_wrap' : True}
                    }
                ],
                'style': {
                    'style_align_right': {'align': 'right', 'border': 1, 'font_size': 9,'text_wrap' : True, 'valign' : 'top'},
                    'style_border': {'font_size': 9, 'border': 1,'text_wrap' : True, 'valign' : 'top'},
                    'style_number': {'font_size': 9, 'num_format': '#,##0.00', 'border': 1,'text_wrap' : True, 'valign' : 'top'},
                    'style_index': {'font_size': 10, 'border': 1, 'valign' : 'top', 'align' : 'center'},
                    'center_header' : {'font_size': 15, 'bold' : 1, 'align' : 'center', 'valign' : 'top'}
                },
                'style_format': {
                    'number': 'style_index',
                    'claim_amount': 'style_number',
                    'payout_amount': 'style_number',
                    'submission_date': 'style_align_right',
                    'approval_date': 'style_align_right',
                    'customer_name': 'style_border',
                    'loan_id': 'style_border',
                    'insurance_company': 'style_border',
                    'claim_status': 'style_border'
                },
                'excel_template': [
                    {'add_header': 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)', 'style': {'align' : 'center', 'bold' : '1', 'font_size' : 15}},
                    'add_space',
                    {'add_header': 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', 'style': {'align' : 'center'}},
                    'add_space'
                ]
            }
            gen_excel(template_input)
    elif data == 'RPCL002':
        template_input = {
            'file_name' : 'RPCL002',
            'data_per_page' : 20,
            'table' : 'insurance_company_report',
            'user_name' : 'test test',
            'header' : 'รายงานการเรียกร้องตามบริษัทประกันภัย (Claims by Insurance Company Report)',
            'select_date' : 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567',
            'set_column' : 9,
            'output_header': [
                {'title': 'number', 'variable': ' '},
                {'title': 'insurance_company', 'variable': 'Insurance Company'},
                {'title': 'number_of_claims', 'variable': 'Number of Claims'},
                {'title': 'amount_of_claim', 'variable': 'Amount of Claim'},
                {'title': 'approved_claims', 'variable': 'Approved Claims'},
                {'title': 'approved_died', 'variable': 'Amount of Died'},
                {'title': 'approved_permanent_disability', 'variable': 'Amount of Permanent Disability'},
                {'title': 'approved_temporary_disability', 'variable': 'Amount of Temporary Disability'},
                {'title': 'denied_claims', 'variable': 'Denied Claims'},
                {'title': 'denied_died', 'variable': 'Amount of Died'},
                {'title': 'denied_permanent_disability', 'variable': 'Amount of Permanent Disability'},
                {'title': 'denied_temporary_disability', 'variable': 'Amount of Temporary Disability'},
                {'title': 'pending_claims', 'variable': 'Pending Claims'},
                {'title': 'pending_died', 'variable': 'Amount of Died'},
                {'title': 'pending_permanent_disability', 'variable': 'Amount of Permanent Disability'},
                {'title': 'pending_temporary_disability', 'variable': 'Amount of Temporary Disability'}
            ],
            'header_style' : [
                {
                    'column' : ['number','insurance_company','number_of_claims','amount_of_claim'], 
                    'style' : {'align': 'center', 'valign': 'top', 'border': 1, 'font_size': 9, 'bold' : 1, 'text_wrap' : True}
                },
                {
                    'column' : ['approved_claims','approved_died','approved_permanent_disability','approved_temporary_disability'], 
                    'style' : {'align': 'center', 'valign': 'top', 'border': 1, 'font_size': 9, 'bold' : 1, 'text_wrap' : True, 'fg_color' : '#00ff00'}
                },
                {
                    'column' : ['denied_claims','denied_died','denied_permanent_disability','denied_temporary_disability'], 
                    'style' : {'align': 'center', 'valign': 'top', 'border': 1, 'font_size': 9, 'bold' : 1, 'text_wrap' : True, 'fg_color' : '#f4cccc'}
                },
                {
                    'column' : ['pending_claims','pending_died','pending_permanent_disability','pending_temporary_disability'], 
                    'style' : {'align': 'center', 'valign': 'top', 'border': 1, 'font_size': 9, 'bold' : 1, 'text_wrap' : True, 'fg_color' : '#ffff00'}
                }
            ],
            'style': {
                'style_align_center': {'align': 'center', 'valign': 'top','border' : 1},
                'style_border': {'font_size': 9, 'border': 1, 'text_wrap' : True},
                'style_number': {'font_size': 9, 'border': 1, 'text_wrap' : True, 'num_format': '#,##0.00'},
                'style_approve': {'font_size': 9, 'border': 1,'text_wrap' : True, 'fg_color' : '#00ff00'},
                'style_number_approve': {'font_size': 9, 'num_format': '#,##0.00', 'border': 1,'text_wrap' : True, 'fg_color' : '#00ff00'},
                'style_denied': {'font_size': 9, 'border': 1,'text_wrap' : True, 'fg_color' : '#f4cccc'},
                'style_number_denied': {'font_size': 9, 'num_format': '#,##0.00', 'border': 1,'text_wrap' : True, 'fg_color' : '#f4cccc'},
                'style_pending': {'font_size': 9, 'border': 1,'text_wrap' : True, 'fg_color' : '#ffff00'},
                'style_number_pending': {'font_size': 9, 'num_format': '#,##0.00', 'border': 1,'text_wrap' : True, 'fg_color' : '#ffff00'}
            },
            'style_format': {
                'number' : 'style_align_center',
                'number_of_claims': 'style_border',
                'amount_of_claim': 'style_number',
                'approved_claims': 'style_approve',
                'approved_died': 'style_number_approve',
                'approved_permanent_disability': 'style_number_approve',
                'approved_temporary_disability': 'style_number_approve',
                'denied_claims': 'style_denied',
                'denied_died': 'style_number_denied',
                'denied_permanent_disability': 'style_number_denied',
                'denied_temporary_disability': 'style_number_denied',
                'pending_claims': 'style_pending',
                'pending_died': 'style_number_pending',
                'pending_permanent_disability': 'style_number_pending',
                'pending_temporary_disability': 'style_number_pending',
                'insurance_company': 'style_border'
            },
            'excel_template': [
                {'add_header': 'รายงานการเรียกร้องตามบริษัทประกันภัย (Claims by Insurance Company Report)', 'style': {'font_size' : 14, 'align' : 'center', 'bold' : '1'}},
                'add_space',
                {'add_header': 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', 'style': { 'align': 'center'}},
                'add_space'
            ]
        }
        gen_excel(template_input)
    elif data == 'RPCL003':
        template_input = {
            'file_name' : 'RPCL003',
            'data_per_page' : 20,
            'set_column' : 17,
            'table' : 'claims_cause_analysis_report',
            'user_name' : 'test test',
            'header' : 'รายงานการวิเคราะห์สาเหตุการเรียกร้อง (Claims Cause Analysis Report)',
            'select_date' : 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567',
            'output_header': [
                {'title': 'number', 'variable': ' '},
                {'title': 'claim_cause', 'variable': 'Claim Cause'},
                {'title': 'number_of_claims', 'variable': 'Number of \nClaims'},
                {'title': 'percentage_total', 'variable': 'Percentage of \nTotal Claims (%)'},
                {'title': 'percentage_died', 'variable': 'Percentage of \nDied (%)'},
                {'title': 'percentage_permanent', 'variable': 'Percentage of \nPermanent Disability (%)'},
                {'title': 'percentage_temporary', 'variable': 'Percentage of \nTemporary Disability (%)'}
            ],
            'header_style' : [
                {
                    'column' : ['number','claim_cause','number_of_claims','percentage_total','percentage_died','percentage_permanent','percentage_temporary'], 
                    'style' : {'align': 'center', 'valign': 'top', 'border': 1, 'font_size': 9, 'bold' : 1,'text_wrap' : False}
                }
            ],
            'style': {
                'style_align_center': {'align': 'left', 'valign': 'top','border' : 1,'font_size': 9},
                'style_percentage': { 'align': 'right','border' : 1,'font_size': 9},
                'style_index': { 'align': 'center','valign' : 'top' ,'border' : 1,'font_size': 9}
            },
            'style_format': {
                'number' : 'style_index',
                'claim_cause' : 'style_align_center',
                'number_of_claims' : 'style_percentage',
                'percentage_total' : 'style_percentage',
                'percentage_died' : 'style_percentage',
                'percentage_permanent' : 'style_percentage',
                'percentage_temporary' : 'style_percentage'
            },
            'excel_template': [
                {'add_header': 'รายงานการเรียกร้องตามบริษัทประกันภัย (Claims by Insurance Company Report)', 'style': {'font_size' : 14, 'align' : 'center', 'bold' : '1'}},
                'add_space',
                {'add_header': 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', 'style': { 'align': 'center'}},
                'add_space'
            ]
        }
        gen_excel(template_input)

def start_function(functions_to_call, pd_dataframe):
    function_dict = {
        'add_space' : add_space,
        'add_header' : lambda df, header: add_header(df, header)
    }
    for func_info in reversed(functions_to_call):
        if isinstance(func_info, dict):
            func_name, arg = list(func_info.items())[0]
            func = function_dict.get(func_name)
            if func:
                pd_dataframe = func(pd_dataframe, arg)
        else:
            func = function_dict.get(func_info)
            if func:
                pd_dataframe = func(pd_dataframe)

    return pd_dataframe

def get_customer_report(chunk, language = 'th'):
    data_transaction = []
    number = 1
    for item in chunk:
        sub_date = item.get("created_date", '')
        approval_date = item.get("export_date", '')

        insurance_company = json.loads(item.get("insurer", '{}'))
        if 'en' not in insurance_company.get('insurer_name', {}).get('M', {}):
            insurance_company_name_th = insurance_company.get('insurer_name', {}).get('M', {}).get('th', {}).get('S', {})
        else:
            insurance_company_name_th = insurance_company.get('insurer_name', {}).get('M', {}).get(language, {}).get('S', {})
        loan_product = json.loads(item.get("loan_product", '{}'))
        loan_product_name_th = loan_product.get('name', {}).get('M', {}).get(language, {}).get('S', {})
        activity_status = json.loads(item.get("activity_status", '{}'))
        activity_status = activity_status.get('name', {}).get('M', {}).get(language, {}).get('S', {})

        total_fee_amount_str = item.get("total_fee_amount", '0')
        total_fee_amount = float(total_fee_amount_str) if total_fee_amount_str else 0

        total_fee_rate_str = item.get("total_fee_rate", '0')
        total_fee_rate = float(total_fee_rate_str) if total_fee_rate_str else 0
        input_data = {
            'number' : number,
            'customer_name': item.get("customer_full_name", ''),
            'loan_id': loan_product_name_th,
            'insurance_company': insurance_company_name_th,
            'claim_amount': total_fee_amount,
            'claim_status': activity_status,
            'submission_date': convert_time(sub_date),
            'approval_date': convert_time(approval_date),
            'payout_amount': total_fee_rate
        }
        data_transaction.append(input_data)
        number += 1
    return data_transaction

def get_insurance_company_report(chunk):
    data_transaction = []
    number = 1
    for company, data in chunk:
        input_data = {
            'number' : number,
            'insurance_company': company,
            'number_of_claims': data['total_claims'],
            'amount_of_claim': data['amount_types']['sum'],
            'approved_claims': data['amount_types']['approved_claims']['total_claim'],
            'approved_died': data['amount_types']['approved_claims']['died'],
            'approved_permanent_disability': data['amount_types']['approved_claims']['permanent_disability'],
            'approved_temporary_disability': data['amount_types']['approved_claims']['temporary_disability'],
            'denied_claims': data['amount_types']['denied_claims']['total_claim'],
            'denied_died': data['amount_types']['denied_claims']['died'],
            'denied_permanent_disability': data['amount_types']['denied_claims']['permanent_disability'],
            'denied_temporary_disability': data['amount_types']['denied_claims']['temporary_disability'],
            'pending_claims': data['amount_types']['pending_claims']['total_claim'],
            'pending_died': data['amount_types']['pending_claims']['died'],
            'pending_permanent_disability': data['amount_types']['pending_claims']['permanent_disability'],
            'pending_temporary_disability': data['amount_types']['pending_claims']['temporary_disability']
        }
        number += 1
        data_transaction.append(input_data)
    return data_transaction

def get_claims_cause_analysis_report(chunk):
    data_transaction = []
    number = 1
    for claim_cause, data in chunk:
        input_data = {
            'number' : number,
            'claim_cause': claim_cause,
            'number_of_claims': data['sum'],
            'percentage_total': data['percentage_of_total'],
            'percentage_died' : data['types']['died']['percentage'],
            'percentage_permanent' : data['types']['permanent_disability']['percentage'],
            'percentage_temporary' : data['types']['temporary_disability']['percentage']
        }
        number += 1
        data_transaction.append(input_data)
    return data_transaction

def object_RPCL003(items):
    claims_cause_analysis_report = {}
    
    for item in items:
        claim_cause = item.get('claim_cause')
        percentage = item.get('percentage')
        
        if claim_cause and percentage:
            if claim_cause not in claims_cause_analysis_report:
                claims_cause_analysis_report[claim_cause] = {
                    'types': {},
                    'sum': 0
                }

            if percentage not in claims_cause_analysis_report[claim_cause]['types']:
                claims_cause_analysis_report[claim_cause]['types'][percentage] = 0

            claims_cause_analysis_report[claim_cause]['types'][percentage] += 1
            claims_cause_analysis_report[claim_cause]['sum'] += 1

    for claim_cause, data in claims_cause_analysis_report.items():
        for percentage, count in data['types'].items():
            data['types'][percentage] = {
                'count': count,
                'percentage': f'{int(round((count / data['sum']) * 100, 0))}%'
            }

    total_sum = sum(data['sum'] for data in claims_cause_analysis_report.values())
    for claim_cause, data in claims_cause_analysis_report.items():
        data['percentage_of_total'] = f'{int(round((data['sum'] / total_sum) * 100, 0))}%'
        
    items_list = list(claims_cause_analysis_report.items())
    return items_list

def object_RPCL002(items):
    insurance_company_data = {}
    for item in items:
        insurance_company = item.get('insurance_company')
        type_claim = item.get('type_claim')
        amount_type = item.get('amount_type')
        amount = item.get('amount', 0)

        if insurance_company and type_claim:
            if insurance_company not in insurance_company_data:
                insurance_company_data[insurance_company] = {
                    'type_claims': {},
                    'amount_types': {
                        'approved_claims': {
                            'permanent_disability': 0,
                            'died': 0,
                            'temporary_disability': 0,
                            'total_claim' : 0
                        },
                        'denied_claims': {
                            'permanent_disability': 0,
                            'died': 0,
                            'temporary_disability': 0,
                            'total_claim' : 0
                        },
                        'pending_claims': {
                            'permanent_disability': 0,
                            'died': 0,
                            'temporary_disability': 0,
                            'total_claim' : 0
                        },
                        'sum' : 0
                    },
                    'total_claims' : 0
                }
            insurance_company_data[insurance_company]['total_claims'] += 1
            
            if amount_type:
                if type_claim == 'approved_claims':
                    insurance_company_data[insurance_company]['amount_types']['approved_claims'][amount_type] += amount
                    insurance_company_data[insurance_company]['amount_types']['approved_claims']['total_claim'] += 1
                elif type_claim == 'denied_claims':
                    insurance_company_data[insurance_company]['amount_types']['denied_claims'][amount_type] += amount
                    insurance_company_data[insurance_company]['amount_types']['denied_claims']['total_claim'] += 1
                elif type_claim == 'pending_claims':
                    insurance_company_data[insurance_company]['amount_types']['pending_claims'][amount_type] += amount
                    insurance_company_data[insurance_company]['amount_types']['pending_claims']['total_claim'] += 1
            
            insurance_company_data[insurance_company]['amount_types']['sum'] += amount
        
    items_list = list(insurance_company_data.items())
    return items_list

def group_export(data, filter, language='th'):
    grouped = defaultdict(list)
    if filter == 'insurance':
        for item in data:
            insurer_data = json.loads(item['insurer'])
            if 'en' in insurer_data.get('insurer_name', {}).get('M', {}):
                insurer_name = insurer_data.get('insurer_name', {}).get('M', {}).get('th', {}).get('S')
            else:
                insurer_name = insurer_data.get('insurer_name', {}).get('M', {}).get(language, {}).get('S')
            if insurer_name:
                grouped[insurer_name].append(item)

        result = [group for i, group in enumerate(grouped.values())]
        return result
    elif filter == 'product':
        for item in data:
            insurer_data = json.loads(item['loan_product'])
            insurer_name = insurer_data.get('name', {}).get('M', {}).get(language, {}).get('S')
            if insurer_name:
                grouped[insurer_name].append(item)

        result = [group for i, group in enumerate(grouped.values())]
        return result

def split_data(data, chunk_size, group= False):
    if group:   
        flattened_list = []
        for item in data:
            split_data = [item[i:i + chunk_size] for i in range(0, len(item), chunk_size)]
            flattened_list.append(split_data)
        result = [inner for outer in flattened_list for inner in outer]
        return result
    else:
        result = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
        return result

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

def gen_excel(template_input):
    user_name = template_input['user_name']
    chunk_size = template_input['data_per_page']

    filter_type = template_input.get('filter')
    file_name = template_input.get('file_name', '')
    language = template_input.get('language')

    if filter_type == None:
        group = False
    else:
        filter = template_input['filter']
        group = True
    
    if filter_type is None:
        file_name = f'{file_name}.xlsx'
    else:
        suffix = ''
        if filter_type == 'insurance':
            suffix = '_insurance'
        elif filter_type == 'product':
            suffix = '_product'
        file_name = f'{file_name}{suffix}.xlsx'

    table = dynamodb.Table(template_input['table'])
    response = table.scan()
    items = response['Items']
    if template_input['table'] == 'insurance_company_report':
        data = object_RPCL002(items)
    elif template_input['table'] == 'claims_cause_analysis_report':
        data = object_RPCL003(items)
    else:
        data = items

    if group and template_input['table'] == 'sales_premium_transaction':
        data = group_export(data, filter, language)
        chunks = split_data(data, chunk_size, group)
    else:
        chunks = split_data(data, chunk_size)

    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        for i, chunk in enumerate(chunks):
            if template_input['table'] == 'sales_premium_transaction':
                data_transaction = get_customer_report(chunk, language)
            elif template_input['table'] == 'insurance_company_report':
                data_transaction = get_insurance_company_report(chunk)
            elif template_input['table'] == 'claims_cause_analysis_report':
                data_transaction = get_claims_cause_analysis_report(chunk)
            pd_dataframe = pd.DataFrame(data_transaction)

            # Prepare the header row
            header_titles = [item['variable'] for item in template_input['output_header']]
            header_dict = {item['variable']: item['title'] for item in template_input['output_header']}
            pd_dataframe.rename(columns=header_dict, inplace=True)

            style = custom_style(template_input['style'], template_input['style_format'])
            header_row = pd.DataFrame([header_titles], columns=pd_dataframe.columns)
            pd_dataframe = pd.concat([header_row, pd_dataframe], ignore_index=True)

            pd_dataframe = start_function(template_input['excel_template'], pd_dataframe)
            sheet_name = f'Sheet_{i+1}'
            export_xlsx(pd_dataframe, writer, sheet_name=sheet_name)
            # add style
            len_template = len(template_input['excel_template'])
            custom_header = template_input['excel_template']
            header_style = template_input['header_style']
            file = template_input['file_name']
            add_style(pd_dataframe, style, writer, len_template, custom_header, header_style, header_titles, file, sheet_name=sheet_name)

            # add footer
            set_paper(sheet_name, writer, user_name, len(chunks))
            worksheet = writer.sheets[sheet_name]
            if template_input['table'] == 'sales_premium_transaction' or 'claims_cause_analysis_report':
                worksheet.set_portrait()
                worksheet.center_horizontally()
                worksheet.set_margins(left=0.25, right=0.25)
            elif template_input['table'] == 'insurance_company_report':
                worksheet.set_landscape()
                worksheet.center_horizontally()
                worksheet.set_margins(left=0.25, right=0.25)

            worksheet.set_paper(9)
            for col_num, column in enumerate(pd_dataframe.columns):
                if column != 'number':
                    worksheet.set_column(col_num, col_num, template_input['set_column'])
                else:
                    worksheet.set_column(col_num, col_num, 5)

    array_holder.add_value(f'{file_name}')

@dataclass
class ArrayHolder:
    values: List[int] = field(default_factory=list)

    def add_value(self, value: int):
        """Add a new value to the array."""
        self.values.append(value)

    def __repr__(self):
        return f"ArrayHolder(values={self.values})"

array_holder = ArrayHolder()

def main():
    report_excel(data = 'RPCL001')
    report_excel(data = 'RPCL002')
    report_excel(data = 'RPCL003')
    create_password_protected_zip(array_holder.values)

if __name__ == "__main__":
    main()