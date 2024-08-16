def add_style_02(workbook, worksheet, df_concatenated, chunks, user_name,):
    worksheet.set_column('A:O', 9)
    center_format = workbook.add_format({'align': 'center'})

    worksheet.merge_range('A2:O2', 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)', center_format)
    worksheet.merge_range('A5:O5', 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', center_format)
    columns_format = workbook.add_format({'border': 0,'font_color': 'white'})
    columns_to_border = ['#1', '#2', '#3', '#4', '#5', '#6', '#7', '#8', '#9', '#10', '#11', '#12', '#13', '#14', '#15']
    approved_format = workbook.add_format({'align': 'center','valign':'top', 'bg_color': '#00FF00', 'border': 1, 'text_wrap': True,'num_format': '#,##0.00','font_size': 10})
    denied_format = workbook.add_format({'align': 'center','valign':'top', 'bg_color': '#f4cccc', 'border': 1, 'text_wrap': True,'num_format': '#,##0.00','font_size': 10})
    pending_format = workbook.add_format({'align': 'center','valign':'top', 'bg_color': '#ffff00', 'border': 1, 'text_wrap': True,'num_format': '#,##0.00','font_size': 10})
    company_format = workbook.add_format({'align': 'center','valign':'top', 'border': 1, 'text_wrap': True,'num_format': '#,##0.00','font_size': 10})
    test = workbook.add_format({'bg_color': 'black'},{'bg_color': 'red'})
    header = ['ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567','รายงานการเรียกร้องตามบริษัทประกันภัย (Claims by Insurance Company Report)']
    for col in columns_to_border:
        worksheet.write(0, df_concatenated.columns.get_loc(col), col, columns_format)
    for row in range(df_concatenated.shape[0]):
        for col in columns_to_border:
            cell_value = df_concatenated.at[row, col]
            if col in ['#1','#2','#3'] and pd.notna(cell_value) and cell_value not in header:
                worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, company_format)
            if col in ['#4','#5','#6','#7'] and pd.notna(cell_value):
                worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, approved_format)
            if col in ['#8','#9','#10','#11'] and pd.notna(cell_value):
                worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, denied_format)
            if col in ['#12','#13','#14','#15'] and pd.notna(cell_value):
                worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, pending_format)

    worksheet.center_horizontally()
    now = datetime.now()
    thai_year = now.year + 543
    thai_date = now.strftime(f'%d/%m/{thai_year}')
    footer_left = f'&L Users : {user_name}'
    footer_right = f'&R Create At : {thai_date}'
    footer_center = f'&C หน้าที่ &P / {len(chunks)}'

    worksheet.set_footer(footer_left + '    ' + footer_center + '    ' + footer_right)
    worksheet.set_landscape()



    import pandas as pd
import boto3
import io
import xlwt,json
from datetime import datetime
from decimal import Decimal

config = {
    "dynamodb_endpoint": "http://localhost:8000",
    "table_name": "insurance_company_report"
}
dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])
table = dynamodb.Table(config['table_name'])
response = table.scan()
items = response['Items']

def export_xlsx(dataframe: pd.DataFrame, styles: dict, xlsx_index: bool):
    def get_style_in_list(column_name: str, styles: dict):
        style = styles.get(column_name)
        if not style:
            style = styles.get(column_name.split('_last')[0])
        return style

    if len(dataframe.index) != 0:
        buff = io.BytesIO()
        writer = pd.ExcelWriter(buff, engine='xlsxwriter')
        dataframe.to_excel(writer, sheet_name='sheet1', index=xlsx_index)

        if styles:
            workbook = writer.book
            worksheet = writer.sheets['sheet1']
            custom_styles = {key: workbook.add_format(value) for key, value in styles.items()}
            payout_not_equal_claim_format = workbook.add_format({'color': 'red', 'align': 'right', 'border': 1})

            for col_num, column in enumerate(dataframe.columns):
                if column in styles:
                    for row_num, value in enumerate(dataframe[column], start=1):
                        if (column == 'Payout Amount' and dataframe['Payout Amount'][row_num-1] != dataframe['Claim Amount'][row_num-1] and dataframe['Claim Status'][row_num-1] != 'Denied'):
                            worksheet.write(row_num, col_num, value, payout_not_equal_claim_format)
                        else:
                            style = get_style_in_list(column + "_last", custom_styles) if len(dataframe[column]) == 1 else get_style_in_list(column, custom_styles)
                            worksheet.write(row_num, col_num, value, style)
                else:
                    for row_num, value in enumerate(dataframe[column], start=1):
                        worksheet.write(row_num, col_num, value)
                max_len = max(dataframe[column].astype(str).map(len).max(), len(column))
                worksheet.set_column(col_num, col_num, max_len + 2)

        writer.close()
        buff.seek(0)
        return buff.getvalue()
    return ""

def report_excel(event, context): 
    data = json.loads(event['body'])
    if 'template' in data and data['template'] == 'test':
        test()
    elif 'template' in data and data['template'] == 'insurance':
        insurance()
    elif 'template' in data and data['template'] == 'RPCL001':
        header = 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)'
        select_date = 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567'
        RPCL001(header,select_date)
    elif 'template' in data and data['template'] == 'RPCL002':
        header = 'รายงานการเรียกร้องตามบริษัทประกันภัย (Claims by Insurance Company Report)'
        select_date = 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567'
        RPCL002(header,select_date)

def convert_time(date):
    datetime_str = date
    datetime_obj = datetime.fromisoformat(datetime_str)
    formatted_date = datetime_obj.strftime('%Y-%m-%d')
    return formatted_date

def add_style_01(workbook, worksheet, df_concatenated, chunks, user_name):
    worksheet.set_column('A:J', 11)
    center_format = workbook.add_format({'align': 'center'})

    worksheet.merge_range('A2:H2', 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)', center_format)
    worksheet.merge_range('A5:H5', 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', center_format)

    border_format = workbook.add_format({'border': 1,'num_format': '#,##0.00','font_size': 9 })
    date_format = workbook.add_format({'border': 1,'font_size': 9,'align': 'right'})
    columns_format = workbook.add_format({'border': 0,'font_color': 'white'})
    header_format = workbook.add_format({'border': 0,'font_size': 15,'align': 'center','bold': 1})
    columns_data_format = workbook.add_format({'border': 1,'font_size': 9,'align': 'center','bold': 1, 'bg_color' : '#E6F7FF'})
    red_text_format = workbook.add_format({'border': 1,'font_size': 9,'num_format': '#,##0.00','font_color': 'red'})

    columns_to_border = ['#1', '#2', '#3', '#4', '#5', '#6', '#7', '#8']
    data_format = ['Customer Name', 'Loan ID', 'Insurance Company', 'Claim Amount', 'Claim Status', 'Submission Date', 'Approval Date', 'Payout Amount']
    for col in columns_to_border:
        worksheet.write(0, df_concatenated.columns.get_loc(col), col, columns_format)

    for row in range(df_concatenated.shape[0]):
        for col in columns_to_border:
            cell_value = df_concatenated.at[row, col]
            if cell_value in data_format:
                worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, columns_data_format)
            elif cell_value == 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)':
                worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, header_format)
            elif pd.notna(cell_value) and cell_value not in ['รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)', 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567']:
                if col == '#8' and row > 0 and pd.notna(df_concatenated.at[row, '#4']) and cell_value != df_concatenated.at[row, '#4'] and df_concatenated.at[row, '#8'] != 0:
                    worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, red_text_format)
                elif col in ['#6','#7'] :
                    worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, date_format)
                else:
                    worksheet.write(row + 1, df_concatenated.columns.get_loc(col), cell_value, border_format)


    worksheet.center_horizontally()
    now = datetime.now()
    thai_year = now.year + 543
    thai_date = now.strftime(f'%d/%m/{thai_year}')
    footer_left = f'&L Users : {user_name}'
    footer_right = f'&R Create At : {thai_date}'
    footer_center = f'&C หน้าที่ &P / {len(chunks)}'

    worksheet.set_footer(footer_left + '    ' + footer_center + '    ' + footer_right)

def add_style_02(dataframe: pd.DataFrame, styles: dict):
    def get_style_in_list(column_name: str, styles: dict):
        style = styles.get(column_name)
        if not style:
            style = styles.get(column_name.split('_last')[0])
        return style

    if len(dataframe.index) != 0:
        buff = io.BytesIO()
        writer = pd.ExcelWriter(buff, engine='xlsxwriter')

        dataframe.to_excel(writer, sheet_name='sheet1', index=False)

        # Get the xlsxwriter workbook and worksheet objects
        if styles:
            workbook = writer.book
            worksheet = writer.sheets['sheet1']

            custom_styles = {}
            for style_key, style_value in styles.items():
                custom_styles[style_key] = workbook.add_format(style_value)

            for style_key, style_value in styles.items():
                for col_num, column in enumerate(dataframe.columns):
                    amount_row = dataframe[column].__len__()
                    if column == style_key:
                        for row_num, value in enumerate(dataframe[column], start=1):
                            print(row_num)
                            if value != 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567' and value != 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)':
                                style = get_style_in_list(column + "_last", custom_styles) if row_num == amount_row else get_style_in_list(column, custom_styles)
                                worksheet.write(row_num, col_num, dataframe[column][row_num - 1], style)

        writer.close()
        buff.seek(0)
        buff2 = buff.getvalue()

        return buff2
    return ""


def split_data(data, chunk_size):
    return [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]

def RPCL002(header, select_date):
    company_format = {'align': 'center','valign':'top', 'border': 1, 'text_wrap': True,'num_format': '#,##0.00','font_size': 10}
    styles_mapping = {
        '#2' : company_format,
    }
    response = table.scan()
    items = response.get('Items', [])
    user_name = 'test user name'
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
    chunk_size = 20
    chunks = split_data(items_list, chunk_size)
    # with pd.ExcelWriter('RPCL002.xlsx', engine='xlsxwriter') as writer:
    for i, chunk in enumerate(chunks):
        header = {
                '#1': [header, None, None, select_date, None, None],
                '#2': None,
                '#3': None,
                '#4': None,
                '#5': None,
                '#6': None,
                '#7': None,
                '#8': None,
                '#9': None,
                '#10': None,
                '#11': None,
                '#12': None,
                '#13': None,
                '#14': None,
                '#15': None,
            }
        output_header = {
            '#1',
            '#2',
            '#3',
            '#4',
            '#5',
            '#6',
            '#7',
            '#8',
            '#9',
            '#10',
            '#11',
            '#12',
            '#13',
            '#14',
            '#15',
        }
        data_dataframe = {
            '#1': ['Insurance Company'],
            '#2': ['Number of Claims'],
            '#3': ['Amount of Claim'],
            '#4': ['Approved Claims'],
            '#5': ['Amount of Died'],
            '#6': ['Amount of Permanent Disability'],
            '#7': ['Amount of Temporary Disability'],
            '#8': ['Denied Claims'],
            '#9': ['Amount of Died'],
            '#10': ['Amount of Permanent Disability'],
            '#11': ['Amount of Temporary Disability'],
            '#12': ['Pending Claims'],
            '#13': ['Amount of Died'],
            '#14': ['Amount of Permanent Disability'],
            '#15': ['Amount of Temporary Disability']
        }
        data_transaction = []
        for company, data in chunk:
            input_data = {
                '#1': company,
                '#2': data['total_claims'],
                '#3': data['amount_types']['sum'],
                '#4': data['amount_types']['approved_claims']['total_claim'],
                '#5': data['amount_types']['approved_claims']['died'],
                '#6': data['amount_types']['approved_claims']['permanent_disability'],
                '#7': data['amount_types']['approved_claims']['temporary_disability'],
                '#8': data['amount_types']['denied_claims']['total_claim'],
                '#9': data['amount_types']['denied_claims']['died'],
                '#10': data['amount_types']['denied_claims']['permanent_disability'],
                '#11': data['amount_types']['denied_claims']['temporary_disability'],
                '#12': data['amount_types']['pending_claims']['total_claim'],
                '#13': data['amount_types']['pending_claims']['died'],
                '#14': data['amount_types']['pending_claims']['permanent_disability'],
                '#15': data['amount_types']['pending_claims']['temporary_disability'],
            }
            data_transaction.append(input_data)
            
        for transaction in data_transaction:
            data_dataframe['#1'].append(transaction['#1'])
            data_dataframe['#2'].append(transaction['#2'])
            data_dataframe['#3'].append(transaction['#3'])
            data_dataframe['#4'].append(transaction['#4'])
            data_dataframe['#5'].append(transaction['#5'])
            data_dataframe['#6'].append(transaction['#6'])
            data_dataframe['#7'].append(transaction['#7'])
            data_dataframe['#8'].append(transaction['#8'])
            data_dataframe['#9'].append(transaction['#9'])
            data_dataframe['#10'].append(transaction['#10'])
            data_dataframe['#11'].append(transaction['#11'])
            data_dataframe['#12'].append(transaction['#12'])
            data_dataframe['#13'].append(transaction['#13'])
            data_dataframe['#14'].append(transaction['#14'])
            data_dataframe['#15'].append(transaction['#15'])

        data_dataframe = pd.DataFrame(data_dataframe)
        dataframe = pd.DataFrame(header)
        concatenated_df = pd.concat([dataframe, data_dataframe], ignore_index=True)
        # df_concatenated.to_excel(writer, sheet_name=f'Sheet_{i+1}', index=False)
        # workbook = writer.book
        # worksheet = writer.sheets[f'Sheet_{i+1}']
        result = add_style_02(concatenated_df, styles_mapping)
        with open(f'RPCL002_{i+1}.xlsx', 'wb') as f:
            f.write(result)
    
def RPCL001(header, select_date):
    def split_data(data, chunk_size):
        return [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
        
    user_name = 'test user name'
    response = table.scan()
    items = response.get('Items', [])
    data = items
    chunk_size = 20
    chunks = split_data(data, chunk_size)
    writer = pd.ExcelWriter('RPCL001.xlsx', engine='xlsxwriter')
    for i, chunk in enumerate(chunks):
        header = {
            '#1': [header, None, None, select_date, None, None],
            '#2': None,
            '#3': None,
            '#4': None,
            '#5': None,
            '#6': None,
            '#7': None,
            '#8': None
        }
        data_dataframe = {
            '#1': ['Customer Name'],
            '#2': ['Loan ID'],
            '#3': ['Insurance Company'],
            '#4': ['Claim Amount'],
            '#5': ['Claim Status'],
            '#6': ['Submission Date'],
            '#7': ['Approval Date'],
            '#8': ['Payout Amount']
        }
        data_transaction = []
        for item in chunk:
            sub_date = item.get("submission_date", '')
            approval_date = item.get("approval_date", '')
            input_data = {
                '#1': item.get("customer_name", ''),
                '#2': item.get("loan_id", ''),
                '#3': item.get("insurance_company", ''),
                '#4': item.get("claim_amount", ''),
                '#5': item.get("claim_status", ''),
                '#6': convert_time(sub_date),
                '#7': convert_time(approval_date),
                '#8': item.get("payout_amount", '')
            }
            data_transaction.append(input_data)
        
        for transaction in data_transaction:
            data_dataframe['#1'].append(transaction['#1'])
            data_dataframe['#2'].append(transaction['#2'])
            data_dataframe['#3'].append(transaction['#3'])
            data_dataframe['#4'].append(transaction['#4'])
            data_dataframe['#5'].append(transaction['#5'])
            data_dataframe['#6'].append(transaction['#6'])
            data_dataframe['#7'].append(transaction['#7'])
            data_dataframe['#8'].append(transaction['#8'])
        
        data_dataframe = pd.DataFrame(data_dataframe)
        dataframe = pd.DataFrame(header)
        df_concatenated = pd.concat([dataframe, data_dataframe], ignore_index=True)
        df_concatenated.to_excel(writer, sheet_name=f'Sheet_{i+1}', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets[f'Sheet_{i+1}']

        add_style_01(workbook, worksheet, df_concatenated, chunks, user_name)

    writer._save()


    # style_align_right = {'align': 'right', 'border': 1}
    # style_border = {'border': 1}
    # style_hide = {'color': 'white', 'border': 0}
    # styles_mapping = {
    #     'Claim Amount': style_align_right,
    #     'Payout Amount': style_align_right,
    #     'Submission Date': style_align_right,
    #     'Approval Date': style_align_right,
    #     'Customer Name': style_border,
    #     'Loan ID': style_border,
    #     'Insurance Company': style_border,
    #     'Claim Status': style_border,
    #     ' ': style_hide
    # }
    # data_header_mapper = {
    #     '#': ' ',
    #     'customer_name': 'Customer Name',
    #     'loan_id': 'Loan ID',
    #     'insurance_company': 'Insurance Company',
    #     'claim_amount': 'Claim Amount',
    #     'claim_status': 'Claim Status',
    #     'submission_date': 'Submission Date',
    #     'approval_date': 'Approval Date',
    #     'payout_amount': 'Payout Amount'
    # }
    # output_header = [' ', 'Customer Name', 'Loan ID', 'Insurance Company', 'Claim Amount', 'Claim Status', 'Submission Date', 'Approval Date', 'Payout Amount']
    # data_transaction = []

    # for item in items:
    #     sub_date = item.get("submission_date", '')
    #     approval_date = item.get("approval_date", '')
    #     input_data = {
    #         '#': None,
    #         'customer_name': item.get("customer_name", ''),
    #         'loan_id': item.get("loan_id", ''),
    #         'insurance_company': item.get("insurance_company", ''),
    #         'claim_amount': item.get("claim_amount", ''),
    #         'claim_status': item.get("claim_status", ''),
    #         'submission_date': convert_time(sub_date),
    #         'approval_date': convert_time(approval_date),
    #         'payout_amount': item.get("payout_amount", '')
    #     }
    #     data_transaction.append(input_data)

    # pd_dataframe = pd.DataFrame(data_transaction)
    # pd_dataframe.rename(columns=data_header_mapper, inplace=True)
    # pd_dataframe = pd_dataframe.reindex(columns=output_header, fill_value='')

    # result = export_xlsx(pd_dataframe, styles_mapping, xlsx_index=False)
    # with open('RPCL001.xlsx', 'wb') as f:
    #     f.write(result)

def insurance():
    style_align_center = {'align': 'center', 'fg_color': '#fe0000'}
    styles_mapping = {
        'เบี้ยประกันชีวิต' : style_align_center
    }

    data_header_mapper = {
        '#1': ('','',''),
        '#2': ('','',''),
        'life_insurance_company': ('','','บริษัทประกันชีวิต'),
        'first_year_01': ('เบี้ยประกันชีวิต','ประเภทสามัญ', 'ปีแรก'),
        'next_year_01': ('เบี้ยประกันชีวิต','ประเภทสามัญ', 'ปีต่อไป'),
        'first_year_02': ('เบี้ยประกันชีวิต','ประเภทอุตสาหกรรม', 'ปีแรก'),
        'next_year_02': ('เบี้ยประกันชีวิต','ประเภทอุตสาหกรรม', 'ปีต่อไป'),
        'first_year_03': ('เบี้ยประกันชีวิต','ประเภทอุตสาหกรรม', 'ปีแรก'),
        'next_year_03': ('เบี้ยประกันชีวิต','ประเภทอุตสาหกรรม', 'ปีต่อไป'),
        'personal_accident_01': ('เบี้ยประกันชีวิต','การรับประกันภัย','อุบัติเหตุส่วนบุคคล'),
        'first_year_04': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทสามัญ', 'ปีแรก'),
        'next_year_04': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทสามัญ', 'ปีต่อไป'),
        'first_year_05': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทอุตสาหกรรม', 'ปีแรก'),
        'next_year_05': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทอุตสาหกรรม', 'ปีต่อไป'),
        'first_year_06': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทอุตสาหกรรม', 'ปีแรก'),
        'next_year_06': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทอุตสาหกรรม', 'ปีต่อไป'),
        'personal_accident_02': ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','การรับประกันภัย','อุบัติเหตุส่วนบุคคล'),
        'this_month': ('','','ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้'),
        'due_at_month_end': ('','','ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน'),
        'insurance_premiums': ('','','เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน')
    }

    output_header = [
        ('','',''),
        ('','','บริษัทประกันชีวิต'),
        ('เบี้ยประกันชีวิต','ประเภทสามัญ', 'ปีแรก'),
        ('เบี้ยประกันชีวิต','ประเภทสามัญ', 'ปีต่อไป'),
        ('เบี้ยประกันชีวิต','ประเภทกลุ่ม', 'ปีแรก'),
        ('เบี้ยประกันชีวิต','ประเภทกลุ่ม', 'ปีต่อไป'),
        ('เบี้ยประกันชีวิต','ประเภทอุตสาหกรรม', 'ปีแรก'),
        ('เบี้ยประกันชีวิต','ประเภทอุตสาหกรรม', 'ปีต่อไป'),
        ('เบี้ยประกันชีวิต','การรับประกันภัย','อุบัติเหตุส่วนบุคคล'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทสามัญ', 'ปีแรก'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทสามัญ', 'ปีต่อไป'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทกลุ่ม', 'ปีแรก'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทกลุ่ม', 'ปีต่อไป'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทอุตสาหกรรม', 'ปีแรก'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','ประเภทอุตสาหกรรม', 'ปีต่อไป'),
        ('ค้าจ้างหรือค่าบำเหน็จในเดือนนี้','การรับประกันภัย','อุบัติเหตุส่วนบุคคล'),
        (' ',' ','ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้'),
        (' ',' ','ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน'),
        (' ',' ','เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน')
    ]
    data_transaction = []
    for index, item in enumerate(items):
        input_data = {
            '#': '',
            'life_insurance_company': item.get("insurance", ''),
            'first_year_01': item.get("premium", ''),
            'next_year_01': item.get("rtu", ''),
            'first_year_02': item.get("premium", ''),
            'next_year_02': item.get("rtu", ''),
            'first_year_03': item.get("premium", ''),
            'next_year_03': item.get("rtu", ''),
            'personal_accident_01': item.get("name_insurance", ''),
            'first_year_04': item.get("premium", ''),
            'next_year_04': item.get("rtu", ''),
            'first_year_05': item.get("premium", ''),
            'next_year_05': item.get("rtu", ''),
            'first_year_06': item.get("premium", ''),
            'next_year_06': item.get("rtu", ''),
            'personal_accident_02': '',
            'this_month': '',
            'due_at_month_end': '',
            'insurance_premiums': ''
        }
        data_transaction.append(input_data)
    pd_dataframe = pd.DataFrame.from_dict(data_transaction)
    pd_dataframe.rename(columns=data_header_mapper, inplace=True)
    header = pd.MultiIndex.from_tuples(output_header)
    pd_dataframe.columns = header
    xlsx_index = True
    result = export_xlsx(pd_dataframe, styles_mapping, xlsx_index)
    with open('insurance_report.xlsx', 'wb') as f:
        f.write(result)

def test():
    style_align_center = {'align': 'center','fg_color': '#fe0000'}
    style_font_red = {'fg_color': '#fe0000'}
    styles_mapping = {
        'ชื่อ' : style_align_center
    }
    data_header_mapper = {
        '#': '#',
        'type_insurance': 'ประเภทความคุ้มครอง',
        'name_insurance': 'แผนความคุ้มครอง',
        'insurance': 'ชื่อ',
        'rpp': 'เบี้ยประกันหลักเพื่อความคุ้มครอง',
        'rtu': 'เบี้ยประกันภัยต่อเนื่อง',
        'sum_insured': 'จำนวนประกันภัยรวมที่ต้องชำระต่องวด',
        'amount_insured': 'จำนวนเงินเอาประกัน'
    }
    output_header = [
        '#',
        'ประเภทความคุ้มครอง',
        'แผนความคุ้มครอง',
        'ชื่อ',
        'เบี้ยประกันหลักเพื่อความคุ้มครอง',
        'เบี้ยประกันภัยต่อเนื่อง',
        'จำนวนประกันภัยรวมที่ต้องชำระต่องวด',
        'จำนวนเงินเอาประกัน'
    ]
    data = []
    for index in range(len(items)):
        premium = int(items[index].get('premium', 0))
        rtu = int(items[index].get('rtu', 0))
        input = {
            '#': index,
            'type_insurance': items[index].get('type_insurance'),
            'name_insurance': items[index].get('name_insurance'),
            'insurance': items[index].get('insurance'),
            'rpp': premium,
            'rtu': rtu,
            'sum_insured': premium + rtu,
            'amount_insured': items[index].get('amount_insured')
        }
        data.append(input)

    data_frame = pd.DataFrame(columns=['#', 'type_insurance', 'name_insurance',
                            'insurance', 'rpp', 'rtu', 'sum_insured', 'amount_insured'], data=data)
    data_transaction = []
    for index, row in data_frame.iterrows():
        data = {
            '#': index + 1,
            'type_insurance': row["type_insurance"],
            'name_insurance': row["name_insurance"],
            'insurance': row["insurance"],
            'rpp': row["rpp"],
            'rtu': row["rtu"],
            'sum_insured': row["sum_insured"],
            'amount_insured': row["amount_insured"]
        }
        data_transaction.append(data)

    pd_dataframe = pd.DataFrame.from_dict(data_transaction)
    pd_dataframe.rename(columns=data_header_mapper, inplace=True)
    pd_dataframe = pd_dataframe.reindex(
        columns=output_header, fill_value='')
    xlsx_index = False   
    result = export_xlsx(pd_dataframe, styles_mapping, xlsx_index)

    with open('test1.xlsx', 'wb') as f:
        f.write(result)
