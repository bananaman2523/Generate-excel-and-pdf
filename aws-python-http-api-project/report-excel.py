import pandas as pd
import boto3
import io
import xlwt,json

config = {
    "dynamodb_endpoint": "http://localhost:8000",
    "table_name": "life_insurance"
}
dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])
table = dynamodb.Table(config['table_name'])
response = table.scan()
items = response['Items']

def export_xlsx(dataframe: pd.DataFrame, styles: dict,xlsx_index):
    def get_style_in_list(column_name: str, styles: dict):
        style = styles.get(column_name)
        if not style:
            style = styles.get(column_name.split('_last')[0])
        return style

    if len(dataframe.index) != 0:
        buff = io.BytesIO()
        writer = pd.ExcelWriter(buff, engine='xlsxwriter')

        dataframe.to_excel(writer, sheet_name='sheet1', index=xlsx_index)

        # Get the xlsxwriter workbook and worksheet objects
        if styles:
            workbook = writer.book
            worksheet = writer.sheets['sheet1']

            custom_styles = {}
            for style_key, style_value in styles.items():
                custom_styles[style_key] = workbook.add_format(style_value)
            group_type = workbook.add_format({
                'align': 'center',
                'border': 1,
                'fg_color': '#ccffcc',
                'valign': 'vcenter'
            })
            groupType = ['H2:I2', 'H3', 'I3', 'O2:P2', 'O3', 'P3']
            groupType_text = ['ประเภทกลุ่ม', 'ปีแรก',
                            'ปีต่อไป', 'ประเภทกลุ่ม', 'ปีแรก', 'ปีต่อไป']
            for group, text in zip(groupType, groupType_text):
                worksheet.write(group, text, group_type)
            for style_key, style_value in styles.items():
                for col_num, column in enumerate(dataframe.columns):
                    print("column======>",column)
                    amount_row = dataframe[column].__len__()
                    if column == style_key:
                        for row_num, value in enumerate(dataframe[column], start=1):
                            style = get_style_in_list(
                                column+"_last", custom_styles) if row_num == amount_row else get_style_in_list(column, custom_styles)
                            worksheet.write(0, col_num, dataframe[column][row_num-1], style)

        writer.close()
        buff.seek(0)
        buff2 = buff.getvalue()

        return buff2
    return ""

def report_excel(event, context): 
    data = json.loads(event['body'])
    if 'template' in data and data['template'] == 'test':
        test()
    elif 'template' in data and data['template'] == 'insurance':
        insurance()

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

# def insurance():
#     style_align_center = {'align': 'center', 'fg_color': '#fe0000'}
#     styles_mapping = {
#         'บริษัทประกันชีวิต' : style_align_center
#     }

#     data_header_mapper = {
#         '#1': '',
#         '#2': '',
#         'life_insurance_company': 'บริษัทประกันชีวิต',
#         'first_year_01': 'ปีแรก',
#         'next_year_01': 'ปีต่อไป',
#         'first_year_02': 'ปีแรก',
#         'next_year_02': 'ปีต่อไป',
#         'first_year_03': 'ปีแรก',
#         'next_year_03': 'ปีต่อไป',
#         'personal_accident_01': 'อุบัติเหตุส่วนบุคคล',
#         'first_year_04': 'ปีแรก',
#         'next_year_04': 'ปีต่อไป',
#         'first_year_05': 'ปีแรก',
#         'next_year_05': 'ปีต่อไป',
#         'first_year_06': 'ปีแรก',
#         'next_year_06': 'ปีต่อไป',
#         'personal_accident_02': 'อุบัติเหตุส่วนบุคคล',
#         'this_month': 'ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้',
#         'due_at_month_end': 'ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน',
#         'insurance_premiums': 'เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน'
#     }

#     output_header = [
#         '#1',
#         '#2',
#         'บริษัทประกันชีวิต',
#         'ปีแรก',
#         'ปีต่อไป',
#         'ปีแรก',
#         'ปีต่อไป',
#         'ปีแรก',
#         'ปีต่อไป',
#         'อุบัติเหตุส่วนบุคคล',
#         'ปีแรก',
#         'ปีต่อไป',
#         'ปีแรก',
#         'ปีต่อไป',
#         'ปีแรก',
#         'ปีต่อไป',
#         'อุบัติเหตุส่วนบุคคล',
#         'ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้',
#         'ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน',
#         'เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน'
#     ]
#     data_transaction = []
#     for index, item in enumerate(items):
#         input_data = {
#             '#': '',
#             'life_insurance_company': item.get("insurance", ''),
#             'first_year_01': item.get("premium", ''),
#             'next_year_01': item.get("rtu", ''),
#             'first_year_02': item.get("premium", ''),
#             'next_year_02': item.get("rtu", ''),
#             'first_year_03': item.get("premium", ''),
#             'next_year_03': item.get("rtu", ''),
#             'personal_accident_01': item.get("name_insurance", ''),
#             'first_year_04': item.get("premium", ''),
#             'next_year_04': item.get("rtu", ''),
#             'first_year_05': item.get("premium", ''),
#             'next_year_05': item.get("rtu", ''),
#             'first_year_06': item.get("premium", ''),
#             'next_year_06': item.get("rtu", ''),
#             'personal_accident_02': '',
#             'this_month': '',
#             'due_at_month_end': '',
#             'insurance_premiums': ''
#         }
#         data_transaction.append(input_data)
#     pd_dataframe = pd.DataFrame.from_dict(data_transaction)
#     pd_dataframe.rename(columns=data_header_mapper, inplace=True)
#     pd_dataframe = pd_dataframe.reindex(
#         columns=output_header, fill_value='')
#     xlsx_index = False
#     result = export_xlsx(pd_dataframe, styles_mapping, xlsx_index)

#     with open('test2.xlsx', 'wb') as f:
#         f.write(result)





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
