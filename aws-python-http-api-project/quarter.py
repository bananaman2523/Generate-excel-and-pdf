import pandas as pd
import boto3

config = {
     "dynamodb_endpoint": "http://localhost:8000",
     "table_name" : "life_insurance"
}
dynamodb = boto3.resource('dynamodb', endpoint_url=config['dynamodb_endpoint'])
table = dynamodb.Table(config['table_name'])
response = table.scan()
items = response['Items']

data = [
    [100, 200, 300, 400, 500, 600],
    [150, 250, 350, 450, 550, 650],
    [175, 275, 375, 475, 575, 675],
    [125, 225, 325, 425, 525, 625]
]
blank = pd.MultiIndex.from_product([
    [' '],
    [' '],
    [' ']
])
columns_1 = pd.MultiIndex.from_product([
    [' '],
    [' '],
    ['บริษัทประกันชีวิต']
])
columns_2 = pd.MultiIndex.from_product([
    ['เบี้ยประกันชีวิต'],
    ['ประเภทสามัญ', 'ประเภทอุตสาหกรรม', 'ประเภทกลุ่ม'],
    ['ปีแรก', 'ปีต่อไป']
])
add1 = pd.MultiIndex.from_product([
    ['เบี้ยประกันชีวิต'],
    ['การรับประกันภัย'],
    ['อุบัติเหตุส่วนบุคคล']
])
add2 = pd.MultiIndex.from_product([
    ['ค้าจ้างหรือค่าบำเหน็จในเดือนนี้'],
    ['การรับประกันภัย'],
    ['อุบัติเหตุส่วนบุคคล']
])
columns_3 = pd.MultiIndex.from_product([
    ['ค้าจ้างหรือค่าบำเหน็จในเดือนนี้'],
    ['ประเภทสามัญ', 'ประเภทอุตสาหกรรม', 'ประเภทกลุ่ม'],
    ['ปีแรก', 'ปีต่อไป']
])
columns_4 = pd.MultiIndex.from_product([
    [' '],
    [' '],
    ['ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้','ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน','เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน']
])

joined_columns = blank.append([columns_1,columns_2,add1,columns_3,add2,columns_4])

df = pd.DataFrame(columns=joined_columns)

def highlight_x(v):
    if v in ['เบี้ยประกันชีวิต','ค้าจ้างหรือค่าบำเหน็จในเดือนนี้', 'บริษัทประกันชีวิต', 'ค่าจ้างหรือค่าบำเหน็จรับทั้งสิ้น ในเดือนนี้', 'ค่าจ้างหรือค่าบำเหน็จค้างรับ ณ วันสิ้นเดือน', 'เบี้ยประกันภัยค้างนำส่ง ณ วันสิ้นเดือน']:
        return 'text-align: center; background-color: #888888; color: #ffffff; border: 1px solid black;'
    else :
        return None
def industry_type(v):
        if v in ['ประเภทสามัญ']:
            return 'text-align: center; background-color: #ffcc99; border: 1px solid black;'
        else:
            return None
def test_color(v):
        if v == 'เบี้ยประกันชีวิต':
            return 'text-align: center; background-color: #ffcc99; border: 1px solid black;'
        else:
            return None
        
styled_df = df.style.map_index(highlight_x, axis="columns")


styled_df.to_excel('test.xlsx', engine='openpyxl', index=True)

