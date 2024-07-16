import requests
import random
import uuid

url = "http://localhost:3000/student/create"
url_update = "http://localhost:3000/student/update"
url_insurance = "http://localhost:3000/life_insurance/create"


headers = {
    "Content-Type": "application/json"
}

list = ['6ea24b47-8b6a-46b7-897d-6a2b1a757105','50053811-a9d5-4397-8f00-f2f4c870d980','132ed3fa-e570-486f-8419-cd0ab244c9e6']
type = ['ILP','ORD','PA','PPI','TKF']
ilp = ['ILP','CI','CL','CP']
ord = ['DS','EW','HB','HS']
pa = ['MB','PA','PAF','PB']
ppi = ['RM','SA','SUPER_CP','TL']
tkf = ['TP','UN','WF','WL']
num = 15
datalist = []
def objective(objective):
    print(objective)
    if objective == 'ILP':
        return 'ประกันชีวิตควบการลงทุน'
    elif objective == 'CI':
        return 'ประกันโรคร้ายแรง'
    elif objective == 'CL':
        return 'ประกันสินเชื่อ'
    elif objective == 'CP':
        return 'ประกันรถยนต์'
    elif objective == 'DS':
        return 'ค่าเลี้ยงดูกรณีทุพพลภาพถาวรสิ้นเชิงจากอุบัติเหตุ'
    elif objective == 'EW':
        return 'ประกันอะไหล่รถยนต์'
    elif objective == 'HB':
        return 'คุ้มครองรายได้จากการรักษาตัวในโรงพยาบาล'
    elif objective == 'HS':
        return 'ประกันสุขภาพ'
    elif objective == 'MB':
        return 'ประกันรถจักรยานยนต์'
    elif objective == 'PA':
        return 'ประกันอุบัติเหตุ'
    elif objective == 'PAF':
        return 'ประกันอุบัติเหตุครอบครัว'
    elif objective == 'PB':
        return 'คุ้มครองผู้ชำระเบี้ย'
    elif objective == 'RM':
        return 'ประกันชีวิตวางแผนเพื่อการเกษียณ'
    elif objective == 'SA':
        return 'ประกันชีวิตเพื่อการออม'
    elif objective == 'SUPER_CP':
        return 'คุ้มครอง super car'
    elif objective == 'TL':
        return 'ประกันชีวิตคุ้มครองชั่วระยะเวลา'
    elif objective == 'TP':
        return 'ประกันการเดินทาง'
    elif objective == 'UN':
        return 'ประกันชีวิตควบการลงทุน'
    elif objective == 'WF':
        return 'ประกันชีวิตเพื่อสังคม'
    elif objective == 'WL':
        return 'ประกันชีวิตคุ้มครองตลอดชีพ'
    else:
        return 'Unknown'
for i in range(len(type)):
    if type[i] == 'ILP':
        for j in range(len(ilp)):
            data_insured = {
                "id" : str(uuid.uuid4()),
                "type_insurance" : type[i],
                "name_insurance": ilp[j],
                "insurance" : objective(ilp[j]),
                "premium": 18000,
                "amount_insured" : 450000,
                'rtu' : 12000
            }
            response = requests.post(url_insurance, json=data_insured, headers=headers) 
    elif type[i] == 'ORD':
        for j in range(len(ord)):
            data_insured = {
                "id" : str(uuid.uuid4()),
                "type_insurance" : type[i],
                "name_insurance": ord[j],
                "insurance" : objective(ord[j]),
                "premium": 20000,
                "amount_insured" : 500000,
                'rtu' : 22000
            }
            response = requests.post(url_insurance, json=data_insured, headers=headers) 
    elif type[i] == 'PA':
        for j in range(len(pa)):
            data_insured = {
                "id" : str(uuid.uuid4()),
                "type_insurance" : type[i],
                "name_insurance": pa[j],
                "insurance" : objective(pa[j]),
                "premium": 25000,
                "amount_insured" : 625000,
                'rtu' : 12000
            }
            response = requests.post(url_insurance, json=data_insured, headers=headers) 
    elif type[i] == 'PPI':
        for j in range(len(ppi)):
            data_insured = {
                "id" : str(uuid.uuid4()),
                "type_insurance" : type[i],
                "name_insurance": ppi[j],
                "insurance" : objective(ppi[j]),
                "premium": 28000,
                "amount_insured" : 700000,
                'rtu' : 32000
            }
            response = requests.post(url_insurance, json=data_insured, headers=headers) 
    elif type[i] == 'TKF':
        for j in range(len(tkf)):
            data_insured = {
                "id" : str(uuid.uuid4()),
                "type_insurance" : type[i],
                "name_insurance": tkf[j],
                "insurance" : objective(tkf[j]),
                "premium": 18500,
                "amount_insured" : 463000,
                'rtu' : 25000
            }
            response = requests.post(url_insurance, json=data_insured, headers=headers) 
    else:
        name_insurance = 'Unknown'
data_insured = {
    "id" : str(uuid.uuid4()),
    "type_insurance" : 'ILP',
    "name_insurance": 'ILP',
    "insurance" : 'ประกันชีวิตควบการลงทุน',
    "amount_insured" : 450000
}
response = requests.post(url_insurance, json=data_insured, headers=headers) 
 

print('create success')
# for i in range(1, num + 1):
#     data = {
#         "id" : f"{i}",
#         "name": f"test{i}",
#         "lastname": f"test{i}",
#         "year" : random.randint(1,4),
#         "score": random.randint(0, 100),
#         "field": "dii"
#     }
#     response = requests.post(url, json=data, headers=headers)
#     datalist.append(data)
    
# print('create success')
# for i in range(len(datalist)):
#     data_uni = {
#         "id": datalist[i]['id'],
#         "mode_update": "university",
#         "new_value": random.choice(list)
#     }

#     response = requests.patch(url_update, json=data_uni, headers=headers)
# print('update success')
