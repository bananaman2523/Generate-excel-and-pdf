import json
from datetime import datetime


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


def get_claims_cause_analysis_report(chunk):
    data_transaction = []
    number = 1
    for claim_cause, data in chunk:
        input_data = {
            'number': number,
            'claim_cause': claim_cause,
            'number_of_claims': data['sum'],
            'percentage_total': data['percentage_of_total'],
            'percentage_died': data['types']['died']['percentage'],
            'percentage_permanent': data['types']['permanent_disability']['percentage'],
            'percentage_temporary': data['types']['temporary_disability']['percentage']
        }
        number += 1
        data_transaction.append(input_data)
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


def convert_time(date):
    if date:
        # datetime_str = date
        datetime_obj = datetime.fromisoformat(date)
        formatted_date = datetime_obj.strftime('%Y-%m-%d')
        return formatted_date
