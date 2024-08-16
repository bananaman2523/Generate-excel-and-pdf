import json
import numpy as np
import pandas as pd
from collections import Counter
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

        total_fee_amount_str = item.get("received_premium", '0')
        total_fee_amount = float(total_fee_amount_str) if total_fee_amount_str else 0

        total_fee_rate_str = item.get("received_premium_variances", '0')
        cleaned_string = total_fee_rate_str.strip("'")
        total_fee_rate = float(cleaned_string) if cleaned_string else 0
        if total_fee_rate < 0:
            payout_amount = f'({abs(total_fee_rate)})'
        else:
            payout_amount = total_fee_rate
        input_data = {
            'number' : number,
            'customer_name': item.get("customer_full_name", ''),
            'loan_id': loan_product_name_th,
            'insurance_company': insurance_company_name_th,
            'claim_amount': total_fee_amount,
            'claim_status': activity_status,
            'submission_date': convert_time(sub_date, language),
            'approval_date': convert_time(approval_date, language),
            'payout_amount': payout_amount
        }
        data_transaction.append(input_data)
        number += 1
    return data_transaction

def get_value_counts(df, column_name, new_column_name):
    value_counts_df = df[column_name].value_counts().reset_index()
    value_counts_df.columns = [f'{column_name}_data', new_column_name]
    return value_counts_df

def get_status_counts(df):
    insurance_status_counts = df.groupby(['insurance_company', 'claim_status']).size().reset_index(name='insurance_count')
    insurance_status_counts.columns = ['name_insurance', 'insurance_status', 'insurance_count']
    
    return insurance_status_counts

def extract_month_year(date_str, language = 'th'):
    if date_str:
        if language == 'th':
            thai_months = {
                "ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
                "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
                "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12
            }
            try:
                day, thai_month, thai_year = date_str.split()
                gregorian_year = int(thai_year) + 2500 - 543
                month = thai_months.get(thai_month, None)
                if month:
                    return thai_month
                else:
                    return None
                    
            except ValueError:
                return None
        else:
            try:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                return date_obj.strftime('%b')
            except ValueError:
                return None
    return None

def generate_month_range(language = 'th'):
    if language == 'th':
        return [
            "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
            "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
        ]
    else:
        return [datetime(2024, month, 1).strftime('%B') for month in range(1, 13)]

def has_no_parentheses(value):
    value_str = str(value)
    return '(' not in value_str and ')' not in value_str

def remove_parentheses(value: str) -> str:
    return float(value.strip('()'))

def get_approve_date(df, language = 'th'):
    approval_dates = df['approval_date'].dropna().tolist()
    month_years = [extract_month_year(date, language) for date in approval_dates if extract_month_year(date, language)]

    counts = Counter(month_years)
    all_months = generate_month_range(language)
    result = {month: counts.get(month, 0) for month in all_months}

    out_balance = {month: 0 for month in all_months}
    difference = {month: 0 for month in all_months}
    paid = {month: 0 for month in all_months}

    for _, row in df.iterrows():
        month_year = extract_month_year(row['approval_date'], language)
        if month_year in all_months:
            claim_amount = row['claim_amount']
            payout_amount = row['payout_amount']
            if payout_amount == 0:
                paid[month_year] += claim_amount
            elif claim_amount != payout_amount and has_no_parentheses(payout_amount) and payout_amount != 0:
                difference[month_year] += payout_amount
            elif isinstance(payout_amount, str):
                out_balance[month_year] += remove_parentheses(payout_amount)

    df_approve_date = pd.DataFrame({
        'month': all_months,
        'approve_date': [result[month] for month in all_months],
        'out_balance': [round(out_balance[month], 2) for month in all_months],
        'difference': [round(difference[month], 2) for month in all_months],
        'paid': [round(paid[month], 2) for month in all_months]
    })

    sum_data = pd.DataFrame({
        'categories_approve': ['ยอดค้างชำระ', 'ยอดเกินรวม', 'จำนวนเงินทั้งหมด'],
        'value_approve': [
            df_approve_date['out_balance'].sum(),
            df_approve_date['difference'].sum(),
            df_approve_date['paid'].sum()
        ]
    })

    pd_dataframe = pd.concat([df_approve_date, sum_data], axis=1)
    return pd_dataframe

def get_customer_report_data(data, language='th'):
    data_transaction = []
    for i, item in enumerate(data):
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

        total_fee_amount_str = item.get("received_premium", '0')
        total_fee_amount = float(total_fee_amount_str) if total_fee_amount_str else 0

        total_fee_rate_str = item.get("received_premium_variances", '0')
        cleaned_string = total_fee_rate_str.strip("'")
        total_fee_rate = float(cleaned_string) if cleaned_string else 0
        if total_fee_rate < 0:
            payout_amount = f'({abs(total_fee_rate)})'
        else:
            payout_amount = total_fee_rate
        input_data = {
            'customer_name': item.get("customer_full_name", ''),
            'loan_id': loan_product_name_th,
            'insurance_company': insurance_company_name_th,
            'claim_amount': total_fee_amount,
            'claim_status': activity_status,
            'submission_date': convert_time(sub_date, language),
            'approval_date': convert_time(approval_date, language),
            'payout_amount': payout_amount
        }
        data_transaction.append(input_data)

    df = pd.DataFrame(data_transaction)
    insurance_company_counts = get_value_counts(df, 'insurance_company', 'insurance_company_count')
    loan_product_counts = get_value_counts(df, 'loan_id', 'loan_product_count')
    claim_status_counts = get_value_counts(df, 'claim_status', 'claim_status_count')
    status_count_each = get_status_counts(df)

    df_approve_date = get_approve_date(df, language)
    # pd_dataframe = pd.concat([df, insurance_company_counts, loan_product_counts, claim_status_counts, approveal_date_counts, status_count_each], axis=1)
    pd_dataframe = concatenate_dataframes(df, insurance_company_counts, loan_product_counts, claim_status_counts, status_count_each, df_approve_date)

    return pd_dataframe

def concatenate_dataframes(*dfs):
    return pd.concat(dfs, axis=1)

def convert_time(date, language = 'th'):
    if date:
        if language == 'en':
            datetime_obj = datetime.fromisoformat(date)
            formatted_date = datetime_obj.strftime('%d %B %Y')
            return formatted_date
        else:
            datetime_obj = datetime.fromisoformat(date)
            thai_year = datetime_obj.year + 543
            thai_year_short = str(thai_year)[-2:]
            thai_months_short = [
                "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
            ]
            thai_month = thai_months_short[datetime_obj.month - 1]
            formatted_date = datetime_obj.strftime(f'{datetime_obj.day} {thai_month} {thai_year_short}')
            return formatted_date