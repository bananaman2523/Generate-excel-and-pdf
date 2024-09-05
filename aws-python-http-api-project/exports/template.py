def template_pdf(file_name):
    if file_name == 'RPCL001':
        template_input = {
            'file_name' : 'RPCL001',
            'language' : 'th',
            'table' : 'sales_premium_transaction',
            'user_name' : 'test test',
            'set_column' : 11,
            'cell_widths' : [10, 23, 23, 23, 23, 23, 23, 23, 23],
            'row_height' : 5,
            'columns_styles' : [
                {'columns': 'id', 'style': {'font_size': 10 , 'bold' : 'B', 'color' : (230, 247, 255), 'bg_color' : (230, 247, 255)}},
                {'columns': 'customer_name', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'loan_id', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'insurance_company', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'claim_amount', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'claim_status', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'submission_date', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'approval_date', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
                {'columns': 'payout_amount', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (230, 247, 255)}},
            ],
            'header_mapping' : {
                'id' : 'ID\n ',
                'customer_name': 'Customer Name\n ',
                'loan_id': 'Loan ID\n ',
                'insurance_company': 'Insurance Company',
                'claim_amount': 'Claim Amount\n ',
                'claim_status': 'Claim Status\n ',
                'submission_date': 'Submission Date\n ',
                'approval_date': 'Approval Date\n ',
                'payout_amount': 'Payout Amount\n '
            },
            'rows_styles' : [
                {'rows': 'id', 'style': {'font_size': 10, 'align': 'C'}},
                {'rows': 'customer_name', 'style': {'font_size': 10}},
                {'rows': 'loan_id', 'style': {'font_size': 10}},
                {'rows': 'insurance_company', 'style': {'font_size': 10}},
                {'rows': 'claim_amount', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'claim_status', 'style': {'font_size': 10}},
                {'rows': 'submission_date', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'approval_date', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'payout_amount', 'style': {'font_size': 10, 'align': 'R'}},
            ],
            'header_template': [
                {'header': 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)', 'style' : {'font_size' : 16}},
                {'header': 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', 'style': {'font_size' : 10}}
            ]
        }
        return template_input
    elif file_name == 'RPCL002':
        template_input = {
            'file_name' : 'RPCL002',
            'language' : 'th',
            'table' : 'sales_premium_transaction',
            'user_name' : 'test test',
            'set_column' : 11,
            'cell_widths' : [9, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18],
            'row_height' : 5,
            'columns_styles' : [
                {'columns': 'id', 'style': {'font_size': 10 , 'bold' : 'B', 'bg_color' : (255, 255, 255)}},
                {'columns': 'insurance_company', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 255)}},
                {'columns': 'number_of_claims', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 255)}},
                {'columns': 'amount_of_claim', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 255)}},
                {'columns': 'approved_claims', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (0, 255, 0)}},
                {'columns': 'approved_died', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (0, 255, 0)}},
                {'columns': 'approved_permanent_disability', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (0, 255, 0)}},
                {'columns': 'approved_temporary_disability', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (0, 255, 0)}},
                {'columns': 'denied_claims', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (244, 204, 204)}},
                {'columns': 'denied_died', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (244, 204, 204)}},
                {'columns': 'denied_permanent_disability', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (244, 204, 204)}},
                {'columns': 'denied_temporary_disability', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (244, 204, 204)}},
                {'columns': 'pending_claims', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 0)}},
                {'columns': 'pending_died', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 0)}},
                {'columns': 'pending_permanent_disability', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 0)}},
                {'columns': 'pending_temporary_disability', 'style': {'font_size': 10, 'bold' : 'B', 'bg_color' : (255, 255, 0)}},
            ],
            'header_mapping' : {
                'id' : 'ID\n',
                'insurance_company': 'Insurance Company',
                'number_of_claims': 'Number of Claims',
                'amount_of_claim': 'Amount of Claim',
                'approved_claims': 'Approved Claims\n ',
                'approved_died': 'Amount of Died\n ',
                'approved_permanent_disability': 'Amount of Permanent Disability',
                'approved_temporary_disability': 'Amount of Temporary Disability',
                'denied_claims': 'Denied Claims\n ',
                'denied_died': 'Amount of Died\n ',
                'denied_permanent_disability': 'Amount of Permanent Disability',
                'denied_temporary_disability': 'Amount of Temporary Disability',
                'pending_claims': 'Pending Claims\n ',
                'pending_died': 'Amount of Died\n ',
                'pending_permanent_disability': 'Amount of Permanent Disability',
                'pending_temporary_disability': 'Amount of Temporary Disability'
            },
            'rows_styles' : [
                {'rows': 'id', 'style': {'font_size': 10 , 'bg_color' : (255, 255, 255)}},
                {'rows': 'insurance_company', 'style': {'font_size': 10, 'bg_color' : (255, 255, 255)}},
                {'rows': 'number_of_claims', 'style': {'font_size': 10, 'bg_color' : (255, 255, 255)}},
                {'rows': 'amount_of_claim', 'style': {'font_size': 10, 'bg_color' : (255, 255, 255)}},
                {'rows': 'approved_claims', 'style': {'font_size': 10, 'bg_color' : (0, 255, 0)}},
                {'rows': 'approved_died', 'style': {'font_size': 10, 'bg_color' : (0, 255, 0)}},
                {'rows': 'approved_permanent_disability', 'style': {'font_size': 10, 'bg_color' : (0, 255, 0)}},
                {'rows': 'approved_temporary_disability', 'style': {'font_size': 10, 'bg_color' : (0, 255, 0)}},
                {'rows': 'denied_claims', 'style': {'font_size': 10, 'bg_color' : (244, 204, 204)}},
                {'rows': 'denied_died', 'style': {'font_size': 10, 'bg_color' : (244, 204, 204)}},
                {'rows': 'denied_permanent_disability', 'style': {'font_size': 10, 'bg_color' : (244, 204, 204)}},
                {'rows': 'denied_temporary_disability', 'style': {'font_size': 10, 'bg_color' : (244, 204, 204)}},
                {'rows': 'pending_claims', 'style': {'font_size': 10, 'bg_color' : (255, 255, 0)}},
                {'rows': 'pending_died', 'style': {'font_size': 10, 'bg_color' : (255, 255, 0)}},
                {'rows': 'pending_permanent_disability', 'style': {'font_size': 10, 'bg_color' : (255, 255, 0)}},
                {'rows': 'pending_temporary_disability', 'style': {'font_size': 10, 'bg_color' : (255, 255, 0)}},
            ],
            'header_template': [
                {'header': 'รายงานการเรียกร้องตามลูกค้า (Claims by Customer Report)', 'style' : {'font_size' : 16}},
                {'header': 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', 'style': {'font_size' : 10}}
            ]
        }
        return template_input
    elif file_name == 'RPCL003':
        template_input = {
            'file_name' : 'RPCL003',
            'language' : 'th',
            'table' : 'sales_premium_transaction',
            'user_name' : 'test test',
            'set_column' : 11,
            'cell_widths' : [10, 30, 30, 30, 30, 30, 30],
            'row_height' : 5,
            'columns_styles' : [
                {'columns': 'id', 'style': {'font_size': 10 , 'bold' : 'B'}},
                {'columns': 'claim_cause', 'style': {'font_size': 10, 'bold' : 'B'}},
                {'columns': 'number_of_claims', 'style': {'font_size': 10, 'bold' : 'B'}},
                {'columns': 'percentage_total', 'style': {'font_size': 10, 'bold' : 'B'}},
                {'columns': 'percentage_died', 'style': {'font_size': 10, 'bold' : 'B'}},
                {'columns': 'percentage_permanent', 'style': {'font_size': 10, 'bold' : 'B'}},
                {'columns': 'percentage_temporary', 'style': {'font_size': 10, 'bold' : 'B'}}
            ],
            'header_mapping' : {
                'id' : 'ID\n ',
                'claim_cause': 'Claim Cause\n ',
                'number_of_claims': 'Number of Claims',
                'percentage_total': 'Percentage of Total Claims (%)',
                'percentage_died': 'Percentage of Died (%)',
                'percentage_permanent': 'Percentage of Permanent Disability (%)',
                'percentage_temporary': 'Percentage of Temporary Disability (%)'
            },
            'rows_styles' : [
                {'rows': 'id', 'style': {'font_size': 10, 'align': 'C'}},
                {'rows': 'claim_cause', 'style': {'font_size': 10}},
                {'rows': 'number_of_claims', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'percentage_total', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'percentage_died', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'percentage_permanent', 'style': {'font_size': 10, 'align': 'R'}},
                {'rows': 'percentage_temporary', 'style': {'font_size': 10, 'align': 'R'}}
            ],
            'header_template': [
                {'header': 'รายงานการเรียกร้องตามบริษัทประกันภัย (Claims by Insurance Company Report)', 'style' : {'font_size' : 16}},
                {'header': 'ข้อมูลระหว่างวันที่ 01/07/2567 - 20/07/2567', 'style': {'font_size' : 10}}
            ]
        }
        return template_input

def template_xlsx(file_name, filter_items):
    if file_name == 'RPCL001':
        template_input = {
            'file_name' : 'RPCL001',
            'data_per_page' : 1000,
            'filter' : filter_items,
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
                'style_number': {'font_size': 9, 'num_format': '#,##0.00', 'border': 1,'text_wrap' : True, 'valign' : 'top', 'align' : 'right'},
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
        return template_input
    elif file_name == 'RPCL002':
        template_input = {
            'file_name' : 'RPCL002',
            'data_per_page' : 1000,
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
        return template_input
    elif file_name == 'RPCL003':
        template_input = {
            'file_name' : 'RPCL003',
            'data_per_page' : 1000,
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
        return template_input