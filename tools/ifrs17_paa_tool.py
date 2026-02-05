"""
Creates a comprehensive IFRS 17 PAA Excel tool with formulas and functionality
"""

import pandas as pd
from datetime import datetime, timedelta
import xlsxwriter

def create_comprehensive_ifrs17_excel():
    """
    Create a full-featured IFRS 17 PAA Excel workbook
    """
    
    filename = "IFRS17_PAA_Complete_Tool.xlsx"
    
    # Create Excel writer
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book
    
    # ========================================================================
    # FORMATS
    # ========================================================================
    
    # Header format
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })
    
    # Title format
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'bg_color': '#203864',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter'
    })
    
    # Subtitle format
    subtitle_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'bg_color': '#305496',
        'font_color': 'white',
        'align': 'left'
    })
    
    # Currency format
    currency_format = workbook.add_format({
        'num_format': '$#,##0.00',
        'border': 1
    })
    
    # Percentage format
    percent_format = workbook.add_format({
        'num_format': '0.00%',
        'border': 1
    })
    
    # Date format
    date_format = workbook.add_format({
        'num_format': 'yyyy-mm-dd',
        'border': 1
    })
    
    # Number format
    number_format = workbook.add_format({
        'num_format': '#,##0',
        'border': 1
    })
    
    # Calculation format (light blue background)
    calc_format = workbook.add_format({
        'num_format': '$#,##0.00',
        'bg_color': '#E7E6E6',
        'border': 1
    })
    
    # Warning format (red)
    warning_format = workbook.add_format({
        'bg_color': '#FFC7CE',
        'font_color': '#9C0006',
        'border': 1
    })
    
    # Success format (green)
    success_format = workbook.add_format({
        'bg_color': '#C6EFCE',
        'font_color': '#006100',
        'border': 1
    })
    
    # ========================================================================
    # SHEET 1: INSTRUCTIONS
    # ========================================================================
    
    instructions_data = {
        'Step': [1, 2, 3, 4, 5, 6, 7],
        'Action': [
            'Start Here: Read all instructions',
            'Configure Settings: Go to "Config" sheet and set your parameters',
            'Enter Contract Data: Fill in contract details in "Contract Input" sheet',
            'Review PAA Eligibility: Check "PAA Eligibility Test" sheet',
            'Review Onerous Tests: Check "Onerous Test" sheet',
            'View Reports: See "Liability Report", "Income Statement", and "Roll Forward"',
            'Dashboard: View summary metrics in "Dashboard" sheet'
        ],
        'Description': [
            'This tool calculates IFRS 17 PAA measurements automatically',
            'Set discount rates, risk adjustment %, and other assumptions',
            'Enter all contract details - formulas will calculate automatically',
            'Automated test shows which contracts qualify for PAA',
            'Identifies onerous contracts and calculates loss components',
            'Comprehensive financial reports with full calculations',
            'Visual summary of key metrics and portfolio overview'
        ]
    }
    
    df_instructions = pd.DataFrame(instructions_data)
    df_instructions.to_excel(writer, sheet_name='Instructions', index=False, startrow=3)
    
    worksheet_instructions = writer.sheets['Instructions']
    worksheet_instructions.merge_range('A1:C1', 'IFRS 17 PAA ENGINE - USER GUIDE', title_format)
    worksheet_instructions.set_column('A:A', 8)
    worksheet_instructions.set_column('B:B', 50)
    worksheet_instructions.set_column('C:C', 60)
    
    # Add instructions header
    for col_num, value in enumerate(df_instructions.columns.values):
        worksheet_instructions.write(3, col_num, value, header_format)
    
    # Key features section
    worksheet_instructions.merge_range('A12:C12', 'KEY FEATURES', subtitle_format)
    features = [
        ['✓', 'PAA Eligibility Testing', 'Automatic test for coverage period and materiality'],
        ['✓', 'Onerous Contract Test', 'Identifies loss-making contracts'],
        ['✓', 'Liability Measurement', 'LRC, LIC, and Loss Component calculations'],
        ['✓', 'Premium Earning', 'Straight-line amortization over coverage period'],
        ['✓', 'Roll-forward', 'Period-over-period reconciliation'],
        ['✓', 'Income Statement', 'Insurance revenue and service expenses'],
        ['✓', 'Dashboard', 'Visual summary of key metrics']
    ]
    
    for i, feature in enumerate(features):
        for j, val in enumerate(feature):
            worksheet_instructions.write(13 + i, j, val)
    
    # ========================================================================
    # SHEET 2: CONFIGURATION
    # ========================================================================
    
    config_data = {
        'Parameter': [
            'Reporting Currency',
            'Reporting Date',
            'Default Discount Rate',
            'Risk Adjustment Method',
            'Risk Adjustment %',
            'Coverage Period Threshold (months)',
            'Materiality Threshold %',
            'Acquisition Cost Treatment',
            'Premium Earning Method',
            'Fiscal Year End'
        ],
        'Value': [
            'USD',
            '2025-12-31',
            '5.0%',
            'Percentage',
            '10.0%',
            12,
            '5.0%',
            'Defer',
            'Straight-line',
            '12-31'
        ],
        'Description': [
            'Currency for all monetary values',
            'Current reporting date for calculations',
            'Annual discount rate for present value calculations',
            'Method for calculating risk adjustment',
            'Risk adjustment as % of expected claims',
            'Maximum coverage period for automatic PAA eligibility',
            'Threshold for discount impact materiality test',
            'Defer acquisition costs or expense immediately',
            'Method for recognizing premium revenue',
            'Company fiscal year end date'
        ]
    }
    
    df_config = pd.DataFrame(config_data)
    df_config.to_excel(writer, sheet_name='Config', index=False, startrow=2)
    
    worksheet_config = writer.sheets['Config']
    worksheet_config.merge_range('A1:C1', 'CONFIGURATION & ASSUMPTIONS', title_format)
    worksheet_config.set_column('A:A', 35)
    worksheet_config.set_column('B:B', 20)
    worksheet_config.set_column('C:C', 60)
    
    for col_num, value in enumerate(df_config.columns.values):
        worksheet_config.write(2, col_num, value, header_format)
    
    # ========================================================================
    # SHEET 3: CONTRACT INPUT
    # ========================================================================
    
    # Sample contract data
    contract_data = {
        'Contract_ID': ['POL-2025-001', 'POL-2025-002', 'POL-2025-003', 'POL-2025-004', 'POL-2025-005'],
        'Policy_Holder': ['ABC Corporation', 'XYZ Limited', 'High Risk Inc', 'Safe Drivers Co', 'Premium Properties'],
        'Contract_Type': ['Property Insurance', 'Auto Insurance', 'Liability Insurance', 'Auto Insurance', 'Property Insurance'],
        'Inception_Date': ['2025-01-01', '2025-02-01', '2025-01-15', '2025-03-01', '2025-01-10'],
        'Coverage_Period_Months': [12, 6, 12, 12, 24],
        'Premium': [50000, 25000, 40000, 30000, 100000],
        'Expected_Claims': [35000, 18000, 45000, 21000, 70000],
        'Acquisition_Costs': [5000, 2500, 4000, 3000, 10000],
        'Discount_Rate': [0.05, 0.05, 0.05, 0.05, 0.05],
        'Risk_Adjustment': [3500, 1800, 4500, 2100, 7000]
    }
    
    df_contracts = pd.DataFrame(contract_data)
    df_contracts.to_excel(writer, sheet_name='Contract Input', index=False, startrow=2)
    
    worksheet_contracts = writer.sheets['Contract Input']
    worksheet_contracts.merge_range('A1:J1', 'CONTRACT DATA INPUT', title_format)
    worksheet_contracts.set_column('A:A', 15)
    worksheet_contracts.set_column('B:B', 20)
    worksheet_contracts.set_column('C:C', 20)
    worksheet_contracts.set_column('D:D', 15)
    worksheet_contracts.set_column('E:E', 20)
    worksheet_contracts.set_column('F:J', 18)
    
    for col_num, value in enumerate(df_contracts.columns.values):
        worksheet_contracts.write(2, col_num, value, header_format)
    
    # Format the data
    for row in range(len(df_contracts)):
        worksheet_contracts.write(3 + row, 3, df_contracts.iloc[row]['Inception_Date'], date_format)
        worksheet_contracts.write(3 + row, 4, df_contracts.iloc[row]['Coverage_Period_Months'], number_format)
        worksheet_contracts.write(3 + row, 5, df_contracts.iloc[row]['Premium'], currency_format)
        worksheet_contracts.write(3 + row, 6, df_contracts.iloc[row]['Expected_Claims'], currency_format)
        worksheet_contracts.write(3 + row, 7, df_contracts.iloc[row]['Acquisition_Costs'], currency_format)
        worksheet_contracts.write(3 + row, 8, df_contracts.iloc[row]['Discount_Rate'], percent_format)
        worksheet_contracts.write(3 + row, 9, df_contracts.iloc[row]['Risk_Adjustment'], currency_format)
    
    # ========================================================================
    # SHEET 4: PAA ELIGIBILITY TEST
    # ========================================================================
    
    worksheet_paa = workbook.add_worksheet('PAA Eligibility Test')
    worksheet_paa.merge_range('A1:G1', 'PAA ELIGIBILITY TESTING', title_format)
    
    # Headers
    paa_headers = ['Contract_ID', 'Coverage_Period_Months', 'Coverage_Test', 'Discount_Impact', 'Materiality_Test', 'PAA_Eligible', 'Recommended_Model']
    for col, header in enumerate(paa_headers):
        worksheet_paa.write(2, col, header, header_format)
    
    worksheet_paa.set_column('A:A', 15)
    worksheet_paa.set_column('B:B', 20)
    worksheet_paa.set_column('C:G', 18)
    
    # Formulas for PAA eligibility
    for row in range(5):
        # Contract ID
        worksheet_paa.write_formula(3 + row, 0, f"='Contract Input'!A{4+row}")
        
        # Coverage Period
        worksheet_paa.write_formula(3 + row, 1, f"='Contract Input'!E{4+row}", number_format)
        
        # Coverage Test (Pass if <= 12 months)
        worksheet_paa.write_formula(3 + row, 2, f'=IF(B{4+row}<=12,"PASS","FAIL")')
        
        # Discount Impact
        worksheet_paa.write_formula(3 + row, 3, 
            f"='Contract Input'!G{4+row}*'Contract Input'!I{4+row}*('Contract Input'!E{4+row}/12)", 
            currency_format)
        
        # Materiality Test (Pass if discount impact < 5% of premium)
        worksheet_paa.write_formula(3 + row, 4, 
            f'=IF(D{4+row}<\'Contract Input\'!F{4+row}*0.05,"PASS","FAIL")')
        
        # PAA Eligible (Pass if either test passes)
        worksheet_paa.write_formula(3 + row, 5, 
            f'=IF(OR(C{4+row}="PASS",E{4+row}="PASS"),"YES","NO")')
        
        # Recommended Model
        worksheet_paa.write_formula(3 + row, 6, 
            f'=IF(F{4+row}="YES","PAA","GMM")')
    
    # Add conditional formatting
    worksheet_paa.conditional_format('F4:F8', {
        'type': 'text',
        'criteria': 'containing',
        'value': 'YES',
        'format': success_format
    })
    
    worksheet_paa.conditional_format('F4:F8', {
        'type': 'text',
        'criteria': 'containing',
        'value': 'NO',
        'format': warning_format
    })
    
    # ========================================================================
    # SHEET 5: ONEROUS CONTRACT TEST
    # ========================================================================
    
    worksheet_onerous = workbook.add_worksheet('Onerous Test')
    worksheet_onerous.merge_range('A1:H1', 'ONEROUS CONTRACT TEST', title_format)
    
    # Headers
    onerous_headers = ['Contract_ID', 'Premium', 'Acquisition_Costs', 'Net_Premium', 
                       'Expected_Claims', 'Risk_Adjustment', 'Fulfilment_CF', 'Shortfall', 'Is_Onerous']
    for col, header in enumerate(onerous_headers):
        worksheet_onerous.write(2, col, header, header_format)
    
    worksheet_onerous.set_column('A:A', 15)
    worksheet_onerous.set_column('B:I', 18)
    
    # Formulas for onerous test
    for row in range(5):
        # Contract ID
        worksheet_onerous.write_formula(3 + row, 0, f"='Contract Input'!A{4+row}")
        
        # Premium
        worksheet_onerous.write_formula(3 + row, 1, f"='Contract Input'!F{4+row}", currency_format)
        
        # Acquisition Costs
        worksheet_onerous.write_formula(3 + row, 2, f"='Contract Input'!H{4+row}", currency_format)
        
        # Net Premium
        worksheet_onerous.write_formula(3 + row, 3, f"=B{4+row}-C{4+row}", currency_format)
        
        # Expected Claims
        worksheet_onerous.write_formula(3 + row, 4, f"='Contract Input'!G{4+row}", currency_format)
        
        # Risk Adjustment
        worksheet_onerous.write_formula(3 + row, 5, f"='Contract Input'!J{4+row}", currency_format)
        
        # Fulfilment Cash Flows
        worksheet_onerous.write_formula(3 + row, 6, f"=E{4+row}+F{4+row}", currency_format)
        
        # Shortfall (Loss Component)
        worksheet_onerous.write_formula(3 + row, 7, f"=MAX(G{4+row}-D{4+row},0)", currency_format)
        
        # Is Onerous
        worksheet_onerous.write_formula(3 + row, 8, f'=IF(H{4+row}>0,"ONEROUS","NOT ONEROUS")')
    
    # Conditional formatting
    worksheet_onerous.conditional_format('I4:I8', {
        'type': 'text',
        'criteria': 'containing',
        'value': 'ONEROUS',
        'format': warning_format
    })
    
    worksheet_onerous.conditional_format('I4:I8', {
        'type': 'text',
        'criteria': 'containing',
        'value': 'NOT ONEROUS',
        'format': success_format
    })
    
    # ========================================================================
    # SHEET 6: LIABILITY REPORT
    # ========================================================================
    
    worksheet_liability = workbook.add_worksheet('Liability Report')
    worksheet_liability.merge_range('A1:M1', 'LIABILITY MEASUREMENT REPORT', title_format)
    
    # Add reporting date reference
    worksheet_liability.write('A2', 'Reporting Date:', subtitle_format)
    worksheet_liability.write_formula('B2', "=Config!B3", date_format)
    
    # Headers
    liability_headers = ['Contract_ID', 'Inception_Date', 'Coverage_Period_Days', 'Days_Elapsed', 
                        'Premium', 'Earned_Premium', 'Unearned_Premium', 'LRC', 
                        'LIC', 'Loss_Component', 'Total_Liability']
    for col, header in enumerate(liability_headers):
        worksheet_liability.write(3, col, header, header_format)
    
    worksheet_liability.set_column('A:A', 15)
    worksheet_liability.set_column('B:K', 18)
    
    # Formulas for liability calculations
    for row in range(5):
        # Contract ID
        worksheet_liability.write_formula(4 + row, 0, f"='Contract Input'!A{4+row}")
        
        # Inception Date
        worksheet_liability.write_formula(4 + row, 1, f"='Contract Input'!D{4+row}", date_format)
        
        # Coverage Period Days
        worksheet_liability.write_formula(4 + row, 2, f"='Contract Input'!E{4+row}*30", number_format)
        
        # Days Elapsed
        worksheet_liability.write_formula(4 + row, 3, 
            f"=MIN(MAX($B$2-B{5+row},0),C{5+row})", number_format)
        
        # Premium
        worksheet_liability.write_formula(4 + row, 4, f"='Contract Input'!F{4+row}", currency_format)
        
        # Earned Premium (straight-line)
        worksheet_liability.write_formula(4 + row, 5, 
            f"=IF(D{5+row}=0,0,E{5+row}*D{5+row}/C{5+row})", currency_format)
        
        # Unearned Premium
        worksheet_liability.write_formula(4 + row, 6, f"=E{5+row}-F{5+row}", currency_format)
        
        # LRC (Liability for Remaining Coverage)
        worksheet_liability.write_formula(4 + row, 7, f"=G{5+row}+J{5+row}", currency_format)
        
        # LIC (Liability for Incurred Claims) - would come from actual claims data
        worksheet_liability.write(4 + row, 8, 0, currency_format)
        
        # Loss Component
        worksheet_liability.write_formula(4 + row, 9, f"='Onerous Test'!H{4+row}", currency_format)
        
        # Total Liability
        worksheet_liability.write_formula(4 + row, 10, f"=H{5+row}+I{5+row}", currency_format)
    
    # Add totals row
    worksheet_liability.write(9, 0, 'TOTAL', header_format)
    for col in [4, 5, 6, 7, 8, 9, 10]:
        worksheet_liability.write_formula(9, col, f"=SUM({chr(65+col)}5:{chr(65+col)}9)", calc_format)
    
    # ========================================================================
    # SHEET 7: INCOME STATEMENT
    # ========================================================================
    
    worksheet_income = workbook.add_worksheet('Income Statement')
    worksheet_income.merge_range('A1:G1', 'INSURANCE SERVICE RESULT', title_format)
    
    # Add period reference
    worksheet_income.write('A2', 'Period Ending:', subtitle_format)
    worksheet_income.write_formula('B2', "=Config!B3", date_format)
    
    # Headers
    income_headers = ['Contract_ID', 'Insurance_Revenue', 'Claims_Incurred', 
                     'Acquisition_Costs', 'Total_Expenses', 'Insurance_Service_Result', 'Margin_%']
    for col, header in enumerate(income_headers):
        worksheet_income.write(3, col, header, header_format)
    
    worksheet_income.set_column('A:A', 15)
    worksheet_income.set_column('B:G', 20)
    
    # Formulas for income statement
    for row in range(5):
        # Contract ID
        worksheet_income.write_formula(4 + row, 0, f"='Contract Input'!A{4+row}")
        
        # Insurance Revenue (Earned Premium)
        worksheet_income.write_formula(4 + row, 1, f"='Liability Report'!F{5+row}", currency_format)
        
        # Claims Incurred (would be actual data - using expected for demo)
        worksheet_income.write_formula(4 + row, 2, 
            f"='Contract Input'!G{4+row}*'Liability Report'!D{5+row}/'Liability Report'!C{5+row}", 
            currency_format)
        
        # Acquisition Costs (amortized like premium)
        worksheet_income.write_formula(4 + row, 3, 
            f"='Contract Input'!H{4+row}*'Liability Report'!D{5+row}/'Liability Report'!C{5+row}", 
            currency_format)
        
        # Total Expenses
        worksheet_income.write_formula(4 + row, 4, f"=C{5+row}+D{5+row}", currency_format)
        
        # Insurance Service Result
        worksheet_income.write_formula(4 + row, 5, f"=B{5+row}-E{5+row}", currency_format)
        
        # Margin %
        worksheet_income.write_formula(4 + row, 6, 
            f"=IF(B{5+row}=0,0,F{5+row}/B{5+row})", percent_format)
    
    # Add totals
    worksheet_income.write(9, 0, 'TOTAL', header_format)
    for col in [1, 2, 3, 4, 5]:
        worksheet_income.write_formula(9, col, f"=SUM({chr(65+col)}5:{chr(65+col)}9)", calc_format)
    worksheet_income.write_formula(9, 6, "=IF(B10=0,0,F10/B10)", calc_format)
    
    # Conditional formatting for service result
    worksheet_income.conditional_format('F5:F9', {
        'type': 'cell',
        'criteria': '<',
        'value': 0,
        'format': warning_format
    })
    
    worksheet_income.conditional_format('F5:F9', {
        'type': 'cell',
        'criteria': '>=',
        'value': 0,
        'format': success_format
    })
    
    # ========================================================================
    # SHEET 8: ROLL FORWARD
    # ========================================================================
    
    worksheet_rollforward = workbook.add_worksheet('Roll Forward')
    worksheet_rollforward.merge_range('A1:H1', 'LIABILITY ROLL-FORWARD RECONCILIATION', title_format)
    
    # Headers
    rollforward_headers = ['Contract_ID', 'Opening_LRC', 'Premium_Received', 'Premium_Earned', 
                          'Loss_Component', 'Closing_LRC', 'Opening_LIC', 'Claims_Paid', 'Closing_LIC']
    for col, header in enumerate(rollforward_headers):
        worksheet_rollforward.write(2, col, header, header_format)
    
    worksheet_rollforward.set_column('A:A', 15)
    worksheet_rollforward.set_column('B:I', 18)
    
    # Formulas for roll-forward
    for row in range(5):
        # Contract ID
        worksheet_rollforward.write_formula(3 + row, 0, f"='Contract Input'!A{4+row}")
        
        # Opening LRC (assume zero at inception for demo)
        worksheet_rollforward.write(3 + row, 1, 0, currency_format)
        
        # Premium Received
        worksheet_rollforward.write_formula(3 + row, 2, f"='Contract Input'!F{4+row}", currency_format)
        
        # Premium Earned (negative - reduces LRC)
        worksheet_rollforward.write_formula(3 + row, 3, f"=-'Liability Report'!F{5+row}", currency_format)
        
        # Loss Component
        worksheet_rollforward.write_formula(3 + row, 4, f"='Onerous Test'!H{4+row}", currency_format)
        
        # Closing LRC
        worksheet_rollforward.write_formula(3 + row, 5, f"=B{4+row}+C{4+row}+D{4+row}+E{4+row}", currency_format)
        
        # Opening LIC
        worksheet_rollforward.write(3 + row, 6, 0, currency_format)
        
        # Claims Paid (demo data)
        worksheet_rollforward.write(3 + row, 7, 0, currency_format)
        
        # Closing LIC
        worksheet_rollforward.write_formula(3 + row, 8, f"=G{4+row}+H{4+row}", currency_format)
    
    # ========================================================================
    # SHEET 9: DASHBOARD
    # ========================================================================
    
    worksheet_dashboard = workbook.add_worksheet('Dashboard')
    worksheet_dashboard.merge_range('A1:F1', 'EXECUTIVE DASHBOARD', title_format)
    
    # Key Metrics
    metrics = [
        ['KEY METRICS', ''],
        ['Total Contracts', "=COUNTA('Contract Input'!A4:A8)"],
        ['PAA Eligible Contracts', '=COUNTIF(\'PAA Eligibility Test\'!F4:F8,"YES")'],
        ['Onerous Contracts', '=COUNTIF(\'Onerous Test\'!I4:I8,"ONEROUS")'],
        ['', ''],
        ['FINANCIAL SUMMARY', ''],
        ['Total Written Premium', "=SUM('Contract Input'!F4:F8)"],
        ['Total Earned Premium', "=SUM('Liability Report'!F5:F9)"],
        ['Total Unearned Premium', "=SUM('Liability Report'!G5:G9)"],
        ['Total Expected Claims', "=SUM('Contract Input'!G4:G8)"],
        ['', ''],
        ['LIABILITIES', ''],
        ['Total LRC', "=SUM('Liability Report'!H5:H9)"],
        ['Total LIC', "=SUM('Liability Report'!I5:I9)"],
        ['Total Loss Component', "=SUM('Liability Report'!J5:J9)"],
        ['Total Liability', "=SUM('Liability Report'!K5:K9)"],
        ['', ''],
        ['PROFITABILITY', ''],
        ['Insurance Revenue', "=SUM('Income Statement'!B5:B9)"],
        ['Insurance Service Expenses', "=SUM('Income Statement'!E5:E9)"],
        ['Insurance Service Result', "=SUM('Income Statement'!F5:F9)"],
        ['Average Margin %', "='Income Statement'!G10"],
    ]
    
    worksheet_dashboard.set_column('A:A', 30)
    worksheet_dashboard.set_column('B:B', 25)
    
    for row, metric in enumerate(metrics):
        if metric[0] in ['KEY METRICS', 'FINANCIAL SUMMARY', 'LIABILITIES', 'PROFITABILITY']:
            worksheet_dashboard.write(2 + row, 0, metric[0], subtitle_format)
        elif metric[0] == '':
            continue
        else:
            worksheet_dashboard.write(2 + row, 0, metric[0])
            if metric[1].startswith('='):
                if 'SUM' in metric[1] or 'COUNT' in metric[1]:
                    worksheet_dashboard.write_formula(2 + row, 1, metric[1], calc_format)
                else:
                    worksheet_dashboard.write_formula(2 + row, 1, metric[1], percent_format if 'Margin' in metric[0] else calc_format)
            else:
                worksheet_dashboard.write(2 + row, 1, metric[1])
    
    # ========================================================================
    # SHEET 10: FORMULAS REFERENCE
    # ========================================================================
    
    formulas_data = {
        'Calculation': [
            'PAA Eligibility - Coverage Test',
            'PAA Eligibility - Materiality Test',
            'Onerous Test - Net Premium',
            'Onerous Test - Fulfilment Cash Flows',
            'Onerous Test - Loss Component',
            'Earned Premium',
            'Unearned Premium',
            'LRC (Liability Remaining Coverage)',
            'LIC (Liability Incurred Claims)',
            'Insurance Revenue',
            'Insurance Service Result'
        ],
        'Formula': [
            'Coverage Period <= 12 months',
            'Discount Impact < 5% of Premium',
            'Premium - Acquisition Costs',
            'Expected Claims + Risk Adjustment',
            'MAX(Fulfilment CF - Net Premium, 0)',
            'Premium × (Days Elapsed / Total Days)',
            'Premium - Earned Premium',
            'Unearned Premium + Loss Component',
            'Claims Incurred - Claims Paid',
            'Earned Premium - Loss Component Release',
            'Insurance Revenue - Insurance Service Expenses'
        ],
        'IFRS_17_Reference': [
            'Para 53',
            'Para 53',
            'Para 47',
            'Para 32-33',
            'Para 47-52',
            'Para 55',
            'Para 55',
            'Para 55',
            'Para 55',
            'Para 83',
            'Para 80-81'
        ]
    }
    
    df_formulas = pd.DataFrame(formulas_data)
    df_formulas.to_excel(writer, sheet_name='Formula Reference', index=False, startrow=2)
    
    worksheet_formulas = writer.sheets['Formula Reference']
    worksheet_formulas.merge_range('A1:C1', 'CALCULATION FORMULAS & IFRS 17 REFERENCES', title_format)
    worksheet_formulas.set_column('A:A', 35)
    worksheet_formulas.set_column('B:B', 50)
    worksheet_formulas.set_column('C:C', 20)
    
    for col_num, value in enumerate(df_formulas.columns.values):
        worksheet_formulas.write(2, col_num, value, header_format)
    
    # ========================================================================
    # SAVE WORKBOOK
    # ========================================================================
    
    writer.close()
    print(f"\n{'='*80}")
    print(f"✓ IFRS 17 PAA Complete Tool Created Successfully!")
    print(f"{'='*80}")
    print(f"\nFile: {filename}")
    print(f"\nSheets included:")
    print(f"  1. Instructions - User guide")
    print(f"  2. Config - Settings and assumptions")
    print(f"  3. Contract Input - Enter contract data here")
    print(f"  4. PAA Eligibility Test - Automatic eligibility testing")
    print(f"  5. Onerous Test - Identifies loss-making contracts")
    print(f"  6. Liability Report - Comprehensive liability measurements")
    print(f"  7. Income Statement - Insurance service results")
    print(f"  8. Roll Forward - Liability reconciliation")
    print(f"  9. Dashboard - Executive summary")
    print(f" 10. Formula Reference - Calculation details & IFRS 17 refs")
    print(f"\n{'='*80}")
    print(f"Ready to use! Open the file and start with the Instructions sheet.")
    print(f"{'='*80}\n")

if __name__ == "__main__":
    create_comprehensive_ifrs17_excel()
