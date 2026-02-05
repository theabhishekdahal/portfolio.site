# IFRS 17 PAA Excel Tool

A comprehensive Python tool for creating IFRS 17 Premium Allocation Approach (PAA) Excel workbooks with automated calculations and reporting.

## Overview

This tool generates a fully-functional Excel workbook that implements IFRS 17 Premium Allocation Approach calculations, including:

- PAA eligibility testing
- Onerous contract identification
- Liability measurement (LRC & LIC)
- Premium earning calculations
- Insurance service results
- Roll-forward reconciliation
- Executive dashboard

## Requirements

- Python 3.7 or higher
- pandas >= 2.0.0
- xlsxwriter >= 3.0.0

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script to generate the Excel workbook:

```bash
python ifrs17_paa_tool.py
```

This will create a file named `IFRS17_PAA_Complete_Tool.xlsx` in the current directory.

## Features

### 10 Comprehensive Worksheets

1. **Instructions** - User guide with step-by-step instructions
2. **Config** - Configuration parameters and assumptions
3. **Contract Input** - Input sheet for contract data
4. **PAA Eligibility Test** - Automated testing for PAA qualification
5. **Onerous Test** - Identifies loss-making contracts
6. **Liability Report** - Comprehensive liability measurements
7. **Income Statement** - Insurance service results
8. **Roll Forward** - Period-over-period reconciliation
9. **Dashboard** - Executive summary with key metrics
10. **Formula Reference** - Calculation formulas and IFRS 17 references

### Key Capabilities

- ✓ **PAA Eligibility Testing** - Automatic test for coverage period and materiality
- ✓ **Onerous Contract Test** - Identifies loss-making contracts
- ✓ **Liability Measurement** - LRC, LIC, and Loss Component calculations
- ✓ **Premium Earning** - Straight-line amortization over coverage period
- ✓ **Roll-forward** - Period-over-period reconciliation
- ✓ **Income Statement** - Insurance revenue and service expenses
- ✓ **Dashboard** - Visual summary of key metrics

### Sample Data Included

The tool comes with 5 sample insurance contracts demonstrating:
- Property Insurance
- Auto Insurance  
- Liability Insurance
- Various coverage periods (6-24 months)
- Mixed profitability scenarios

## Customization

The generated Excel file includes:
- **Pre-built formulas** that automatically calculate based on input data
- **Conditional formatting** for visual indicators (red for warnings, green for success)
- **Professional formatting** with color-coded headers and sections
- **Flexible configuration** through the Config sheet

## IFRS 17 Compliance

All calculations follow IFRS 17 standard requirements:
- Coverage period threshold testing (Para 53)
- Materiality testing for discount effects (Para 53)
- Onerous contract identification (Para 47-52)
- Premium earning recognition (Para 55)
- Liability measurement (Para 32-33, 55)
- Insurance service result presentation (Para 80-83)

## Workflow

1. **Start** - Read Instructions sheet
2. **Configure** - Set parameters in Config sheet
3. **Input Data** - Enter contract details in Contract Input sheet
4. **Review** - Check PAA Eligibility and Onerous Test results
5. **Reports** - View Liability Report and Income Statement
6. **Dashboard** - Review executive summary

## Output Format

The Excel file includes:
- Professional color-coded formatting
- Automated formulas with cell references
- Conditional formatting for status indicators
- Comprehensive documentation
- Ready for immediate use

## Notes

- All monetary values are in USD by default (configurable)
- Dates use yyyy-mm-dd format
- Coverage periods are in months
- Discount rates are annual percentages
- The tool is designed for demonstration and can be extended with actual claims data

## Author

Created as part of the portfolio.site project by Abhishek Dahal.

## License

See the main repository LICENSE file for details.
