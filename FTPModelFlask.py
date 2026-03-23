from flask import Flask, request, jsonify, send_file
import pandas as pd
import io
import openpyxl
import numpy as np
from datetime import datetime, timedelta

app = Flask(__name__)

# Define the tenor points (in days)
tenors = [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800]

# USD FTP curve data (matching your provided values)
usd_rates = [3.29, 3.36, 6.15, 10.97, 11.02, 11.13, 12.22, 12.22, 12.22, 13.96, 15.41, 18.32, 18.32]

# ZWG FTP curve data (matching your provided values)
zwg_rates = [16.90, 16.90, 16.90, 16.90, 17.90, 18.10, 19.10, 19.10, 20.10, 23.47, 26.54, 32.67, 32.67]

# Global variable to store the latest uploaded data (in-memory, no database)
latest_data = {
    'filename': None,
    'sheets': {},  # sheet name -> DataFrame as dict
    'ftp_results': None,
    'full_dataframes': {},  # Store full dataframes for download
    'summaries': {},  # Store summaries
    'period': {}  # Store period info
}

@app.route('/download-pdf', methods=['GET'])
def download_pdf():
    """Download the generated PDF report"""
    try:
        if not latest_data.get('summaries') or not latest_data.get('period'):
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404
        
        pdf_buffer = generate_pdf_report()
        
        month = latest_data.get('period', {}).get('month', 'Report')
        year = latest_data.get('period', {}).get('year', '')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"FTP_Report_{month}_{year}_{timestamp}.pdf"
        
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"Error downloading PDF: {str(e)}")
        return jsonify({'error': f'Failed to download PDF: {str(e)}'}), 500

def compute_ftp_components(deposit, loan, tenure):
    """Helper to compute FTP charge, gain, net (matches frontend logic)"""
    try:
        d = float(deposit) if deposit else 0
        l = float(loan) if loan else 0
        t = float(tenure) if tenure else 1
        
        if d <= 0 or l <= 0 or t <= 0:
            return {'charge': '1.12', 'gain': '0.72', 'net': '1.84'}
        
        risk_factor = min(l / d, 2.0)
        tenure_factor = min(t * 0.07, 1.2)
        
        charge = round(1.0 + (risk_factor * 0.3) + (tenure_factor * 0.2), 2)
        gain = round(0.6 + (min(d / 200000, 1.5) * 0.3) + (t * 0.04), 2)
        net = round(gain - charge, 2)
        
        return {
            'charge': f"{charge:.2f}",
            'gain': f"{gain:.2f}",
            'net': f"{net:.2f}"
        }
    except:
        return {'charge': '1.12', 'gain': '0.72', 'net': '1.84'}

@app.route('/')
def index():
    """Serve the main HTML page"""
    from flask import send_from_directory
    return send_from_directory('.', 'index.html')


@app.route('/download-excel', methods=['GET'])
def download_excel():
    """Download summary results as Excel file"""
    try:
        # Check if we have summaries
        if not latest_data.get('summaries') or not latest_data.get('period'):
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Create summary data
            summary_data = []
            
            # Add period information
            summary_data.append({
                'Type': 'PERIOD INFORMATION',
                'Detail': '',
                'Value': ''
            })
            summary_data.append({
                'Type': 'File Name',
                'Detail': latest_data.get('filename', 'N/A'),
                'Value': ''
            })
            summary_data.append({
                'Type': 'Period',
                'Detail': f"{latest_data['period'].get('first_day', 'N/A')} to {latest_data['period'].get('last_day', 'N/A')}",
                'Value': ''
            })
            summary_data.append({
                'Type': 'Month',
                'Detail': latest_data['period'].get('month', 'N/A'),
                'Value': ''
            })
            summary_data.append({
                'Type': 'Year',
                'Detail': latest_data['period'].get('year', 'N/A'),
                'Value': ''
            })
            summary_data.append({
                'Type': '',
                'Detail': '',
                'Value': ''
            })
            summary_data.append({
                'Type': 'FTP SUMMARY BY CURRENCY',
                'Detail': '',
                'Value': ''
            })
            summary_data.append({
                'Type': '',
                'Detail': '',
                'Value': ''
            })
            
            # Add summary by currency
            for currency, sheets in latest_data['summaries'].items():
                for sheet_name, sheet_data in sheets.items():
                    summary_data.append({
                        'Type': f'CURRENCY: {currency}',
                        'Detail': sheet_name,
                        'Value': ''
                    })
                    summary_data.append({
                        'Type': '  Total Exposure',
                        'Detail': f"{sheet_data['total_exposure']:,.2f}",
                        'Value': currency
                    })
                    summary_data.append({
                        'Type': '  Total FTP Charge',
                        'Detail': f"{sheet_data['total_ftp_charge']:,.2f}",
                        'Value': currency
                    })
                    summary_data.append({
                        'Type': '  Number of Records',
                        'Detail': f"{sheet_data['row_count']:,}",
                        'Value': ''
                    })
                    
                    # Add SBU breakdown
                    if sheet_data.get('by_sbu'):
                        summary_data.append({
                            'Type': '  Breakdown by SBU',
                            'Detail': '',
                            'Value': ''
                        })
                        for sbu in sheet_data['by_sbu']:
                            summary_data.append({
                                'Type': f'    {sbu["SBU"]}',
                                'Detail': f"Exposure: {sbu['Currency Exposure + Currency Accrued Reporting']:,.2f}",
                                'Value': f"FTP Charge: {sbu['FTP Charge']:,.2f}"
                            })
                    
                    summary_data.append({
                        'Type': '',
                        'Detail': '',
                        'Value': ''
                    })
            
            # Add grand totals
            total_exposure_zwg = 0
            total_ftp_zwg = 0
            total_exposure_fx = 0
            total_ftp_fx = 0
            
            for currency, sheets in latest_data['summaries'].items():
                for sheet_data in sheets.values():
                    if currency == 'ZWG':
                        total_exposure_zwg += sheet_data['total_exposure']
                        total_ftp_zwg += sheet_data['total_ftp_charge']
                    else:
                        total_exposure_fx += sheet_data['total_exposure']
                        total_ftp_fx += sheet_data['total_ftp_charge']
            
            summary_data.append({
                'Type': 'GRAND TOTALS',
                'Detail': '',
                'Value': ''
            })
            
            if total_exposure_zwg > 0:
                summary_data.append({
                    'Type': 'ZWG Total Exposure',
                    'Detail': f"{total_exposure_zwg:,.2f} ZWG",
                    'Value': ''
                })
                summary_data.append({
                    'Type': 'ZWG Total FTP Charge',
                    'Detail': f"{total_ftp_zwg:,.2f} ZWG",
                    'Value': ''
                })
            
            if total_exposure_fx > 0:
                summary_data.append({
                    'Type': 'USD Total Exposure',
                    'Detail': f"{total_exposure_fx:,.2f} USD",
                    'Value': ''
                })
                summary_data.append({
                    'Type': 'USD Total FTP Charge',
                    'Detail': f"{total_ftp_fx:,.2f} USD",
                    'Value': ''
                })
            
            summary_data.append({
                'Type': '',
                'Detail': '',
                'Value': ''
            })
            summary_data.append({
                'Type': 'Generated by FTP Central System',
                'Detail': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Value': ''
            })
            
            # Create DataFrame and write to Excel
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='FTP Summary', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['FTP Summary']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        
        # Generate filename
        month = latest_data.get('period', {}).get('month', 'Report')
        year = latest_data.get('period', {}).get('year', '')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"FTP_Summary_{month}_{year}_{timestamp}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Error downloading Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to download Excel: {str(e)}'}), 500


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle Excel file upload, read all sheets into DataFrames"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
    
    try:
        import re
        from datetime import datetime, timedelta
        import pandas as pd
        import datetime as dt
        
        # Parse filename to get month and year
        filename_match = re.search(r'FTP Input File (\w+) (\d{4})', file.filename)
        if not filename_match:
            return jsonify({'error': 'Filename must be in format: FTP Input File Month Year.xlsx'}), 400
        
        month_name = filename_match.group(1)
        year = int(filename_match.group(2))
        
        # Convert month name to month number
        month_map = {
            'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
            'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12
        }
        month_num = month_map.get(month_name.lower())
        if not month_num:
            return jsonify({'error': f'Invalid month name: {month_name}'}), 400
        
        # Get first and last day of the month
        first_day = datetime(year, month_num, 1)
        if month_num == 12:
            last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(year, month_num + 1, 1) - timedelta(days=1)
        
        # Read all sheet names
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        print(f"Found sheets: {sheet_names}")
        
        # Dictionary to store processed data
        sheets_data = {}
        
        # Store summaries by currency and SBU
        global_summaries = {
            'ZWG': {},
            'FX': {}
        }
        
        # Process each sheet separately
        for sheet in sheet_names:
            print(f"Processing sheet: {sheet}")
            
            try:
                # Read the sheet
                df = pd.read_excel(file, sheet_name=sheet)
            except Exception as e:
                print(f"  Error reading sheet {sheet}: {e}")
                continue
            
            original_shape = df.shape
            print(f"  Original shape: {original_shape}")
            
            # Process only LOANS sheets
            if sheet in ["ZWG LOANS", "FX LOANS"]:
                print(f"  Applying {sheet} special handling...")
                
                df_processed = df.copy()
                del df  # Free memory
                
                # Branch mapping dictionary (same as before)
                branch_sbu_map = {
                    '106': {'unit': 'Agribusiness', 'sbu': 'Corporate Banking'},
                    '118': {'unit': 'Bureau De Change Hre', 'sbu': 'Shared Services'},
                    '45': {'unit': 'Business Banking', 'sbu': 'Corporate Banking'},
                    '108': {'unit': 'Business Banking', 'sbu': 'Corporate Banking'},
                    '47': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
                    '113': {'unit': 'Custodial Services', 'sbu': 'Corporate Banking'},
                    '48': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
                    '53': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
                    '602': {'unit': 'Mortgage Finance', 'sbu': 'Retail Banking'},
                    '107': {'unit': 'Institutional Banking', 'sbu': 'Corporate Banking'},
                    '66': {'unit': 'Treasury', 'sbu': 'Treasury'},
                    '601': {'unit': 'Treasury', 'sbu': 'Treasury'},
                    '0': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '35': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '36': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '37': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '38': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '39': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '40': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '41': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '43': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '46': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '49': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '50': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '54': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '56': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '57': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '58': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '61': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '65': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '67': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '68': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '69': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '70': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '105': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '115': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '117': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '123': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '124': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '600': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '141': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '116': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '11': {'unit': 'Kwame Nkrumah', 'sbu': 'Retail Banking'},
                    '12': {'unit': '8Th Avenue', 'sbu': 'Retail Banking'},
                    '13': {'unit': 'Mutare', 'sbu': 'Retail Banking'},
                    '14': {'unit': 'Kwekwe', 'sbu': 'Retail Banking'},
                    '15': {'unit': 'Chitungwiza', 'sbu': 'Retail Banking'},
                    '17': {'unit': 'Gokwe', 'sbu': 'Retail Banking'},
                    '18': {'unit': 'Gweru', 'sbu': 'Retail Banking'},
                    '20': {'unit': 'Chivhu', 'sbu': 'Retail Banking'},
                    '21': {'unit': 'Selous', 'sbu': 'Retail Banking'},
                    '23': {'unit': 'Southerton', 'sbu': 'Retail Banking'},
                    '24': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '25': {'unit': 'Masvingo', 'sbu': 'Retail Banking'},
                    '26': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
                    '27': {'unit': 'Cash Depot Bulawayo', 'sbu': 'Retail Banking'},
                    '28': {'unit': 'Chiredzi', 'sbu': 'Retail Banking'},
                    '29': {'unit': 'Borrowdale', 'sbu': 'Retail Banking'},
                    '30': {'unit': 'Avondale', 'sbu': 'Retail Banking'},
                    '31': {'unit': 'Chinhoyi', 'sbu': 'Retail Banking'},
                    '32': {'unit': 'Kwekwe', 'sbu': 'Retail Banking'},
                    '33': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '34': {'unit': 'Cash Depot Harare', 'sbu': 'Retail Banking'},
                    '44': {'unit': 'Wealth Management', 'sbu': 'Retail Banking'},
                    '87': {'unit': 'Chipinge', 'sbu': 'Retail Banking'},
                    '88': {'unit': '8Th Avenue', 'sbu': 'Retail Banking'},
                    '89': {'unit': 'Highfield', 'sbu': 'Retail Banking'},
                    '90': {'unit': 'Marondera', 'sbu': 'Retail Banking'},
                    '91': {'unit': 'Chitungwiza', 'sbu': 'Retail Banking'},
                    '92': {'unit': 'Gokwe', 'sbu': 'Retail Banking'},
                    '93': {'unit': 'Beitbridge', 'sbu': 'Retail Banking'},
                    '95': {'unit': 'Kariba', 'sbu': 'Retail Banking'},
                    '96': {'unit': 'Kariba', 'sbu': 'Retail Banking'},
                    '97': {'unit': 'Karoi', 'sbu': 'Retail Banking'},
                    '98': {'unit': 'Chinhoyi', 'sbu': 'Retail Banking'},
                    '99': {'unit': 'Masvingo', 'sbu': 'Retail Banking'},
                    '100': {'unit': 'Mvurwi', 'sbu': 'Retail Banking'},
                    '101': {'unit': 'Chipinge', 'sbu': 'Retail Banking'},
                    '102': {'unit': 'Rusape', 'sbu': 'Retail Banking'},
                    '103': {'unit': 'Murehwa', 'sbu': 'Retail Banking'},
                    '104': {'unit': 'Victoria Falls', 'sbu': 'Retail Banking'},
                    '109': {'unit': 'Chiredzi', 'sbu': 'Retail Banking'},
                    '110': {'unit': 'Selous', 'sbu': 'Retail Banking'},
                    '111': {'unit': 'Selous', 'sbu': 'Retail Banking'},
                    '112': {'unit': 'Mvurwi', 'sbu': 'Retail Banking'},
                    '51': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
                    '52': {'unit': 'Kwame Nkrumah', 'sbu': 'Retail Banking'},
                    '55': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '62': {'unit': '8Th Avenue', 'sbu': 'Retail Banking'},
                    '114': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
                    '120': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '121': {'unit': 'Retail Centraslised Back Office', 'sbu': 'Retail Banking'},
                    '122': {'unit': 'Mta Centre Fife Street', 'sbu': 'Retail Banking'},
                    '611': {'unit': 'Masvingo', 'sbu': 'Retail Banking'},
                    '612': {'unit': 'Chiredzi', 'sbu': 'Retail Banking'},
                    '613': {'unit': 'Masvingo', 'sbu': 'Retail Banking'},
                    '614': {'unit': 'Zvishavane', 'sbu': 'Retail Banking'},
                    '615': {'unit': 'Gweru', 'sbu': 'Retail Banking'},
                    '616': {'unit': 'Kwekwe', 'sbu': 'Retail Banking'},
                    '617': {'unit': 'Kadoma', 'sbu': 'Retail Banking'},
                    '618': {'unit': 'Kadoma', 'sbu': 'Retail Banking'},
                    '619': {'unit': 'Gokwe', 'sbu': 'Retail Banking'},
                    '629': {'unit': 'Chipinge', 'sbu': 'Retail Banking'},
                    '630': {'unit': 'Chipinge', 'sbu': 'Retail Banking'},
                    '631': {'unit': 'Mutare', 'sbu': 'Retail Banking'},
                    '632': {'unit': 'Mutare', 'sbu': 'Retail Banking'},
                    '633': {'unit': 'Mutare', 'sbu': 'Retail Banking'},
                    '634': {'unit': 'Rusape', 'sbu': 'Retail Banking'},
                    '644': {'unit': '8Th Avenue', 'sbu': 'Retail Banking'},
                    '645': {'unit': '8Th Avenue', 'sbu': 'Retail Banking'},
                    '646': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
                    '647': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
                    '648': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
                    '649': {'unit': 'Gwanda', 'sbu': 'Retail Banking'},
                    '650': {'unit': 'Cash Depot Bulawayo', 'sbu': 'Retail Banking'},
                    '660': {'unit': 'Samora Machel', 'sbu': 'Retail Banking'},
                    '661': {'unit': 'Avondale', 'sbu': 'Retail Banking'},
                    '662': {'unit': 'Bindura', 'sbu': 'Retail Banking'},
                    '663': {'unit': 'Msasa', 'sbu': 'Retail Banking'},
                    '664': {'unit': 'Chinhoyi', 'sbu': 'Retail Banking'},
                    '665': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '667': {'unit': 'Karoi', 'sbu': 'Retail Banking'},
                    '668': {'unit': 'Murehwa', 'sbu': 'Retail Banking'},
                    '669': {'unit': 'Samora Machel', 'sbu': 'Retail Banking'},
                    '670': {'unit': 'Samora Machel', 'sbu': 'Retail Banking'},
                    '671': {'unit': 'Cash Depot Harare', 'sbu': 'Retail Banking'},
                    '672': {'unit': 'Kariba', 'sbu': 'Retail Banking'},
                    '681': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '682': {'unit': 'Cripps', 'sbu': 'Retail Banking'},
                    '683': {'unit': 'Chitungwiza', 'sbu': 'Retail Banking'},
                    '684': {'unit': 'Chivhu', 'sbu': 'Retail Banking'},
                    '685': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '686': {'unit': 'Highfield', 'sbu': 'Retail Banking'},
                    '687': {'unit': 'Marondera', 'sbu': 'Retail Banking'},
                    '688': {'unit': 'Msasa', 'sbu': 'Retail Banking'},
                    '689': {'unit': 'Msasa', 'sbu': 'Retail Banking'},
                    '690': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
                    '125': {'unit': 'Passport Centre Harare', 'sbu': 'Shared Services'},
                    '127': {'unit': 'Passport Centre Bulawayo', 'sbu': 'Shared Services'},
                    '126': {'unit': 'Virtual Branch', 'sbu': 'Shared Services'},
                    '128': {'unit': 'Passport Centre Chitungwiza', 'sbu': 'Shared Services'},
                    '129': {'unit': 'Passport Centre Lupane', 'sbu': 'Shared Services'},
                    '130': {'unit': 'Passport Centre Hwange', 'sbu': 'Shared Services'},
                    '131': {'unit': 'Passport Centre Gweru', 'sbu': 'Shared Services'},
                    '132': {'unit': 'Passport Centre Beitbridge', 'sbu': 'Shared Services'},
                    '133': {'unit': 'Passport Centre Chinhoyi', 'sbu': 'Shared Services'},
                    '134': {'unit': 'Passport Centre Marondera', 'sbu': 'Shared Services'},
                    '135': {'unit': 'Passport Centre Bindura', 'sbu': 'Shared Services'},
                    '136': {'unit': 'Passport Centre Gwanda', 'sbu': 'Shared Services'},
                    '137': {'unit': 'Passport Centre Mutare', 'sbu': 'Shared Services'},
                    '138': {'unit': 'Passport Centre Masvingo', 'sbu': 'Shared Services'},
                    '139': {'unit': 'Passport Centre Zvishavane', 'sbu': 'Shared Services'},
                    '140': {'unit': 'Passport Centre Murehwa', 'sbu': 'Shared Services'},
                    '142': {'unit': 'Retail Centralised Byo', 'sbu': 'Retail Banking'},
                    '145': {'unit': 'Borrowdale', 'sbu': 'Retail Banking'},
                    '146': {'unit': 'Passport Centre Mwenezi', 'sbu': 'Shared Services'},
                    '200': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
                    '147': {'unit': 'Passport Centre Gokwe', 'sbu': 'Shared Services'},
                    '143': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
                    '144': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'}
                }
                
                print(f"  Created branch mapping with {len(branch_sbu_map)} unique branch codes")
                
                # Add SBU column
                branch_code_col = None
                possible_names = ['Branch Code', 'BRANCHCODE', 'BRANCH_CODE', 'BRANCH', 'BR_CODE']
                for col in possible_names:
                    if col in df_processed.columns:
                        branch_code_col = col
                        break
                
                if branch_code_col:
                    df_processed[branch_code_col] = df_processed[branch_code_col].astype(str).str.strip()
                    branch_info = df_processed[branch_code_col].map(branch_sbu_map)
                    df_processed['ACC MANAGEMENT UNIT'] = branch_info.apply(lambda x: x['unit'] if isinstance(x, dict) else 'Unknown')
                    df_processed['SBU'] = branch_info.apply(lambda x: x['sbu'] if isinstance(x, dict) else 'Unknown')
                    
                    # Override SBU for LoC rows
                    if 'Source of Funding' in df_processed.columns:
                        df_processed['Source of Funding'] = df_processed['Source of Funding'].astype(str).str.strip()
                        loc_mask = df_processed['Source of Funding'].str.upper() == 'LOC'
                        df_processed.loc[loc_mask, 'SBU'] = 'Corporate Banking'
                        print(f"  Found {loc_mask.sum()} rows with Source of Funding = 'LoC'")
                    
                    unknown_count = df_processed['ACC MANAGEMENT UNIT'].isna().sum()
                    df_processed['ACC MANAGEMENT UNIT'] = df_processed['ACC MANAGEMENT UNIT'].fillna('Unknown')
                    df_processed['SBU'] = df_processed['SBU'].fillna('Unknown')
                    print(f"  {unknown_count} rows with unknown branch codes")
                else:
                    df_processed['ACC MANAGEMENT UNIT'] = 'Unknown'
                    df_processed['SBU'] = 'Unknown'
                
                # Date processing
                if 'BOOKING_DATE' in df_processed.columns:
                    df_processed['BOOKING_DATE'] = pd.to_datetime(df_processed['BOOKING_DATE'], errors='coerce')
                    booking_date_mask = df_processed['BOOKING_DATE'].isna()
                    df_processed.loc[booking_date_mask, 'BOOKING_DATE'] = first_day
                    print(f"  Updated {booking_date_mask.sum()} rows with BOOKING_DATE = {first_day.strftime('%Y-%m-%d')}")
                
                if 'MATURITY_DATE' in df_processed.columns:
                    df_processed['MATURITY_DATE'] = pd.to_datetime(df_processed['MATURITY_DATE'], errors='coerce')
                    maturity_date_mask = df_processed['MATURITY_DATE'].isna()
                    df_processed.loc[maturity_date_mask, 'MATURITY_DATE'] = first_day + timedelta(days=365)
                    print(f"  Updated {maturity_date_mask.sum()} rows with MATURITY_DATE")
                
                # Calculate TENOR
                if 'BOOKING_DATE' in df_processed.columns and 'MATURITY_DATE' in df_processed.columns:
                    df_processed['TENOR'] = (df_processed['MATURITY_DATE'] - df_processed['BOOKING_DATE']).dt.days
                    df_processed.loc[df_processed['TENOR'] < 0, 'TENOR'] = 0
                
                # Calculate DimDays (overlapping days with FTP period)
                def calculate_dim_days(row):
                    booking_date = row['BOOKING_DATE'].date() if hasattr(row['BOOKING_DATE'], 'date') else row['BOOKING_DATE']
                    maturity_date = row['MATURITY_DATE'].date() if hasattr(row['MATURITY_DATE'], 'date') else row['MATURITY_DATE']
                    
                    if booking_date <= first_day.date() and maturity_date >= last_day.date():
                        return (last_day.date() - first_day.date()).days + 1
                    elif booking_date >= first_day.date() and maturity_date >= last_day.date():
                        return (last_day.date() - booking_date).days + 1
                    elif booking_date >= first_day.date() and maturity_date <= last_day.date():
                        return (maturity_date - booking_date).days
                    elif booking_date < first_day.date() and maturity_date < last_day.date():
                        return (last_day.date() - first_day.date()).days
                    else:
                        return (maturity_date - first_day.date()).days
                
                df_processed['DimDays'] = df_processed.apply(calculate_dim_days, axis=1)
                
                # Calculate DTM and MTM
                last_day_ts = pd.Timestamp(last_day.date())
                df_processed['DTM'] = df_processed.apply(
                    lambda row: (row['MATURITY_DATE'] - last_day_ts).days if row['MATURITY_DATE'] > last_day_ts else 0,
                    axis=1
                )
                df_processed['MTM'] = (df_processed['DTM'] / 30).round(1)
                
                # Create bucket columns
                tenors = [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800]
                bucket_labels = []
                for i in range(len(tenors)):
                    if i == 0:
                        bucket_labels.append(f'<{tenors[i]}days')
                    else:
                        bucket_labels.append(f'{tenors[i-1]}-{tenors[i]}days')
                bucket_labels.append(f'+{tenors[-1]}days')
                
                for label in bucket_labels:
                    df_processed[label] = 0
                
                # Allocate to buckets
                exposure = df_processed['Currency Exposure + Currency Accrued Reporting']
                mtm_days = df_processed['MTM'] * 30
                bin_edges = [0] + tenors + [float('inf')]
                bucket_indices = pd.cut(mtm_days, bins=bin_edges, labels=False, right=False, include_lowest=True)
                bucket_indices = bucket_indices.fillna(len(bucket_labels) - 1).astype(int)
                
                for i, label in enumerate(bucket_labels):
                    mask = (bucket_indices == i) & (exposure > 0)
                    df_processed.loc[mask, label] = exposure.loc[mask]
                
                df_processed['Check'] = df_processed[bucket_labels].sum(axis=1)
                
                # Calculate FTP Charge based on currency
                if sheet == "ZWG LOANS":
                    rates = zwg_rates
                else:
                    rates = usd_rates
                
                rate_array = np.zeros(len(df_processed))
                for i, label in enumerate(bucket_labels):
                    if i < len(rates):
                        rate = rates[i]
                    else:
                        rate = rates[-1]
                    rate_array += df_processed[label] * (rate / 100)
                
                df_processed['FTP Charge'] = rate_array
                
                # Create summaries by SBU and currency
                currency = 'ZWG' if sheet == "ZWG LOANS" else 'FX'
                
                # Group by SBU and calculate totals
                sbu_summary = df_processed.groupby('SBU').agg({
                    'Currency Exposure + Currency Accrued Reporting': 'sum',
                    'FTP Charge': 'sum'
                }).reset_index()
                
                # Store in global summaries
                global_summaries[currency][sheet] = {
                    'total_exposure': df_processed['Currency Exposure + Currency Accrued Reporting'].sum(),
                    'total_ftp_charge': df_processed['FTP Charge'].sum(),
                    'by_sbu': sbu_summary.to_dict(orient='records'),
                    'row_count': len(df_processed)
                }
                
                # Store preview data (first 100 rows) for display
                preview_data = df_processed.head(100).copy()
                
                # Convert datetime columns to strings for JSON serialization
                for col in preview_data.select_dtypes(include=['datetime64', 'datetime64[ns]']).columns:
                    preview_data[col] = preview_data[col].astype(str).replace('NaT', None)
                
                sheets_data[sheet] = {
                    'columns': df_processed.columns.tolist(),
                    'data': preview_data.to_dict(orient='records'),
                    'shape': df_processed.shape,
                    'summary': {
                        'total_exposure': float(df_processed['Currency Exposure + Currency Accrued Reporting'].sum()),
                        'total_ftp_charge': float(df_processed['FTP Charge'].sum()),
                        'by_sbu': sbu_summary.to_dict(orient='records')
                    }
                }
                latest_data['full_dataframes'][sheet] = df_processed.copy()
                # Free memory
                del df_processed
                
            else:
                # For non-loan sheets, store preview only
                preview_data = df.head(100).copy()
                sheets_data[sheet] = {
                    'columns': df.columns.tolist(),
                    'data': preview_data.to_dict(orient='records'),
                    'shape': df.shape,
                    'summary': None
                }
                latest_data['full_dataframes'][sheet] = df.copy()


                del df
            
            print(f"  Completed processing {sheet}")
        
        # Store in global variable
        latest_data['filename'] = file.filename
        latest_data['sheets'] = sheets_data
        latest_data['summaries'] = global_summaries
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'year': year
        }
        
        # Return summary without trying to access df_processed
        return jsonify({
            'status': 'success',
            'message': f'Successfully processed {len(sheet_names)} sheets',
            'summary': global_summaries,
            'period': latest_data['period']
        })
    
    except Exception as e:
        print(f"Error processing upload: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/calculate', methods=['POST'])
def calculate():
    """Calculate FTP based on inputs and return results"""
    data = request.json
    deposit = data.get('deposit')
    loan = data.get('loan')
    tenure = data.get('tenure')
    
    result = compute_ftp_components(deposit, loan, tenure)
    
    # Store results for download
    latest_data['ftp_results'] = {
        'deposit': deposit,
        'loan': loan,
        'tenure': tenure,
        'charge': result['charge'],
        'gain': result['gain'],
        'net': result['net'],
        'timestamp': datetime.now().isoformat()
    }
    
    return jsonify(result)

@app.route('/download-results', methods=['GET'])
def download_results():
    """Generate and download FTP results as Excel file"""
    results = latest_data.get('ftp_results')
    
    if not results:
        # Default results if none calculated
        results = {
            'deposit': '250000',
            'loan': '300000',
            'tenure': '10',
            'charge': '1.84',
            'gain': '1.12',
            'net': '0.72',
            'timestamp': datetime.now().isoformat()
        }
    
    # Create a DataFrame for results
    results_df = pd.DataFrame([{
        'Parameter': 'Deposit Amount',
        'Value': f"${float(results['deposit']):,.2f}",
        'Unit': 'USD'
    }, {
        'Parameter': 'Loan Amount',
        'Value': f"${float(results['loan']):,.2f}",
        'Unit': 'USD'
    }, {
        'Parameter': 'Loan Tenure',
        'Value': results['tenure'],
        'Unit': 'years'
    }, {
        'Parameter': 'FTP Charge',
        'Value': results['charge'],
        'Unit': '%'
    }, {
        'Parameter': 'FTP Gain',
        'Value': results['gain'],
        'Unit': '%'
    }, {
        'Parameter': 'Net FTP',
        'Value': results['net'],
        'Unit': '%'
    }])
    
    # Create an Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='FTP Results', index=False)
        
        # If we have uploaded sheets, include them as well
        if latest_data['sheets']:
            for sheet_name, sheet_data in latest_data['sheets'].items():
                # Reconstruct DataFrame from stored preview (full data not stored)
                # For demo, we'll just note that data was uploaded
                summary_df = pd.DataFrame([{
                    'Sheet': sheet_name,
                    'Columns': ', '.join(sheet_data['columns']),
                    'Rows': sheet_data['shape'][0],
                    'Preview': 'Data available in original upload'
                }])
                summary_df.to_excel(writer, sheet_name=f'{sheet_name}_summary', index=False)
    
    output.seek(0)
    
    # Generate filename with timestamp
    filename = f"FTP_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/ftp-curve-data', methods=['GET'])
def ftp_curve_data():
    """Return FTP curve data for USD and ZWG"""
    
    global tenors
    global usd_rates
    global zwg_rates
    
    # Check if we have uploaded data that might override these defaults
    if latest_data['sheets']:
        # Look for sheets that might contain curve data
        for sheet_name, sheet_data in latest_data['sheets'].items():
            # If you have specific sheet names that contain curve data, you can parse them here
            # For example, if sheet has columns: Tenor, USD_Rate, ZWG_Rate
            pass
    
    return jsonify({
        'tenors': tenors,
        'zwg': {
            'name': 'ZWG FTP Curve',
            'rates': zwg_rates,
            'color': '#b33a3a',  # Red/burgundy to match your theme
            'borderColor': '#921f1f'
        },
        'usd': {
            'name': 'USD FTP Curve',
            'rates': usd_rates,
            'color': '#2563eb',  # Blue
            'borderColor': '#1d4ed8'
        }
    })

@app.route('/get-preview', methods=['GET'])
def get_preview():
    """Return the current preview data (uploaded sheets)"""
    if latest_data['sheets']:
        return jsonify({
            'filename': latest_data['filename'],
            'sheets': latest_data['sheets']
        })
    else:
        return jsonify({'message': 'No data uploaded yet'}), 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)