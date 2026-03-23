from flask import Flask, request, jsonify, send_file
import pandas as pd
import io
import openpyxl
import numpy as np
from datetime import datetime, timedelta
import re
import gzip
import pickle

app = Flask(__name__)

# Define the tenor points (in days)
tenors = [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800]

# USD FTP curve data
usd_rates = [3.29, 3.36, 6.15, 10.97, 11.02, 11.13, 12.22, 12.22, 12.22, 13.96, 15.41, 18.32, 18.32]

# ZWG FTP curve data
zwg_rates = [16.90, 16.90, 16.90, 16.90, 17.90, 18.10, 19.10, 19.10, 20.10, 23.47, 26.54, 32.67, 32.67]

# Global variable - store compressed data
latest_data = {
    'filename': None,
    'excel_file': None,  # Store the generated Excel file bytes
    'compressed_dataframes': {},  # Store compressed dataframes (memory efficient!)
    'summaries': {},  # Store only summaries (small)
    'period': {},  # Store period info
    'sheets_preview': {}  # Store only preview data (first 100 rows)
}

def compress_dataframe(df):
    """Compress dataframe using gzip + pickle (60-70% smaller)"""
    # Convert datetime to string for better compression
    df_copy = df.copy()
    for col in df_copy.select_dtypes(include=['datetime64', 'datetime64[ns]']).columns:
        df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d')
    
    # Replace NaN with None for better compression
    df_copy = df_copy.where(pd.notna(df_copy), None)
    
    # Pickle and compress
    pickled = pickle.dumps(df_copy, protocol=pickle.HIGHEST_PROTOCOL)
    compressed = gzip.compress(pickled, compresslevel=6)
    return compressed

def decompress_dataframe(compressed_bytes):
    """Decompress bytes back to dataframe"""
    pickled = gzip.decompress(compressed_bytes)
    df = pickle.loads(pickled)
    
    # Convert string dates back to datetime
    for col in df.columns:
        if col in ['BOOKING_DATE', 'MATURITY_DATE']:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    return df

def compute_ftp_components(deposit, loan, tenure):
    """Helper to compute FTP charge, gain, net"""
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

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle Excel file upload, process and compress dataframes"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
    
    try:
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
        
        # Store summaries and previews
        summaries = {'ZWG': {}, 'FX': {}}
        sheets_preview = {}
        compressed_dfs = {}
        
        # Create Excel writer for output
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Process each sheet
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
                    df_processed = process_loans_sheet(df, sheet, first_day, last_day, summaries)
                    
                    # COMPRESS the dataframe (60-70% smaller!)
                    compressed_bytes = compress_dataframe(df_processed)
                    compressed_dfs[sheet] = compressed_bytes
                    
                    print(f"  📦 Compression: {len(df_processed):,} rows -> {len(compressed_bytes)/1024:.1f} KB ({(len(compressed_bytes)/(len(df_processed)*100)):.1f} bytes/row)")
                    
                    # Prepare for Excel - convert datetime to string
                    for col in df_processed.select_dtypes(include=['datetime64', 'datetime64[ns]']).columns:
                        df_processed[col] = df_processed[col].dt.strftime('%Y-%m-%d')
                    
                    df_processed = df_processed.fillna('')
                    
                    # Write to Excel
                    sheet_name_clean = sheet[:31]
                    df_processed.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                    
                    # Store preview (first 100 rows)
                    preview_data = df_processed.head(100).copy()
                    sheets_preview[sheet] = {
                        'columns': df_processed.columns.tolist(),
                        'data': preview_data.to_dict(orient='records'),
                        'shape': df_processed.shape
                    }
                    
                    del df_processed
                    
                else:
                    # For non-loan sheets, write directly
                    df = df.fillna('')
                    sheet_name_clean = sheet[:31]
                    df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                    
                    # Compress non-loan sheets too
                    compressed_bytes = compress_dataframe(df)
                    compressed_dfs[sheet] = compressed_bytes
                    
                    # Store preview
                    preview_data = df.head(100).copy()
                    sheets_preview[sheet] = {
                        'columns': df.columns.tolist(),
                        'data': preview_data.to_dict(orient='records'),
                        'shape': original_shape
                    }
                
                del df  # Free memory immediately
                print(f"  Completed processing {sheet}")
            
            # Add summary sheet
            summary_data = []
            for currency, sheets_data in summaries.items():
                for sheet_name, sheet_data in sheets_data.items():
                    summary_data.append({
                        'Currency': currency,
                        'Sheet': sheet_name,
                        'Total Exposure': sheet_data['total_exposure'],
                        'Total FTP Charge': sheet_data['total_ftp_charge'],
                        'Number of Records': sheet_data['row_count']
                    })
                    
                    for sbu in sheet_data['by_sbu']:
                        summary_data.append({
                            'Currency': f"{currency} - {sheet_name}",
                            'Sheet': f"  {sbu['SBU']}",
                            'Total Exposure': sbu['Currency Exposure + Currency Accrued Reporting'],
                            'Total FTP Charge': sbu['FTP Charge'],
                            'Number of Records': ''
                        })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Store everything
        output.seek(0)
        excel_bytes = output.getvalue()
        
        latest_data['filename'] = file.filename
        latest_data['excel_file'] = excel_bytes
        latest_data['compressed_dataframes'] = compressed_dfs  # Store compressed data
        latest_data['summaries'] = summaries
        latest_data['sheets_preview'] = sheets_preview
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'year': year
        }
        
        total_compressed = sum(len(b) for b in compressed_dfs.values())
        print(f"✅ Success! Total compressed size: {total_compressed/1024:.1f} KB")
        print(f"📊 Excel file size: {len(excel_bytes)/1024:.2f} KB")
        
        return jsonify({
            'status': 'success',
            'message': f'Successfully processed {len(sheet_names)} sheets',
            'summary': summaries,
            'period': latest_data['period']
        })
    
    except Exception as e:
        print(f"Error processing upload: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def process_loans_sheet(df, sheet, first_day, last_day, summaries):
    """Process a loans sheet and return the processed dataframe"""
    df_processed = df.copy()
    
    # Branch mapping dictionary (your existing mapping - kept for brevity)
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
        
        if 'Source of Funding' in df_processed.columns:
            df_processed['Source of Funding'] = df_processed['Source of Funding'].astype(str).str.strip()
            loc_mask = df_processed['Source of Funding'].str.upper() == 'LOC'
            df_processed.loc[loc_mask, 'SBU'] = 'Corporate Banking'
        
        df_processed['ACC MANAGEMENT UNIT'] = df_processed['ACC MANAGEMENT UNIT'].fillna('Unknown')
        df_processed['SBU'] = df_processed['SBU'].fillna('Unknown')
    else:
        df_processed['ACC MANAGEMENT UNIT'] = 'Unknown'
        df_processed['SBU'] = 'Unknown'
    
    # Date processing
    if 'BOOKING_DATE' in df_processed.columns:
        df_processed['BOOKING_DATE'] = pd.to_datetime(df_processed['BOOKING_DATE'], errors='coerce')
        df_processed['BOOKING_DATE'] = df_processed['BOOKING_DATE'].fillna(first_day)
    
    if 'MATURITY_DATE' in df_processed.columns:
        df_processed['MATURITY_DATE'] = pd.to_datetime(df_processed['MATURITY_DATE'], errors='coerce')
        df_processed['MATURITY_DATE'] = df_processed['MATURITY_DATE'].fillna(first_day + timedelta(days=365))
    
    # Calculate TENOR
    if 'BOOKING_DATE' in df_processed.columns and 'MATURITY_DATE' in df_processed.columns:
        df_processed['TENOR'] = (df_processed['MATURITY_DATE'] - df_processed['BOOKING_DATE']).dt.days
        df_processed.loc[df_processed['TENOR'] < 0, 'TENOR'] = 0
    
    # Calculate DimDays
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
    tenors_list = [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800]
    bucket_labels = []
    for i in range(len(tenors_list)):
        if i == 0:
            bucket_labels.append(f'<{tenors_list[i]}days')
        else:
            bucket_labels.append(f'{tenors_list[i-1]}-{tenors_list[i]}days')
    bucket_labels.append(f'+{tenors_list[-1]}days')
    
    for label in bucket_labels:
        df_processed[label] = 0
    
    # Allocate to buckets
    exposure = df_processed['Currency Exposure + Currency Accrued Reporting']
    mtm_days = df_processed['MTM'] * 30
    bin_edges = [0] + tenors_list + [float('inf')]
    bucket_indices = pd.cut(mtm_days, bins=bin_edges, labels=False, right=False, include_lowest=True)
    bucket_indices = bucket_indices.fillna(len(bucket_labels) - 1).astype(int)
    
    for i, label in enumerate(bucket_labels):
        mask = (bucket_indices == i) & (exposure > 0)
        df_processed.loc[mask, label] = exposure.loc[mask]
    
    df_processed['Check'] = df_processed[bucket_labels].sum(axis=1)
    
    # Calculate FTP Charge
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
    
    # Store summary
    currency = 'ZWG' if sheet == "ZWG LOANS" else 'FX'
    sbu_summary = df_processed.groupby('SBU').agg({
        'Currency Exposure + Currency Accrued Reporting': 'sum',
        'FTP Charge': 'sum'
    }).reset_index()
    
    summaries[currency][sheet] = {
        'total_exposure': float(df_processed['Currency Exposure + Currency Accrued Reporting'].sum()),
        'total_ftp_charge': float(df_processed['FTP Charge'].sum()),
        'by_sbu': sbu_summary.to_dict(orient='records'),
        'row_count': len(df_processed)
    }
    
    return df_processed

@app.route('/download-excel', methods=['GET'])
def download_excel():
    """Download the pre-generated Excel file (fast, no decompression needed)"""
    try:
        if latest_data.get('excel_file') is None:
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404
        
        month = latest_data.get('period', {}).get('month', 'Report')
        year = latest_data.get('period', {}).get('year', '')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"FTP_Results_{month}_{year}_{timestamp}.xlsx"
        
        return send_file(
            io.BytesIO(latest_data['excel_file']),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Error downloading Excel: {str(e)}")
        return jsonify({'error': f'Failed to download Excel: {str(e)}'}), 500

@app.route('/get-full-data/<sheet_name>', methods=['GET'])
def get_full_data(sheet_name):
    """Get full dataframe for a specific sheet (decompress on demand)"""
    try:
        if sheet_name not in latest_data.get('compressed_dataframes', {}):
            return jsonify({'error': 'Sheet not found'}), 404
        
        # Decompress the dataframe
        df = decompress_dataframe(latest_data['compressed_dataframes'][sheet_name])
        
        # Return as JSON (first 100 rows for preview, but you can adjust)
        return jsonify({
            'sheet': sheet_name,
            'shape': df.shape,
            'data': df.head(100).to_dict(orient='records'),
            'columns': df.columns.tolist()
        })
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/get-preview', methods=['GET'])
def get_preview():
    """Return the current preview data (uploaded sheets)"""
    if latest_data.get('sheets_preview'):
        return jsonify({
            'filename': latest_data['filename'],
            'sheets': latest_data['sheets_preview']
        })
    else:
        return jsonify({'message': 'No data uploaded yet'}), 404

@app.route('/calculate', methods=['POST'])
def calculate():
    """Calculate FTP based on inputs and return results"""
    data = request.json
    deposit = data.get('deposit')
    loan = data.get('loan')
    tenure = data.get('tenure')
    
    result = compute_ftp_components(deposit, loan, tenure)
    
    return jsonify(result)

@app.route('/ftp-curve-data', methods=['GET'])
def ftp_curve_data():
    """Return FTP curve data for USD and ZWG"""
    return jsonify({
        'tenors': tenors,
        'zwg': {
            'name': 'ZWG FTP Curve',
            'rates': zwg_rates,
            'color': '#b33a3a',
            'borderColor': '#921f1f'
        },
        'usd': {
            'name': 'USD FTP Curve',
            'rates': usd_rates,
            'color': '#2563eb',
            'borderColor': '#1d4ed8'
        }
    })

if __name__ == '__main__':
    app.run(debug=True, port=5000)