from flask import Flask, request, jsonify, send_file
import pandas as pd
import io
import openpyxl
from datetime import datetime

app = Flask(__name__)

# Global variable to store the latest uploaded data (in-memory, no database)
latest_data = {
    'filename': None,
    'sheets': {},  # sheet name -> DataFrame as dict
    'ftp_results': None  # store latest calculated results
}


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
        # Extract month and year from filename
        import re
        from datetime import datetime, timedelta
        import pandas as pd
        import datetime as dt
        
        # Memory optimization: Limit rows for large files
        # Instead of reading entire file at once, read in chunks or limit rows
        # For now, we'll read but then only store preview
        
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
        
        # Get first day of the month
        first_day = datetime(year, month_num, 1)
        
        # Get last day of the month
        if month_num == 12:
            last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(year, month_num + 1, 1) - timedelta(days=1)
        
        # Read all sheet names first (fast)
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        print(f"Found sheets: {sheet_names}")
        
        # Dictionary to store processed data for each sheet
        sheets_data = {}
        
        # Process each sheet separately
        for sheet in sheet_names:
            print(f"Processing sheet: {sheet}")
            
            # Read the sheet - use dtypes to optimize memory
            # Specify dtypes for known columns to reduce memory
            dtype_spec = {}
            if sheet in ["ZWG LOANS", "FX LOANS"]:
                # Only specify dtypes for columns we know
                dtype_spec = {
                    'Branch Code': 'str',
                    'CURRENCY': 'str',
                    'Loan Type': 'str',
                    'Staff Status': 'str'
                }
            
            try:
                df = pd.read_excel(file, sheet_name=sheet, dtype=dtype_spec)
            except Exception as e:
                print(f"  Error reading sheet {sheet}: {e}")
                # Try without dtype specification
                df = pd.read_excel(file, sheet_name=sheet)
            
            original_shape = df.shape
            print(f"  Original shape: {original_shape}")
            
            # Special handling for ZWG LOANS and FX LOANS sheets
            if sheet in ["ZWG LOANS", "FX LOANS"]:
                print(f"  Applying {sheet} special handling...")
                
                # Process in chunks if needed - but for now, keep full processing
                df_processed = df.copy()
                
                # Clear original df to free memory
                del df
                
                # --- BRANCH MAPPING (using dictionary - already efficient) ---
                branch_map = {
                    '106': 'Agribusiness', '118': 'Bureau De Change Hre', '45': 'Business Banking',
                    '108': 'Business Banking', '47': 'Private Sector', '113': 'Custodial Services',
                    '48': 'Private Sector', '53': 'Private Sector', '602': 'Mortgage Finance',
                    '107': 'Institutional Banking', '66': 'Treasury', '601': 'Treasury',
                    '0': 'Shared Services', '35': 'Shared Services', '36': 'Shared Services',
                    '37': 'Shared Services', '38': 'Shared Services', '39': 'Shared Services',
                    '40': 'Shared Services', '41': 'Shared Services', '43': 'Shared Services',
                    '46': 'Shared Services', '49': 'Shared Services', '50': 'Shared Services',
                    '54': 'Shared Services', '56': 'Shared Services', '57': 'Shared Services',
                    '58': 'Shared Services', '61': 'Shared Services', '65': 'Shared Services',
                    '67': 'Shared Services', '68': 'Shared Services', '69': 'Shared Services',
                    '70': 'Shared Services', '105': 'Shared Services', '115': 'Shared Services',
                    '117': 'Shared Services', '123': 'Shared Services', '124': 'Shared Services',
                    '600': 'Shared Services', '141': 'Shared Services', '116': 'Shared Services',
                    '11': 'Kwame Nkrumah', '12': '8Th Avenue', '13': 'Mutare', '14': 'Kwekwe',
                    '15': 'Chitungwiza', '17': 'Gokwe', '18': 'Gweru', '20': 'Chivhu',
                    '21': 'Selous', '23': 'Southerton', '24': 'Sapphire', '25': 'Masvingo',
                    '26': 'Belmont', '27': 'Cash Depot Bulawayo', '28': 'Chiredzi', '29': 'Borrowdale',
                    '30': 'Avondale', '31': 'Chinhoyi', '32': 'Kwekwe', '33': 'Sapphire',
                    '34': 'Cash Depot Harare', '44': 'Wealth Management', '87': 'Chipinge',
                    '88': '8Th Avenue', '89': 'Highfield', '90': 'Marondera', '91': 'Chitungwiza',
                    '92': 'Gokwe', '93': 'Beitbridge', '95': 'Kariba', '96': 'Kariba',
                    '97': 'Karoi', '98': 'Chinhoyi', '99': 'Masvingo', '100': 'Mvurwi',
                    '101': 'Chipinge', '102': 'Rusape', '103': 'Murehwa', '104': 'Victoria Falls',
                    '109': 'Chiredzi', '110': 'Selous', '111': 'Selous', '112': 'Mvurwi',
                    '51': 'Retail Head Office', '52': 'Kwame Nkrumah', '55': 'Shared Services',
                    '62': '8Th Avenue', '114': 'Retail Head Office', '120': 'Sapphire',
                    '121': 'Retail Centraslised Back Office', '122': 'Mta Centre Fife Street',
                    '611': 'Masvingo', '612': 'Chiredzi', '613': 'Masvingo', '614': 'Zvishavane',
                    '615': 'Gweru', '616': 'Kwekwe', '617': 'Kadoma', '618': 'Kadoma',
                    '619': 'Gokwe', '629': 'Chipinge', '630': 'Chipinge', '631': 'Mutare',
                    '632': 'Mutare', '633': 'Mutare', '634': 'Rusape', '644': '8Th Avenue',
                    '645': '8Th Avenue', '646': 'Belmont', '647': 'Belmont', '648': 'Belmont',
                    '649': 'Gwanda', '650': 'Cash Depot Bulawayo', '660': 'Samora Machel',
                    '661': 'Avondale', '662': 'Bindura', '663': 'Msasa', '664': 'Chinhoyi',
                    '665': 'Sapphire', '667': 'Karoi', '668': 'Murehwa', '669': 'Samora Machel',
                    '670': 'Samora Machel', '671': 'Cash Depot Harare', '672': 'Kariba',
                    '681': 'Sapphire', '682': 'Cripps', '683': 'Chitungwiza', '684': 'Chivhu',
                    '685': 'Sapphire', '686': 'Highfield', '687': 'Marondera', '688': 'Msasa',
                    '689': 'Msasa', '690': 'Sapphire', '125': 'Passport Centre Harare',
                    '127': 'Passport Centre Bulawayo', '126': 'Virtual Branch',
                    '128': 'Passport Centre Chitungwiza', '129': 'Passport Centre Lupane',
                    '130': 'Passport Centre Hwange', '131': 'Passport Centre Gweru',
                    '132': 'Passport Centre Beitbridge', '133': 'Passport Centre Chinhoyi',
                    '134': 'Passport Centre Marondera', '135': 'Passport Centre Bindura',
                    '136': 'Passport Centre Gwanda', '137': 'Passport Centre Mutare',
                    '138': 'Passport Centre Masvingo', '139': 'Passport Centre Zvishavane',
                    '140': 'Passport Centre Murehwa', '142': 'Retail Centralised Byo',
                    '145': 'Borrowdale', '146': 'Passport Centre Mwenezi', '200': 'Shared Services',
                    '147': 'Passport Centre Gokwe', '143': 'Retail Head Office', '144': 'Retail Head Office'
                }
                print(f"  Created branch mapping with {len(branch_map)} unique branch codes")
                
                # Find branch code column
                branch_code_col = None
                possible_names = ['Branch Code', 'BRANCHCODE', 'BRANCH_CODE', 'BRANCH', 'BR_CODE']
                for col in possible_names:
                    if col in df_processed.columns:
                        branch_code_col = col
                        break
                
                if branch_code_col:
                    # Convert to string and map
                    df_processed[branch_code_col] = df_processed[branch_code_col].astype(str).str.strip()
                    df_processed['ACC MANAGEMENT UNIT'] = df_processed[branch_code_col].map(branch_map)
                    df_processed['ACC MANAGEMENT UNIT'].fillna('Unknown', inplace=True)
                    
                    unknown_count = (df_processed['ACC MANAGEMENT UNIT'] == 'Unknown').sum()
                    print(f"  Added ACC MANAGEMENT UNIT column")
                    print(f"  Found {len(df_processed) - unknown_count} matching branch codes")
                    print(f"  {unknown_count} rows with unknown branch codes")
                    
                    unique_units = df_processed['ACC MANAGEMENT UNIT'].unique()
                    unit_counts = df_processed['ACC MANAGEMENT UNIT'].value_counts().to_dict()
                else:
                    print(f"  Warning: No branch code column found in {sheet}")
                    df_processed['ACC MANAGEMENT UNIT'] = 'Unknown'
                    unknown_count = len(df_processed)
                    unique_units = ['Unknown']
                    unit_counts = {'Unknown': unknown_count}
                
                # --- DATE PROCESSING (optimized) ---
                # Process BOOKING_DATE
                if 'BOOKING_DATE' in df_processed.columns:
                    df_processed['BOOKING_DATE'] = pd.to_datetime(df_processed['BOOKING_DATE'], errors='coerce')
                    booking_date_mask = df_processed['BOOKING_DATE'].isna()
                    df_processed.loc[booking_date_mask, 'BOOKING_DATE'] = first_day
                    print(f"  Updated {booking_date_mask.sum()} rows with BOOKING_DATE = {first_day.strftime('%Y-%m-%d')}")
                else:
                    booking_date_mask = pd.Series([False] * len(df_processed))
                
                # Process MATURITY_DATE with time object handling
                if 'MATURITY_DATE' in df_processed.columns:
                    # Check for time objects
                    time_objects_mask = df_processed['MATURITY_DATE'].apply(lambda x: isinstance(x, dt.time))
                    if time_objects_mask.any():
                        print(f"  Found {time_objects_mask.sum()} time objects in MATURITY_DATE column")
                        df_processed.loc[time_objects_mask, 'MATURITY_DATE'] = first_day
                    
                    df_processed['MATURITY_DATE'] = pd.to_datetime(df_processed['MATURITY_DATE'], errors='coerce')
                    maturity_date_mask = df_processed['MATURITY_DATE'].isna()
                    maturity_default = first_day + timedelta(days=365)
                    df_processed.loc[maturity_date_mask, 'MATURITY_DATE'] = maturity_default
                    print(f"  Updated {maturity_date_mask.sum()} rows with MATURITY_DATE = {maturity_default.strftime('%Y-%m-%d')}")
                else:
                    maturity_date_mask = pd.Series([False] * len(df_processed))
                
                # Create TENOR column
                if 'BOOKING_DATE' in df_processed.columns and 'MATURITY_DATE' in df_processed.columns:
                    df_processed['TENOR'] = (df_processed['MATURITY_DATE'] - df_processed['BOOKING_DATE']).dt.days
                    df_processed.loc[df_processed['TENOR'] < 0, 'TENOR'] = 0
                    print(f"  Created TENOR column")
                    
                    # Only create formatted version for preview (not full column to save memory)
                    def format_tenor(days):
                        if pd.isna(days) or days < 0:
                            return 'N/A'
                        if days < 30:
                            return f"{int(days)}D"
                        elif days < 365:
                            return f"{round(days / 30)}M"
                        else:
                            return f"{round(days / 365, 1)}Y"
                    
                    df_processed['TENOR_FORMATTED'] = df_processed['TENOR'].apply(format_tenor)
                else:
                    df_processed['TENOR'] = 0
                    df_processed['TENOR_FORMATTED'] = 'N/A'
                
                # Store only preview data (first 100 rows) instead of full data to save memory
                preview_data = df_processed.head(100).to_dict(orient='records')
                
                # Store processed data for preview
                sheets_data[sheet] = {
                    'columns': df_processed.columns.tolist(),
                    'data': preview_data,  # Only first 100 rows for preview
                    'shape': df_processed.shape,
                    'processed': True,
                    'booking_date_updates': int(booking_date_mask.sum()),
                    'maturity_date_updates': int(maturity_date_mask.sum()),
                    'branch_code_column': branch_code_col if branch_code_col else 'Not found',
                    'unknown_branch_codes': int(unknown_count),
                    'acc_management_units': list(unique_units),
                    'unit_counts': unit_counts,
                    'tenor_stats': {
                        'min': int(df_processed['TENOR'].min()),
                        'max': int(df_processed['TENOR'].max()),
                        'avg': float(df_processed['TENOR'].mean())
                    },
                    'period': {
                        'first_day': first_day.strftime('%d %B %Y'),
                        'last_day': last_day.strftime('%d %B %Y')
                    }
                }
                
                # Store full dataframe only if needed (remove this if not used elsewhere)
                # latest_data[f'{sheet.lower().replace(" ", "_")}_processed'] = df_processed
                # Instead, store only necessary summary stats to save memory
                latest_data[f'{sheet.lower().replace(" ", "_")}_summary'] = {
                    'row_count': len(df_processed),
                    'columns': df_processed.columns.tolist(),
                    'tenor_stats': {
                        'min': int(df_processed['TENOR'].min()),
                        'max': int(df_processed['TENOR'].max()),
                        'avg': float(df_processed['TENOR'].mean())
                    },
                    'acc_management_units': unit_counts
                }
                
                # Free up memory
                del df_processed
                
                print(f"  Completed processing {sheet}")
                
            else:
                # For other sheets, store only preview
                sheets_data[sheet] = {
                    'columns': df.columns.tolist(),
                    'data': df.head(100).to_dict(orient='records'),
                    'shape': df.shape,
                    'processed': False
                }
                del df  # Free memory
            
            print(f"  Completed processing {sheet}")
        
        # Store only necessary data in global variable
        latest_data['filename'] = file.filename
        latest_data['sheets'] = sheets_data
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'year': year
        }
        
        return jsonify({
            'success': True,
            'filename': file.filename,
            'sheets': sheets_data,
            'period': latest_data['period'],
            'message': f'Successfully loaded {len(sheet_names)} sheet(s)'
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
    
    # Define the tenor points (in days)
    tenors = [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800]
    
    # USD FTP curve data (matching your provided values)
    usd_rates = [3.29, 3.36, 6.15, 10.97, 11.02, 11.13, 12.22, 12.22, 12.22, 13.96, 15.41, 18.32, 18.32]
    
    # ZWG FTP curve data (matching your provided values)
    zwg_rates = [16.90, 16.90, 16.90, 16.90, 17.90, 18.10, 19.10, 19.10, 20.10, 23.47, 26.54, 32.67, 32.67]
    
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