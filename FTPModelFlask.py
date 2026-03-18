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
        # Read all sheets from Excel
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        sheets_data = {}
        
        # Store first 10 rows of each sheet as preview
        for sheet in sheet_names:
            df = pd.read_excel(file, sheet_name=sheet)
            # Convert to dict with records and limit to 10 rows for preview
            sheets_data[sheet] = {
                'columns': df.columns.tolist(),
                'data': df.head(10).to_dict(orient='records'),
                'shape': df.shape
            }
        
        # Store in global variable
        latest_data['filename'] = file.filename
        latest_data['sheets'] = sheets_data
        
        return jsonify({
            'success': True,
            'filename': file.filename,
            'sheets': sheets_data,
            'message': f'Successfully loaded {len(sheet_names)} sheet(s)'
        })
    
    except Exception as e:
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