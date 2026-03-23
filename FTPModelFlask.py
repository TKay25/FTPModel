from flask import Flask, request, jsonify, send_file
import pandas as pd
import io
import openpyxl
import numpy as np
from datetime import datetime, timedelta
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import portrait, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER

app = Flask(__name__)

# Define the tenor points (in days)
tenors = [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800]

# USD FTP curve data
usd_rates = [3.29, 3.36, 6.15, 10.97, 11.02, 11.13, 12.22, 12.22, 12.22, 13.96, 15.41, 18.32, 18.32]

# ZWG FTP curve data
zwg_rates = [16.90, 16.90, 16.90, 16.90, 17.90, 18.10, 19.10, 19.10, 20.10, 23.47, 26.54, 32.67, 32.67]

# Global variable
latest_data = {
    'filename': None,
    'sheets': {},
    'ftp_results': None,
    'summaries': {},
    'period': {}
}

def format_number(num):
    if num is None:
        return '0'
    if abs(num) >= 1000000:
        return f"{num/1000000:.2f}M"
    if abs(num) >= 1000:
        return f"{num/1000:.2f}K"
    return f"{num:.2f}"

def generate_pdf_report():
    """Generate PDF report from the preview data and summaries"""
    buffer = io.BytesIO()
    
    global latest_data
    
    if not latest_data.get('summaries') or not latest_data.get('sheets'):
        doc = SimpleDocTemplate(buffer, pagesize=portrait(A4))
        styles = getSampleStyleSheet()
        story = []
        story.append(Paragraph("No data available", styles['Title']))
        story.append(Paragraph("Please upload a file first.", styles['Normal']))
        doc.build(story)
        buffer.seek(0)
        return buffer
    
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4),
                           rightMargin=72, leftMargin=72,
                           topMargin=72, bottomMargin=72)
    
    styles = getSampleStyleSheet()
    story = []
    
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24,
                                  textColor=colors.HexColor('#921f1f'), alignment=TA_CENTER, spaceAfter=30)
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Heading2'], fontSize=16,
                                   textColor=colors.HexColor('#632424'), spaceAfter=12, spaceBefore=12)
    
    title = Paragraph("FTP Analysis Report", title_style)
    story.append(title)
    
    period_text = f"{latest_data['period'].get('first_day', 'N/A')} to {latest_data['period'].get('last_day', 'N/A')}"
    period_style = ParagraphStyle('PeriodStyle', parent=styles['Normal'], fontSize=14,
                                   alignment=TA_CENTER, textColor=colors.HexColor('#632424'), spaceAfter=20)
    story.append(Paragraph(period_text, period_style))
    story.append(Spacer(1, 20))
    
    # Summary by Currency
    story.append(Paragraph("FTP Summary by Currency", header_style))
    
    table_data = [['Currency', 'Sheet', 'Total Exposure', 'Total FTP Charge', 'Records']]
    for currency, sheets in latest_data['summaries'].items():
        for sheet_name, sheet_data in sheets.items():
            table_data.append([
                currency, sheet_name,
                format_number(sheet_data['total_exposure']),
                format_number(sheet_data['total_ftp_charge']),
                f"{sheet_data['row_count']:,}"
            ])
    
    table = Table(table_data, colWidths=[1.2*inch, 1.8*inch, 1.5*inch, 1.5*inch, 1*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#921f1f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
    ]))
    story.append(table)
    story.append(Spacer(1, 20))
    
    # SBU Breakdown
    story.append(Paragraph("SBU Breakdown by Currency", header_style))
    for currency, sheets in latest_data['summaries'].items():
        story.append(Paragraph(f"<b>{currency}</b>", styles['Normal']))
        sbu_data = [['SBU', 'Exposure', 'FTP Charge']]
        for sheet_data in sheets.values():
            if sheet_data.get('by_sbu'):
                for sbu in sheet_data['by_sbu']:
                    sbu_data.append([
                        sbu['SBU'],
                        format_number(sbu['Currency Exposure + Currency Accrued Reporting']),
                        format_number(sbu['FTP Charge'])
                    ])
        if len(sbu_data) > 1:
            sbu_table = Table(sbu_data, colWidths=[2*inch, 2*inch, 2*inch])
            sbu_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#632424')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ]))
            story.append(sbu_table)
            story.append(Spacer(1, 15))
    
    # Grand Totals
    total_exposure_zwg = sum(sheet_data['total_exposure'] for currency, sheets in latest_data['summaries'].items() 
                              for sheet_data in sheets.values() if currency == 'ZWG')
    total_ftp_zwg = sum(sheet_data['total_ftp_charge'] for currency, sheets in latest_data['summaries'].items() 
                         for sheet_data in sheets.values() if currency == 'ZWG')
    total_exposure_fx = sum(sheet_data['total_exposure'] for currency, sheets in latest_data['summaries'].items() 
                             for sheet_data in sheets.values() if currency == 'FX')
    total_ftp_fx = sum(sheet_data['total_ftp_charge'] for currency, sheets in latest_data['summaries'].items() 
                        for sheet_data in sheets.values() if currency == 'FX')
    
    story.append(Paragraph("Grand Totals", header_style))
    totals_data = [['Currency', 'Total Exposure', 'Total FTP Charge']]
    if total_exposure_zwg > 0:
        totals_data.append(['ZWG', format_number(total_exposure_zwg), format_number(total_ftp_zwg)])
    if total_exposure_fx > 0:
        totals_data.append(['USD', format_number(total_exposure_fx), format_number(total_ftp_fx)])
    
    totals_table = Table(totals_data, colWidths=[2*inch, 2.5*inch, 2.5*inch])
    totals_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#921f1f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ]))
    story.append(totals_table)
    story.append(Spacer(1, 20))
    
    # Preview Data Section
    story.append(PageBreak())
    story.append(Paragraph("Data Preview (First 10 Rows per Sheet)", header_style))
    
    for sheet_name, sheet_data in latest_data['sheets'].items():
        story.append(Paragraph(f"<b>{sheet_name}</b>", styles['Normal']))
        story.append(Spacer(1, 5))
        
        preview_rows = sheet_data.get('data', [])[:10]
        if preview_rows and sheet_data.get('columns'):
            columns = sheet_data['columns'][:5]
            preview_table_data = [columns[:5]]
            for row in preview_rows:
                row_data = []
                for col in columns[:5]:
                    val = row.get(col, '')
                    if isinstance(val, float):
                        row_data.append(format_number(val))
                    elif len(str(val)) > 20:
                        row_data.append(str(val)[:17] + '...')
                    else:
                        row_data.append(str(val))
                preview_table_data.append(row_data)
            
            col_width = (6.5 * inch) / len(columns[:5])
            preview_table = Table(preview_table_data, colWidths=[col_width] * len(columns[:5]))
            preview_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#632424')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTSIZE', (0, 1), (-1, -1), 7),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ]))
            story.append(preview_table)
            story.append(Spacer(1, 15))
    
    story.append(Spacer(1, 30))
    footer_text = f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    footer_style = ParagraphStyle('Footer', parent=styles['Normal'], fontSize=8,
                                   alignment=TA_CENTER, textColor=colors.grey)
    story.append(Paragraph(footer_text, footer_style))
    story.append(Paragraph("FTP Central System", footer_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

@app.route('/download-pdf', methods=['GET'])
def download_pdf():
    try:
        global latest_data
        if not latest_data.get('summaries') or not latest_data.get('period'):
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404
        pdf_buffer = generate_pdf_report()
        month = latest_data.get('period', {}).get('month', 'Report')
        year = latest_data.get('period', {}).get('year', '')
        filename = f"FTP_Report_{month}_{year}.pdf"
        return send_file(pdf_buffer, as_attachment=True, download_name=filename, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Failed to download PDF: {str(e)}'}), 500

def compute_ftp_components(deposit, loan, tenure):
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
        return {'charge': f"{charge:.2f}", 'gain': f"{gain:.2f}", 'net': f"{net:.2f}"}
    except:
        return {'charge': '1.12', 'gain': '0.72', 'net': '1.84'}

@app.route('/')
def index():
    from flask import send_from_directory
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    global latest_data
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    try:
        # Parse filename
        filename_match = re.search(r'FTP Input File (\w+) (\d{4})', file.filename)
        if not filename_match:
            return jsonify({'error': 'Filename must be: FTP Input File Month Year.xlsx'}), 400
        
        month_name = filename_match.group(1)
        year = int(filename_match.group(2))
        
        month_map = {'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
                     'july':7,'august':8,'september':9,'october':10,'november':11,'december':12}
        month_num = month_map.get(month_name.lower())
        if not month_num:
            return jsonify({'error': f'Invalid month: {month_name}'}), 400
        
        first_day = datetime(year, month_num, 1)
        if month_num == 12:
            last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(year, month_num + 1, 1) - timedelta(days=1)
        
        # Get sheet names
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        print(f"Found sheets: {sheet_names}")
        
        sheets_data = {}
        global_summaries = {'ZWG': {}, 'FX': {}}
        
        # Your FULL branch mapping dictionary here (keep your complete one)
        branch_sbu_map = {
            '106': {'unit': 'Agribusiness', 'sbu': 'Corporate Banking'},
            '108': {'unit': 'Business Banking', 'sbu': 'Corporate Banking'},
            '47': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
            '11': {'unit': 'Kwame Nkrumah', 'sbu': 'Retail Banking'},
            '12': {'unit': '8Th Avenue', 'sbu': 'Retail Banking'},
            # ... ADD YOUR FULL MAPPING HERE (all 200+ entries)
        }
        
        for sheet in sheet_names:
            print(f"Processing: {sheet}")
            try:
                df = pd.read_excel(file, sheet_name=sheet)
            except Exception as e:
                print(f"Error: {e}")
                continue
            
            if sheet in ["ZWG LOANS", "FX LOANS"]:
                df_processed = df.copy()
                del df
                
                # Add SBU column
                branch_col = None
                for col in ['Branch Code', 'BRANCHCODE', 'BRANCH_CODE']:
                    if col in df_processed.columns:
                        branch_col = col
                        break
                
                if branch_col:
                    df_processed[branch_col] = df_processed[branch_col].astype(str).str.strip()
                    branch_info = df_processed[branch_col].map(branch_sbu_map)
                    df_processed['SBU'] = branch_info.apply(lambda x: x['sbu'] if isinstance(x, dict) else 'Unknown')
                else:
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
                def calc_days(row):
                    bd = row['BOOKING_DATE'].date() if hasattr(row['BOOKING_DATE'], 'date') else row['BOOKING_DATE']
                    md = row['MATURITY_DATE'].date() if hasattr(row['MATURITY_DATE'], 'date') else row['MATURITY_DATE']
                    if bd <= first_day.date() and md >= last_day.date():
                        return (last_day.date() - first_day.date()).days + 1
                    elif bd >= first_day.date() and md >= last_day.date():
                        return (last_day.date() - bd).days + 1
                    elif bd >= first_day.date() and md <= last_day.date():
                        return (md - bd).days
                    else:
                        return (md - first_day.date()).days
                
                df_processed['DimDays'] = df_processed.apply(calc_days, axis=1)
                
                # Calculate DTM and MTM
                last_day_ts = pd.Timestamp(last_day.date())
                df_processed['DTM'] = df_processed.apply(lambda r: (r['MATURITY_DATE'] - last_day_ts).days if r['MATURITY_DATE'] > last_day_ts else 0, axis=1)
                df_processed['MTM'] = (df_processed['DTM'] / 30).round(1)
                
                # Bucket columns
                bucket_labels = ['<7days', '7-14days', '14-21days', '21-30days', '30-60days', '60-90days', '90-180days', '180-270days', '270-360days', '360-720days', '720-1080days', '1080-1460days', '1460-1800days', '+1800days']
                for label in bucket_labels:
                    df_processed[label] = 0
                
                # Allocate to buckets
                exposure = df_processed['Currency Exposure + Currency Accrued Reporting']
                mtm_days = df_processed['MTM'] * 30
                tenors_list = [7,14,21,30,60,90,180,270,360,720,1080,1460,1800]
                bin_edges = [0] + tenors_list + [float('inf')]
                bucket_indices = pd.cut(mtm_days, bins=bin_edges, labels=False, right=False, include_lowest=True)
                bucket_indices = bucket_indices.fillna(len(bucket_labels)-1).astype(int)
                
                for i, label in enumerate(bucket_labels):
                    mask = (bucket_indices == i) & (exposure > 0)
                    df_processed.loc[mask, label] = exposure.loc[mask]
                
                df_processed['Check'] = df_processed[bucket_labels].sum(axis=1)
                
                # FTP Charge
                rates = zwg_rates if sheet == "ZWG LOANS" else usd_rates
                rate_array = np.zeros(len(df_processed))
                for i, label in enumerate(bucket_labels):
                    rate = rates[i] if i < len(rates) else rates[-1]
                    rate_array += df_processed[label] * (rate / 100)
                df_processed['FTP Charge'] = rate_array
                
                # Summary
                currency = 'ZWG' if sheet == "ZWG LOANS" else 'FX'
                sbu_summary = df_processed.groupby('SBU').agg({
                    'Currency Exposure + Currency Accrued Reporting': 'sum',
                    'FTP Charge': 'sum'
                }).reset_index()
                
                global_summaries[currency][sheet] = {
                    'total_exposure': float(df_processed['Currency Exposure + Currency Accrued Reporting'].sum()),
                    'total_ftp_charge': float(df_processed['FTP Charge'].sum()),
                    'by_sbu': sbu_summary.to_dict(orient='records'),
                    'row_count': len(df_processed)
                }
                
                # Preview
                preview = df_processed.head(100).copy()
                for col in preview.select_dtypes(include=['datetime64']).columns:
                    preview[col] = preview[col].astype(str).replace('NaT', None)
                
                sheets_data[sheet] = {
                    'columns': df_processed.columns.tolist(),
                    'data': preview.to_dict(orient='records'),
                    'shape': df_processed.shape
                }
                del df_processed
                
            else:
                preview = df.head(100).copy()
                sheets_data[sheet] = {
                    'columns': df.columns.tolist(),
                    'data': preview.to_dict(orient='records'),
                    'shape': df.shape
                }
                del df
            
            print(f"Completed: {sheet}")
        
        # Store data
        latest_data['filename'] = file.filename
        latest_data['sheets'] = sheets_data
        latest_data['summaries'] = global_summaries
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'year': year
        }
        
        print(f"✅ Stored summaries: {list(global_summaries.keys())}")
        
        return jsonify({
            'status': 'success',
            'summary': global_summaries,
            'period': latest_data['period']
        })
    
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/calculate', methods=['POST'])
def calculate():
    data = request.json
    result = compute_ftp_components(data.get('deposit'), data.get('loan'), data.get('tenure'))
    return jsonify(result)

@app.route('/ftp-curve-data', methods=['GET'])
def ftp_curve_data():
    return jsonify({
        'tenors': tenors,
        'zwg': {'name': 'ZWG FTP Curve', 'rates': zwg_rates, 'color': '#b33a3a', 'borderColor': '#921f1f'},
        'usd': {'name': 'USD FTP Curve', 'rates': usd_rates, 'color': '#2563eb', 'borderColor': '#1d4ed8'}
    })

@app.route('/get-preview', methods=['GET'])
def get_preview():
    if latest_data['sheets']:
        return jsonify({'filename': latest_data['filename'], 'sheets': latest_data['sheets']})
    return jsonify({'message': 'No data uploaded yet'}), 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)