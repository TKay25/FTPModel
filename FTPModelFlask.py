from flask import Flask, request, jsonify, send_file
import pandas as pd
import io
import openpyxl
import numpy as np
from datetime import datetime, timedelta
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')

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
    'full_dataframes': {},
    'summaries': {},
    'period': {}
}

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

def format_number(num):
    """Format large numbers with K/M suffix"""
    if num is None:
        return '0'
    if abs(num) >= 1000000:
        return f"{num/1000000:.2f}M"
    if abs(num) >= 1000:
        return f"{num/1000:.2f}K"
    return f"{num:.2f}"

def create_ftp_chart():
    """Create FTP curve chart as image"""
    fig, ax = plt.subplots(figsize=(10, 6))
    
    tenor_labels = []
    for t in tenors:
        if t <= 30:
            tenor_labels.append(f'{t}D')
        elif t == 60:
            tenor_labels.append('60D')
        elif t == 90:
            tenor_labels.append('90D')
        elif t == 180:
            tenor_labels.append('6M')
        elif t == 270:
            tenor_labels.append('9M')
        elif t == 360:
            tenor_labels.append('1Y')
        elif t == 720:
            tenor_labels.append('2Y')
        elif t == 1080:
            tenor_labels.append('3Y')
        elif t == 1460:
            tenor_labels.append('4Y')
        elif t == 1800:
            tenor_labels.append('5Y')
        else:
            tenor_labels.append(f'{t}D')
    
    ax.plot(tenor_labels, zwg_rates, 'o-', label='ZWG FTP Curve', color='#b33a3a', linewidth=2, markersize=6)
    ax.plot(tenor_labels, usd_rates, 's-', label='USD FTP Curve', color='#2563eb', linewidth=2, markersize=6)
    
    ax.set_xlabel('Tenor', fontsize=12, fontweight='bold')
    ax.set_ylabel('FTP Rate (%)', fontsize=12, fontweight='bold')
    ax.set_title('FTP Curves - ZWG vs USD', fontsize=14, fontweight='bold', pad=20)
    ax.legend(loc='best', fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_xticklabels(tenor_labels, rotation=45, ha='right')
    
    plt.tight_layout()
    
    # Save to bytes
    img_bytes = io.BytesIO()
    plt.savefig(img_bytes, format='png', dpi=150, bbox_inches='tight')
    plt.close()
    img_bytes.seek(0)
    return img_bytes

def generate_pdf_report():
    """Generate PDF report from the preview data and summaries"""
    buffer = io.BytesIO()
    
    # Create PDF document in portrait mode
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4),
                           rightMargin=72, leftMargin=72,
                           topMargin=72, bottomMargin=72)
    
    styles = getSampleStyleSheet()
    story = []
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#921f1f'),
        alignment=TA_CENTER,
        spaceAfter=30
    )
    
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=colors.HexColor('#632424'),
        spaceAfter=12,
        spaceBefore=12
    )
    
    # Title
    title = Paragraph("FTP Analysis Report", title_style)
    story.append(title)
    
    # Period info
    period_text = f"{latest_data['period'].get('first_day', 'N/A')} to {latest_data['period'].get('last_day', 'N/A')}"
    period_style = ParagraphStyle(
        'PeriodStyle',
        parent=styles['Normal'],
        fontSize=14,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#632424'),
        spaceAfter=20
    )
    story.append(Paragraph(period_text, period_style))
    
    # File info
    file_info = f"Source File: {latest_data.get('filename', 'N/A')}"
    story.append(Paragraph(file_info, styles['Normal']))
    story.append(Spacer(1, 20))
    
    # Add FTP Curve Chart
    story.append(Paragraph("FTP Curve Analysis", header_style))
    chart_img = create_ftp_chart()
    img = Image(chart_img, width=6*inch, height=3.6*inch)
    story.append(img)
    story.append(Spacer(1, 20))
    
    # Summary by Currency
    story.append(Paragraph("FTP Summary by Currency", header_style))
    
    # Create summary table data
    table_data = [['Currency', 'Sheet', 'Total Exposure', 'Total FTP Charge', 'Records']]
    
    for currency, sheets in latest_data['summaries'].items():
        for sheet_name, sheet_data in sheets.items():
            table_data.append([
                currency,
                sheet_name,
                format_number(sheet_data['total_exposure']),
                format_number(sheet_data['total_ftp_charge']),
                f"{sheet_data['row_count']:,}"
            ])
    
    # Create table
    table = Table(table_data, colWidths=[1.2*inch, 1.8*inch, 1.5*inch, 1.5*inch, 1*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#921f1f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
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
        
        # Get preview data (first 10 rows)
        preview_rows = sheet_data.get('data', [])[:10]
        
        if preview_rows and sheet_data.get('columns'):
            columns = sheet_data['columns'][:6]  # Limit to first 6 columns for readability
            
            # Create table data
            preview_table_data = [columns[:6]]
            
            for row in preview_rows:
                row_data = []
                for col in columns[:6]:
                    val = row.get(col, '')
                    if isinstance(val, float):
                        row_data.append(format_number(val))
                    elif len(str(val)) > 20:
                        row_data.append(str(val)[:17] + '...')
                    else:
                        row_data.append(str(val))
                preview_table_data.append(row_data)
            
            # Calculate column widths
            col_width = (6.5 * inch) / len(columns[:6])
            
            preview_table = Table(preview_table_data, colWidths=[col_width] * len(columns[:6]))
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
    
    # Footer
    story.append(Spacer(1, 30))
    footer_text = f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=8,
        alignment=TA_CENTER,
        textColor=colors.grey
    )
    story.append(Paragraph(footer_text, footer_style))
    story.append(Paragraph("FTP Central System", footer_style))
    
    # Build PDF
    doc.build(story)
    buffer.seek(0)
    return buffer

@app.route('/')
def index():
    """Serve the main HTML page"""
    from flask import send_from_directory
    return send_from_directory('.', 'index.html')

@app.route('/download-pdf', methods=['GET'])
def download_pdf():
    """Download the generated PDF report"""
    try:
        if not latest_data.get('summaries') or not latest_data.get('period'):
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404
        
        # Generate PDF
        pdf_buffer = generate_pdf_report()
        
        # Generate filename
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
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to download PDF: {str(e)}'}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle Excel file upload (same as your existing upload function)"""
    # [KEEP YOUR EXISTING UPLOAD FUNCTION HERE - IT REMAINS UNCHANGED]
    # I'm not repeating it here for brevity, but keep your full upload function
    pass

@app.route('/calculate', methods=['POST'])
def calculate():
    """Calculate FTP based on inputs and return results"""
    data = request.json
    deposit = data.get('deposit')
    loan = data.get('loan')
    tenure = data.get('tenure')
    
    result = compute_ftp_components(deposit, loan, tenure)
    
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