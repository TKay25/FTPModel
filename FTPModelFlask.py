from flask import Flask, request, jsonify, send_file
import json
import os
import sqlite3
import pandas as pd
import io
import openpyxl
import numpy as np
from datetime import datetime, timedelta
from time import perf_counter
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import portrait, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER

app = Flask(__name__)

CURVE_CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'curve_config.json')
PROCESSED_DATA_PATH = os.path.join(os.path.dirname(__file__), 'latest_processed_data.json')
PROCESSED_OUTPUTS_DIR = os.path.join(os.path.dirname(__file__), 'processed_outputs')
BRANCH_MAP_DB_PATH = os.path.join(os.path.dirname(__file__), 'branch_sbu_map.db')
os.makedirs(PROCESSED_OUTPUTS_DIR, exist_ok=True)


DEFAULT_BRANCH_SBU_MAP = {
    '0': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
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
    '35': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '36': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '37': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '38': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '39': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '40': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '41': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '43': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '44': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
    '45': {'unit': 'Business Banking', 'sbu': 'Corporate Banking'},
    '46': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '47': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
    '48': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
    '49': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '50': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '51': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
    '52': {'unit': 'Kwame Nkrumah', 'sbu': 'Retail Banking'},
    '53': {'unit': 'Private Sector', 'sbu': 'Corporate Banking'},
    '54': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '55': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '56': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '57': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '58': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '61': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '62': {'unit': '8th Avenue', 'sbu': 'Retail Banking'},
    '65': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '66': {'unit': 'Treasury', 'sbu': 'Treasury'},
    '67': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '68': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '69': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '70': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '87': {'unit': 'Chisipite', 'sbu': 'Retail Banking'},
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
    '105': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '106': {'unit': 'Agribusiness', 'sbu': 'Corporate Banking'},
    '107': {'unit': 'Institutional Banking', 'sbu': 'Corporate Banking'},
    '108': {'unit': 'Business Banking', 'sbu': 'Corporate Banking'},
    '109': {'unit': 'Chiredzi', 'sbu': 'Retail Banking'},
    '110': {'unit': 'Selous', 'sbu': 'Retail Banking'},
    '111': {'unit': 'Selous', 'sbu': 'Retail Banking'},
    '112': {'unit': 'Mvurwi', 'sbu': 'Retail Banking'},
    '113': {'unit': 'Custodial Services', 'sbu': 'Corporate Banking'},
    '114': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
    '115': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '116': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '117': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '118': {'unit': 'Bureau De Change Hire', 'sbu': 'Shared Services'},
    '120': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
    '121': {'unit': 'Retail Centralised', 'sbu': 'Retail Banking'},
    '122': {'unit': 'Mat Centre Fire Street', 'sbu': 'Retail Banking'},
    '123': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '124': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '125': {'unit': 'Passport Centre Harare', 'sbu': 'Retail Banking'},
    '126': {'unit': 'Virtual Branch', 'sbu': 'Retail Banking'},
    '127': {'unit': 'Passport Centre Harare', 'sbu': 'Retail Banking'},
    '128': {'unit': 'Passport Centre Chitungwiza', 'sbu': 'Retail Banking'},
    '129': {'unit': 'Passport Centre Lupar', 'sbu': 'Retail Banking'},
    '130': {'unit': 'Passport Centre Hwange', 'sbu': 'Retail Banking'},
    '131': {'unit': 'Passport Centre Gweru', 'sbu': 'Retail Banking'},
    '132': {'unit': 'Passport Centre Beitbridge', 'sbu': 'Retail Banking'},
    '133': {'unit': 'Passport Centre Chinhoyi', 'sbu': 'Retail Banking'},
    '134': {'unit': 'Passport Centre Marondera', 'sbu': 'Retail Banking'},
    '135': {'unit': 'Passport Centre Bindura', 'sbu': 'Retail Banking'},
    '136': {'unit': 'Passport Centre Gwe', 'sbu': 'Retail Banking'},
    '137': {'unit': 'Passport Centre Mutare', 'sbu': 'Retail Banking'},
    '138': {'unit': 'Passport Centre Mvurwi', 'sbu': 'Retail Banking'},
    '139': {'unit': 'Passport Centre Zvishavane', 'sbu': 'Retail Banking'},
    '140': {'unit': 'Passport Centre Murehwa', 'sbu': 'Retail Banking'},
    '141': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '142': {'unit': 'Retail Centralised', 'sbu': 'Retail Banking'},
    '143': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
    '144': {'unit': 'Retail Head Office', 'sbu': 'Retail Banking'},
    '145': {'unit': 'Borrowdale', 'sbu': 'Retail Banking'},
    '146': {'unit': 'Passport Centre Mwer', 'sbu': 'Retail Banking'},
    '147': {'unit': 'Passport Centre Gokwe', 'sbu': 'Retail Banking'},
    '203': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '600': {'unit': 'Shared Services', 'sbu': 'Shared Services'},
    '601': {'unit': 'Treasury', 'sbu': 'Treasury'},
    '602': {'unit': 'Mortgage Finance', 'sbu': 'Retail Banking'},
    '611': {'unit': 'Masvingo', 'sbu': 'Retail Banking'},
    '612': {'unit': 'Chiredzi', 'sbu': 'Retail Banking'},
    '613': {'unit': 'Chinhoyi', 'sbu': 'Retail Banking'},
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
    '644': {'unit': '8th Avenue', 'sbu': 'Retail Banking'},
    '645': {'unit': '8th Avenue', 'sbu': 'Retail Banking'},
    '646': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
    '647': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
    '648': {'unit': 'Belmont', 'sbu': 'Retail Banking'},
    '649': {'unit': 'Gwanda', 'sbu': 'Retail Banking'},
    '658': {'unit': 'Cash Depot Bulawayo', 'sbu': 'Retail Banking'},
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
    '682': {'unit': 'Chivhu', 'sbu': 'Retail Banking'},
    '683': {'unit': 'Chitungwiza', 'sbu': 'Retail Banking'},
    '684': {'unit': 'Chivhu', 'sbu': 'Retail Banking'},
    '685': {'unit': 'Sapphire', 'sbu': 'Retail Banking'},
    '686': {'unit': 'Highfield', 'sbu': 'Retail Banking'},
    '687': {'unit': 'Marondera', 'sbu': 'Retail Banking'},
    '688': {'unit': 'Msasa', 'sbu': 'Retail Banking'},
    '689': {'unit': 'Msasa', 'sbu': 'Retail Banking'},
    '690': {'unit': 'Sapphire', 'sbu': 'Retail Banking'}
}


def init_branch_map_db():
    with sqlite3.connect(BRANCH_MAP_DB_PATH) as connection:
        connection.execute(
            '''
            CREATE TABLE IF NOT EXISTS branch_sbu_map (
                branch_code TEXT PRIMARY KEY,
                unit TEXT NOT NULL,
                sbu TEXT NOT NULL
            )
            '''
        )
        existing_count = connection.execute('SELECT COUNT(*) FROM branch_sbu_map').fetchone()[0]
        if existing_count < len(DEFAULT_BRANCH_SBU_MAP):
            connection.execute('DELETE FROM branch_sbu_map')
            connection.executemany(
                'INSERT INTO branch_sbu_map (branch_code, unit, sbu) VALUES (?, ?, ?)',
                [(code, data['unit'], data['sbu']) for code, data in DEFAULT_BRANCH_SBU_MAP.items()]
            )
        connection.commit()


def load_branch_sbu_map():
    init_branch_map_db()
    with sqlite3.connect(BRANCH_MAP_DB_PATH) as connection:
        rows = connection.execute(
            'SELECT branch_code, unit, sbu FROM branch_sbu_map ORDER BY CAST(branch_code AS INTEGER), branch_code'
        ).fetchall()
    return {str(code): {'unit': unit, 'sbu': sbu} for code, unit, sbu in rows}


def load_branch_sbu_rows():
    init_branch_map_db()
    with sqlite3.connect(BRANCH_MAP_DB_PATH) as connection:
        rows = connection.execute(
            'SELECT branch_code, unit, sbu FROM branch_sbu_map ORDER BY CAST(branch_code AS INTEGER), branch_code'
        ).fetchall()
    return [{'code': str(code), 'unit': unit, 'sbu': sbu} for code, unit, sbu in rows]


def save_branch_sbu_rows(rows):
    normalized_rows = []
    seen_codes = set()

    for row in rows:
        code = str(row.get('code', '')).strip()
        unit = str(row.get('unit', '')).strip()
        sbu = str(row.get('sbu', '')).strip()

        if not code or not unit or not sbu:
            continue

        if code in seen_codes:
            continue

        seen_codes.add(code)
        normalized_rows.append((code, unit, sbu))

    if not normalized_rows:
        raise ValueError('At least one valid branch mapping is required')

    with sqlite3.connect(BRANCH_MAP_DB_PATH) as connection:
        connection.execute('DELETE FROM branch_sbu_map')
        connection.executemany(
            'INSERT INTO branch_sbu_map (branch_code, unit, sbu) VALUES (?, ?, ?)',
            normalized_rows
        )
        connection.commit()


def reset_branch_sbu_map_to_default():
    save_branch_sbu_rows([
        {'code': code, 'unit': data['unit'], 'sbu': data['sbu']}
        for code, data in DEFAULT_BRANCH_SBU_MAP.items()
    ])


init_branch_map_db()


def load_curve_config():
    default_curves = {
        'tenors': [7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800],
        'usd_rates': [3.29, 3.36, 6.15, 10.97, 11.02, 11.13, 12.22, 12.22, 12.22, 13.96, 15.41, 18.32, 18.32],
        'zwg_rates': [16.90, 16.90, 16.90, 16.90, 17.90, 18.10, 19.10, 19.10, 20.10, 23.47, 26.54, 32.67, 32.67]
    }

    if not os.path.exists(CURVE_CONFIG_PATH):
        return default_curves

    try:
        with open(CURVE_CONFIG_PATH, 'r', encoding='utf-8') as curve_file:
            curve_config = json.load(curve_file)

        tenors = [int(value) for value in curve_config.get('tenors', default_curves['tenors'])]
        usd_curve = [float(value) for value in curve_config.get('usd_rates', default_curves['usd_rates'])]
        zwg_curve = [float(value) for value in curve_config.get('zwg_rates', default_curves['zwg_rates'])]

        if len(tenors) != len(usd_curve) or len(tenors) != len(zwg_curve):
            raise ValueError('Curve config lengths do not match')

        return {
            'tenors': tenors,
            'usd_rates': usd_curve,
            'zwg_rates': zwg_curve
        }
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return default_curves


def save_curve_config(curve_config):
    with open(CURVE_CONFIG_PATH, 'w', encoding='utf-8') as curve_file:
        json.dump(curve_config, curve_file, indent=2)

curve_config = load_curve_config()

# Define the tenor points (in days)
tenors = curve_config['tenors']

# USD FTP curve data
usd_rates = curve_config['usd_rates']

# ZWG FTP curve data
zwg_rates = curve_config['zwg_rates']

# Global variable
latest_data = {
    'filename': None,
    'sheets': {},
    'ftp_results': None,
    'summaries': {},
    'period': {},
    'excel_output_path': None,
    'excel_filename': None
}


def save_latest_data_snapshot():
    def _json_default(value):
        if isinstance(value, np.generic):
            return value.item()
        if isinstance(value, (datetime, timedelta)):
            return str(value)
        return str(value)

    with open(PROCESSED_DATA_PATH, 'w', encoding='utf-8') as snapshot_file:
        json.dump(latest_data, snapshot_file, ensure_ascii=True, default=_json_default)


def load_latest_data_snapshot():
    global latest_data

    if not os.path.exists(PROCESSED_DATA_PATH):
        return False

    try:
        with open(PROCESSED_DATA_PATH, 'r', encoding='utf-8') as snapshot_file:
            snapshot = json.load(snapshot_file)

        latest_data = {
            'filename': snapshot.get('filename'),
            'sheets': snapshot.get('sheets', {}),
            'ftp_results': snapshot.get('ftp_results'),
            'summaries': snapshot.get('summaries', {}),
            'period': snapshot.get('period', {}),
            'excel_output_path': snapshot.get('excel_output_path'),
            'excel_filename': snapshot.get('excel_filename')
        }
        return True
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return False


def ensure_latest_data_available():
    if latest_data.get('summaries') and latest_data.get('period'):
        return True
    return load_latest_data_snapshot()

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
        if not ensure_latest_data_available():
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404
        pdf_buffer = generate_pdf_report()
        month = latest_data.get('period', {}).get('month', 'Report')
        year = latest_data.get('period', {}).get('year', '')
        filename = f"FTP_Report_{month}_{year}.pdf"
        return send_file(pdf_buffer, as_attachment=True, download_name=filename, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'Failed to download PDF: {str(e)}'}), 500


@app.route('/download-excel', methods=['GET'])
def download_excel():
    try:
        global latest_data
        if not ensure_latest_data_available():
            return jsonify({'error': 'No processed data available. Please upload a file first.'}), 404

        excel_path = latest_data.get('excel_output_path')
        excel_filename = latest_data.get('excel_filename') or 'FTP_Results.xlsx'

        if not excel_path or not os.path.exists(excel_path):
            return jsonify({'error': 'Processed Excel output is not available. Please upload the file again.'}), 404

        return send_file(
            excel_path,
            as_attachment=True,
            download_name=excel_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': f'Failed to download Excel: {str(e)}'}), 500

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
        upload_start_time = perf_counter()

        def log_stage(stage_name, stage_start_time):
            elapsed_seconds = perf_counter() - stage_start_time
            print(f"[PERF] {stage_name}: {elapsed_seconds:.3f}s")

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
        read_start_time = perf_counter()
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        log_stage('Read workbook metadata', read_start_time)
        print(f"Found sheets: {sheet_names}")

        output_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f"FTP_Results_{month_name}_{year}_{output_stamp}.xlsx"
        excel_output_path = os.path.join(PROCESSED_OUTPUTS_DIR, excel_filename)
        
        sheets_data = {}
        global_summaries = {'ZWG': {}, 'FX': {}}
        
        branch_sbu_map = load_branch_sbu_map()
        branch_sbu_lookup = {code: value.get('sbu', 'Unknown') for code, value in branch_sbu_map.items()}
        
        excel_write_start_time = perf_counter()
        with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
            for sheet in sheet_names:
                print(f"Processing: {sheet}")
                sheet_start_time = perf_counter()
                try:
                    df = excel_file.parse(sheet_name=sheet)
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
                        df_processed['SBU'] = df_processed[branch_col].map(branch_sbu_lookup).fillna('Unknown')
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

                    # Calculate DimDays with vectorized date logic.
                    first_day_ts = pd.Timestamp(first_day.date())
                    last_day_ts = pd.Timestamp(last_day.date())
                    booking_dates = df_processed['BOOKING_DATE']
                    maturity_dates = df_processed['MATURITY_DATE']

                    full_period_days = (last_day_ts - first_day_ts).days + 1
                    dim_days = np.where(
                        (booking_dates <= first_day_ts) & (maturity_dates >= last_day_ts),
                        full_period_days,
                        np.where(
                            (booking_dates >= first_day_ts) & (maturity_dates >= last_day_ts),
                            (last_day_ts - booking_dates).dt.days + 1,
                            np.where(
                                (booking_dates >= first_day_ts) & (maturity_dates <= last_day_ts),
                                (maturity_dates - booking_dates).dt.days,
                                (maturity_dates - first_day_ts).dt.days
                            )
                        )
                    )
                    df_processed['DimDays'] = dim_days

                    # Calculate DTM and MTM
                    df_processed['DTM'] = np.where(
                        maturity_dates > last_day_ts,
                        (maturity_dates - last_day_ts).dt.days,
                        0
                    )
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

                    bucket_values = np.zeros((len(df_processed), len(bucket_labels)), dtype=float)
                    exposure_values = exposure.to_numpy(dtype=float, copy=False)
                    index_values = bucket_indices.to_numpy(dtype=int, copy=False)
                    positive_exposure_rows = np.nonzero(exposure_values > 0)[0]
                    bucket_values[positive_exposure_rows, index_values[positive_exposure_rows]] = exposure_values[positive_exposure_rows]
                    df_processed[bucket_labels] = bucket_values

                    df_processed['Check'] = df_processed[bucket_labels].sum(axis=1)

                    # FTP Charge
                    rates = zwg_rates if sheet == "ZWG LOANS" else usd_rates
                    rate_vector = np.array(
                        [(rates[i] if i < len(rates) else rates[-1]) / 100 for i in range(len(bucket_labels))],
                        dtype=float
                    )
                    df_processed['FTP Charge'] = df_processed[bucket_labels].to_numpy(dtype=float, copy=False) @ rate_vector

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

                    # Persist full processed worksheet for Excel download
                    df_processed.to_excel(writer, sheet_name=sheet[:31], index=False)
                    del df_processed

                else:
                    preview = df.head(100).copy()
                    sheets_data[sheet] = {
                        'columns': df.columns.tolist(),
                        'data': preview.to_dict(orient='records'),
                        'shape': df.shape
                    }
                    # Include non-loan sheets in Excel output as-is
                    df.to_excel(writer, sheet_name=sheet[:31], index=False)
                    del df

                    log_stage(f'Sheet {sheet}', sheet_start_time)
                print(f"Completed: {sheet}")
                log_stage('Write processed workbook', excel_write_start_time)
        
        # Store data
                snapshot_start_time = perf_counter()
        latest_data['filename'] = file.filename
        latest_data['sheets'] = sheets_data
        latest_data['summaries'] = global_summaries
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'year': year
        }
        latest_data['excel_output_path'] = excel_output_path
        latest_data['excel_filename'] = excel_filename

        save_latest_data_snapshot()
        log_stage('Persist snapshot', snapshot_start_time)
        log_stage('Total upload pipeline', upload_start_time)
        
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


@app.route('/ftp-curve-data', methods=['POST'])
def update_ftp_curve_data():
    global tenors, usd_rates, zwg_rates

    payload = request.get_json(silent=True) or {}
    updated_tenors = payload.get('tenors')
    updated_usd_rates = payload.get('usd_rates')
    updated_zwg_rates = payload.get('zwg_rates')

    if not isinstance(updated_tenors, list) or not isinstance(updated_usd_rates, list) or not isinstance(updated_zwg_rates, list):
        return jsonify({'error': 'tenors, usd_rates, and zwg_rates must all be arrays'}), 400

    if not updated_tenors or len(updated_tenors) != len(updated_usd_rates) or len(updated_tenors) != len(updated_zwg_rates):
        return jsonify({'error': 'Curve arrays must be non-empty and have matching lengths'}), 400

    try:
        normalized_tenors = [int(value) for value in updated_tenors]
        normalized_usd_rates = [float(value) for value in updated_usd_rates]
        normalized_zwg_rates = [float(value) for value in updated_zwg_rates]
    except (TypeError, ValueError):
        return jsonify({'error': 'All curve values must be numeric'}), 400

    if any(value <= 0 for value in normalized_tenors):
        return jsonify({'error': 'Tenors must be greater than zero'}), 400

    if normalized_tenors != sorted(normalized_tenors):
        return jsonify({'error': 'Tenors must be sorted in ascending order'}), 400

    curve_config = {
        'tenors': normalized_tenors,
        'usd_rates': normalized_usd_rates,
        'zwg_rates': normalized_zwg_rates
    }

    try:
        save_curve_config(curve_config)
    except OSError as exc:
        return jsonify({'error': f'Unable to save curve config: {str(exc)}'}), 500

    tenors = normalized_tenors
    usd_rates = normalized_usd_rates
    zwg_rates = normalized_zwg_rates

    return jsonify({
        'status': 'success',
        'tenors': tenors,
        'zwg': {'name': 'ZWG FTP Curve', 'rates': zwg_rates, 'color': '#b33a3a', 'borderColor': '#921f1f'},
        'usd': {'name': 'USD FTP Curve', 'rates': usd_rates, 'color': '#2563eb', 'borderColor': '#1d4ed8'}
    })


@app.route('/branch-sbu-map', methods=['GET'])
def get_branch_sbu_map():
    try:
        return jsonify({'status': 'success', 'mappings': load_branch_sbu_rows()})
    except Exception as e:
        return jsonify({'error': f'Failed to load branch mappings: {str(e)}'}), 500


@app.route('/branch-sbu-map', methods=['POST'])
def update_branch_sbu_map():
    payload = request.get_json(silent=True) or {}
    mappings = payload.get('mappings')

    if not isinstance(mappings, list):
        return jsonify({'error': 'mappings must be an array'}), 400

    try:
        save_branch_sbu_rows(mappings)
        return jsonify({'status': 'success', 'mappings': load_branch_sbu_rows()})
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'Failed to save branch mappings: {str(e)}'}), 500


@app.route('/branch-sbu-map/reset', methods=['POST'])
def reset_branch_sbu_map():
    try:
        reset_branch_sbu_map_to_default()
        return jsonify({'status': 'success', 'mappings': load_branch_sbu_rows()})
    except Exception as e:
        return jsonify({'error': f'Failed to reset branch mappings: {str(e)}'}), 500

@app.route('/get-preview', methods=['GET'])
def get_preview():
    if latest_data['sheets'] or load_latest_data_snapshot():
        return jsonify({'filename': latest_data['filename'], 'sheets': latest_data['sheets']})
    return jsonify({'message': 'No data uploaded yet'}), 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)