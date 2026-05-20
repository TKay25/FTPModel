from flask import Flask, request, jsonify, send_file
import json
import os
import sqlite3
import gc
import threading
import uuid
import shutil
import pandas as pd
import io
import openpyxl
import numpy as np
from datetime import datetime, timedelta
from time import perf_counter
import time
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import portrait, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER

app = Flask(__name__)

CURVE_CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'curve_config.json')
PROCESSED_REPORTS_DB_PATH = os.path.join(os.path.dirname(__file__), 'processed_reports.db')
PROCESSED_OUTPUTS_DIR = os.path.join(os.path.dirname(__file__), 'processed_outputs')
BRANCH_MAP_DB_PATH = os.path.join(os.path.dirname(__file__), 'branch_sbu_map.db')
LOAN_SHEETS = {'ZWG LOANS', 'FX LOANS'}
# Disable non-loan sheet export by default to keep memory use below small-instance limits.
INCLUDE_NON_LOAN_SHEETS = os.getenv('INCLUDE_NON_LOAN_SHEETS', '0').lower() in {'1', 'true', 'yes'}
# Guardrail for low-memory hosts: downgrade full-workbook export if upload is too large.
INCLUDE_WORKINGS_MAX_UPLOAD_MB = float(os.getenv('INCLUDE_WORKINGS_MAX_UPLOAD_MB', '6'))
REPORT_RETENTION_MAX_VERSIONS = int(os.getenv('REPORT_RETENTION_MAX_VERSIONS', '3'))
REPORT_RETENTION_MAX_MONTHS = int(os.getenv('REPORT_RETENTION_MAX_MONTHS', '24'))
FTP_REQUIRE_ROLE = os.getenv('FTP_REQUIRE_ROLE', '0').lower() in {'1', 'true', 'yes', 'on'}
FTP_DEFAULT_ROLE = os.getenv('FTP_DEFAULT_ROLE', 'admin').strip().lower() or 'admin'
ENABLE_SCHEDULER = os.getenv('ENABLE_SCHEDULER', '0').lower() in {'1', 'true', 'yes', 'on'}
SCHEDULER_POLL_SECONDS = int(os.getenv('SCHEDULER_POLL_SECONDS', '30'))
ENABLE_NOTIFICATIONS = os.getenv('ENABLE_NOTIFICATIONS', '1').lower() in {'1', 'true', 'yes', 'on'}
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

MONTH_NAME_TO_NUMBER = {
    'january': 1,
    'february': 2,
    'march': 3,
    'april': 4,
    'may': 5,
    'june': 6,
    'july': 7,
    'august': 8,
    'september': 9,
    'october': 10,
    'november': 11,
    'december': 12,
}
MONTH_NUMBER_TO_NAME = {value: name.title() for name, value in MONTH_NAME_TO_NUMBER.items()}

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
    'excel_filename': None,
    'excel_contains_workings': False,
    'skipped_sheet_count': 0,
    'export_warning': None,
}


def _json_default(value):
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, (datetime, timedelta)):
        return str(value)
    return str(value)


def _ensure_processed_reports_db():
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.execute(
            '''
            CREATE TABLE IF NOT EXISTS processed_reports (
                report_key TEXT PRIMARY KEY,
                month_number INTEGER NOT NULL,
                month_name TEXT NOT NULL,
                year INTEGER NOT NULL,
                filename TEXT,
                period_json TEXT NOT NULL,
                summaries_json TEXT NOT NULL,
                sheets_json TEXT NOT NULL,
                excel_output_path TEXT,
                excel_filename TEXT,
                excel_contains_workings INTEGER NOT NULL DEFAULT 0,
                skipped_sheet_count INTEGER NOT NULL DEFAULT 0,
                export_warning TEXT,
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            )
            '''
        )
        connection.execute(
            '''
            CREATE TABLE IF NOT EXISTS report_notes (
                note_id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_key TEXT NOT NULL,
                note_text TEXT NOT NULL,
                created_by TEXT,
                created_at REAL NOT NULL
            )
            '''
        )
        connection.execute(
            '''
            CREATE TABLE IF NOT EXISTS audit_log (
                event_id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_type TEXT NOT NULL,
                report_key TEXT,
                actor TEXT,
                actor_role TEXT,
                details_json TEXT,
                created_at REAL NOT NULL
            )
            '''
        )
        connection.execute(
            '''
            CREATE TABLE IF NOT EXISTS scheduler_jobs (
                job_id INTEGER PRIMARY KEY AUTOINCREMENT,
                folder_path TEXT NOT NULL,
                include_workings INTEGER NOT NULL DEFAULT 0,
                overwrite_existing INTEGER NOT NULL DEFAULT 0,
                active INTEGER NOT NULL DEFAULT 1,
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL,
                last_run_at REAL,
                last_status TEXT,
                last_message TEXT
            )
            '''
        )
        connection.commit()


def _current_actor():
    actor = (request.headers.get('X-FTP-Actor') if request else None) or request.args.get('actor') if request else None
    if not actor:
        actor = 'local-user'
    return str(actor).strip() or 'local-user'


def _current_role():
    role = (request.headers.get('X-FTP-Role') if request else None) or request.args.get('role') if request else None
    if not role:
        role = FTP_DEFAULT_ROLE
    return str(role).strip().lower() or FTP_DEFAULT_ROLE


def _require_role(*allowed_roles):
    if not FTP_REQUIRE_ROLE:
        return None

    role = _current_role()
    normalized_allowed = {str(value).strip().lower() for value in allowed_roles}
    if role not in normalized_allowed:
        return jsonify({'error': f'Action requires one of the following roles: {", ".join(sorted(normalized_allowed))}'}), 403
    return None


def _audit_event(event_type, report_key=None, details=None):
    _ensure_processed_reports_db()
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.execute(
            '''
            INSERT INTO audit_log (event_type, report_key, actor, actor_role, details_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ''',
            (
                event_type,
                report_key,
                _current_actor() if request else 'system',
                _current_role() if request else 'system',
                json.dumps(details or {}, ensure_ascii=True, default=_json_default),
                time.time(),
            )
        )
        connection.commit()


def _register_notification(message, level='info', report_key=None):
    if not ENABLE_NOTIFICATIONS:
        return
    _audit_event('notification', report_key=report_key, details={'message': message, 'level': level})


def _fetch_report_metadata(report_key):
    _ensure_processed_reports_db()
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.row_factory = sqlite3.Row
        cursor = connection.cursor()
        cursor.execute(
            '''
            SELECT report_key, month_number, month_name, year, filename, excel_filename,
                   excel_output_path, excel_contains_workings, skipped_sheet_count, export_warning,
                   created_at, updated_at, period_json, summaries_json, sheets_json
            FROM processed_reports
            WHERE report_key = ?
            ''',
            (report_key,)
        )
        return cursor.fetchone()


def _load_report_notes(report_key):
    _ensure_processed_reports_db()
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.row_factory = sqlite3.Row
        rows = connection.execute(
            'SELECT note_id, note_text, created_by, created_at FROM report_notes WHERE report_key = ? ORDER BY created_at DESC',
            (report_key,)
        ).fetchall()
    return [dict(row) for row in rows]


def _delete_report_workbook(excel_output_path):
    if excel_output_path and os.path.exists(excel_output_path):
        try:
            os.remove(excel_output_path)
        except OSError as exc:
            print(f"[WARN] Failed to remove archived workbook {excel_output_path}: {exc}")


def _cleanup_report_retention(month_number=None, year=None):
    _ensure_processed_reports_db()
    now = datetime.now()
    cutoff_year = now.year
    cutoff_month = now.month - REPORT_RETENTION_MAX_MONTHS
    while cutoff_month <= 0:
        cutoff_month += 12
        cutoff_year -= 1

    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.row_factory = sqlite3.Row
        rows_to_delete = []

        if REPORT_RETENTION_MAX_MONTHS > 0:
            old_rows = connection.execute(
                '''
                SELECT report_key, excel_output_path
                FROM processed_reports
                WHERE (year < ?) OR (year = ? AND month_number < ?)
                ''',
                (cutoff_year, cutoff_year, cutoff_month)
            ).fetchall()
            rows_to_delete.extend(old_rows)

        if month_number and year and REPORT_RETENTION_MAX_VERSIONS > 0:
            version_rows = connection.execute(
                '''
                SELECT report_key, excel_output_path
                FROM processed_reports
                WHERE month_number = ? AND year = ?
                ORDER BY updated_at DESC
                ''',
                (int(month_number), int(year))
            ).fetchall()
            rows_to_delete.extend(version_rows[REPORT_RETENTION_MAX_VERSIONS:])

        unique_rows = {}
        for row in rows_to_delete:
            unique_rows[row['report_key']] = row['excel_output_path']

        if not unique_rows:
            return 0

        for report_key, excel_output_path in unique_rows.items():
            connection.execute('DELETE FROM report_notes WHERE report_key = ?', (report_key,))
            connection.execute('DELETE FROM processed_reports WHERE report_key = ?', (report_key,))
            _audit_event('retention_delete', report_key=report_key, details={'reason': 'retention_policy'})
            _delete_report_workbook(excel_output_path)

        connection.commit()
        return len(unique_rows)


def save_latest_data_snapshot():
    period = latest_data.get('period') or {}
    month_number = period.get('month_number')
    if not month_number:
        month_value = str(period.get('month', '')).strip().lower()
        month_number = MONTH_NAME_TO_NUMBER.get(month_value)

    year_value = period.get('year')
    if not month_number or year_value is None:
        return

    month_number = int(month_number)
    year_value = int(year_value)
    report_key = period.get('report_key') or f'{year_value:04d}-{month_number:02d}'
    payload = {
        'filename': latest_data.get('filename'),
        'sheets': latest_data.get('sheets', {}),
        'ftp_results': latest_data.get('ftp_results'),
        'summaries': latest_data.get('summaries', {}),
        'period': {
            **period,
            'month_number': month_number,
            'report_key': report_key,
        },
        'excel_output_path': latest_data.get('excel_output_path'),
        'excel_filename': latest_data.get('excel_filename'),
        'excel_contains_workings': bool(latest_data.get('excel_contains_workings')),
        'skipped_sheet_count': int(latest_data.get('skipped_sheet_count') or 0),
        'export_warning': latest_data.get('export_warning'),
    }

    _ensure_processed_reports_db()
    now_ts = time.time()
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.execute(
            '''
            INSERT INTO processed_reports (
                report_key, month_number, month_name, year, filename,
                period_json, summaries_json, sheets_json,
                excel_output_path, excel_filename,
                excel_contains_workings, skipped_sheet_count, export_warning,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(report_key) DO UPDATE SET
                month_number = excluded.month_number,
                month_name = excluded.month_name,
                year = excluded.year,
                filename = excluded.filename,
                period_json = excluded.period_json,
                summaries_json = excluded.summaries_json,
                sheets_json = excluded.sheets_json,
                excel_output_path = excluded.excel_output_path,
                excel_filename = excluded.excel_filename,
                excel_contains_workings = excluded.excel_contains_workings,
                skipped_sheet_count = excluded.skipped_sheet_count,
                export_warning = excluded.export_warning,
                updated_at = excluded.updated_at
            ''',
            (
                report_key,
                month_number,
                MONTH_NUMBER_TO_NAME[month_number],
                year_value,
                payload['filename'],
                json.dumps(payload['period'], ensure_ascii=True, default=_json_default),
                json.dumps(payload['summaries'], ensure_ascii=True, default=_json_default),
                json.dumps(payload['sheets'], ensure_ascii=True, default=_json_default),
                payload['excel_output_path'],
                payload['excel_filename'],
                int(payload['excel_contains_workings']),
                int(payload['skipped_sheet_count']),
                payload['export_warning'],
                now_ts,
                now_ts,
            )
        )
        connection.commit()

    _audit_event(
        'report_saved',
        report_key=report_key,
        details={
            'month': payload['period'].get('month'),
            'year': payload['period'].get('year'),
            'excel_contains_workings': payload['excel_contains_workings'],
            'skipped_sheet_count': payload['skipped_sheet_count'],
            'filename': payload['filename'],
        }
    )
    _register_notification(f'Report {report_key} is ready for download.', level='success', report_key=report_key)
    _cleanup_report_retention(month_number=month_number, year=year_value)


def _normalize_month_number(month_value):
    if month_value is None:
        return None
    month_text = str(month_value).strip().lower()
    if not month_text:
        return None
    if month_text.isdigit():
        month_number = int(month_text)
    else:
        month_number = MONTH_NAME_TO_NUMBER.get(month_text)
    if month_number is None or month_number < 1 or month_number > 12:
        return None
    return month_number


def _make_report_key(year_value, month_number, version_suffix=None):
    base_key = f'{int(year_value):04d}-{int(month_number):02d}'
    if version_suffix is None:
        return base_key
    return f'{base_key}-{int(version_suffix)}'


def _load_processed_report_row(month=None, year=None, report_key=None):
    _ensure_processed_reports_db()
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.row_factory = sqlite3.Row
        cursor = connection.cursor()

        if report_key:
            cursor.execute('SELECT * FROM processed_reports WHERE report_key = ?', (report_key,))
            return cursor.fetchone()

        if month is None and year is None:
            cursor.execute('SELECT * FROM processed_reports ORDER BY updated_at DESC LIMIT 1')
            return cursor.fetchone()

        month_number = _normalize_month_number(month)
        if not month_number or year is None:
            return None

        cursor.execute(
            'SELECT * FROM processed_reports WHERE month_number = ? AND year = ? ORDER BY updated_at DESC LIMIT 1',
            (int(month_number), int(year))
        )
        return cursor.fetchone()


def _next_report_version_suffix(year_value, month_number):
    base_key = _make_report_key(year_value, month_number)
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        cursor = connection.cursor()
        cursor.execute(
            'SELECT report_key FROM processed_reports WHERE report_key = ? OR report_key LIKE ? ORDER BY report_key',
            (base_key, f'{base_key}-%')
        )
        existing_keys = [row[0] for row in cursor.fetchall()]

    if base_key not in existing_keys:
        return None

    suffixes = []
    for existing_key in existing_keys:
        if existing_key == base_key:
            suffixes.append(0)
            continue
        try:
            suffixes.append(int(existing_key.rsplit('-', 1)[1]))
        except (IndexError, ValueError):
            continue

    suffixes = sorted(set(suffixes))
    next_suffix = 1
    for suffix in suffixes:
        if suffix == next_suffix:
            next_suffix += 1
    return next_suffix


def load_latest_data_snapshot(month=None, year=None, report_key=None):
    global latest_data

    try:
        snapshot = _load_processed_report_row(month=month, year=year, report_key=report_key)
        if not snapshot:
            return False

        period = json.loads(snapshot['period_json']) if snapshot['period_json'] else {}
        latest_data = {
            'filename': snapshot['filename'],
            'sheets': json.loads(snapshot['sheets_json']) if snapshot['sheets_json'] else {},
            'ftp_results': None,
            'summaries': json.loads(snapshot['summaries_json']) if snapshot['summaries_json'] else {},
            'period': period,
            'excel_output_path': snapshot['excel_output_path'],
            'excel_filename': snapshot['excel_filename'],
            'excel_contains_workings': bool(snapshot['excel_contains_workings']),
            'skipped_sheet_count': int(snapshot['skipped_sheet_count'] or 0),
            'export_warning': snapshot['export_warning'],
        }
        return True
    except (OSError, ValueError, TypeError, json.JSONDecodeError, sqlite3.Error, KeyError):
        return False


def delete_processed_report(report_key):
    _ensure_processed_reports_db()
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.row_factory = sqlite3.Row
        cursor = connection.cursor()
        cursor.execute('SELECT excel_output_path FROM processed_reports WHERE report_key = ?', (report_key,))
        row = cursor.fetchone()
        if not row:
            return False

        excel_output_path = row['excel_output_path']
        cursor.execute('DELETE FROM report_notes WHERE report_key = ?', (report_key,))
        cursor.execute('DELETE FROM processed_reports WHERE report_key = ?', (report_key,))
        connection.commit()

    _delete_report_workbook(excel_output_path)
    _audit_event('report_deleted', report_key=report_key, details={'deleted_via': 'api'})
    _register_notification(f'Report {report_key} was deleted.', level='warning', report_key=report_key)

    return True


def ensure_latest_data_available(month=None, year=None):
    if latest_data.get('summaries') and latest_data.get('period'):
        return True
    return load_latest_data_snapshot(month=month, year=year)


def _resolve_month_year_args(month_value, year_value):
    if year_value is None:
        return None, None
    month_value = str(month_value).strip().lower() if month_value is not None else ''
    if not month_value:
        return None, None
    if month_value.isdigit():
        month_number = int(month_value)
    else:
        month_number = MONTH_NAME_TO_NUMBER.get(month_value)
    if not month_number:
        return None, None
    return MONTH_NUMBER_TO_NAME[month_number], int(year_value)

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
        report_key = request.args.get('report_key')
        month = request.args.get('month')
        year = request.args.get('year', type=int)
        if report_key:
            if not load_latest_data_snapshot(report_key=report_key):
                return jsonify({'error': 'No processed data found for that report version.'}), 404
        elif month and year:
            if not load_latest_data_snapshot(month=month, year=year):
                return jsonify({'error': 'No processed data found for that month and year.'}), 404
        elif not ensure_latest_data_available():
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
        report_key = request.args.get('report_key')
        month = request.args.get('month')
        year = request.args.get('year', type=int)
        if report_key:
            if not load_latest_data_snapshot(report_key=report_key):
                return jsonify({'error': 'No processed data found for that report version.'}), 404
        elif month and year:
            if not load_latest_data_snapshot(month=month, year=year):
                return jsonify({'error': 'No processed data found for that month and year.'}), 404
        elif not ensure_latest_data_available():
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

# ---------------------------------------------------------------------------
# Async job registry
# ---------------------------------------------------------------------------
_jobs = {}  # job_id -> { status, progress, stage, result, error, created_at, updated_at, completed_at }
_jobs_lock = threading.Lock()
JOB_RETENTION_SECONDS = int(os.environ.get('JOB_RETENTION_SECONDS', '1800'))
UPLOAD_JOBS_DB_PATH = os.path.join(os.path.dirname(__file__), 'upload_jobs.db')

TEMP_UPLOADS_DIR = os.path.join(os.path.dirname(__file__), 'temp_uploads')
os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)


def _update_job(job_id, **kwargs):
    now_ts = time.time()

    with _jobs_lock:
        if job_id in _jobs:
            _jobs[job_id].update(kwargs)
            _jobs[job_id]['updated_at'] = now_ts
            if kwargs.get('status') in {'done', 'error'}:
                _jobs[job_id]['completed_at'] = now_ts

    try:
        with sqlite3.connect(UPLOAD_JOBS_DB_PATH) as conn:
            cur = conn.cursor()
            fields = []
            values = []

            if 'status' in kwargs:
                fields.append('status = ?')
                values.append(kwargs['status'])
            if 'progress' in kwargs:
                fields.append('progress = ?')
                values.append(int(kwargs['progress']))
            if 'stage' in kwargs:
                fields.append('stage = ?')
                values.append(kwargs['stage'])
            if 'result' in kwargs:
                fields.append('result_json = ?')
                values.append(json.dumps(kwargs['result']))
            if 'error' in kwargs:
                fields.append('error = ?')
                values.append(kwargs['error'])

            fields.append('updated_at = ?')
            values.append(now_ts)

            if kwargs.get('status') in {'done', 'error'}:
                fields.append('completed_at = ?')
                values.append(now_ts)

            values.append(job_id)
            cur.execute(
                f"UPDATE upload_jobs SET {', '.join(fields)} WHERE job_id = ?",
                values
            )
            conn.commit()
    except Exception as exc:
        print(f"[WARN] Failed to persist job update for {job_id}: {exc}")


def _cleanup_expired_jobs():
    now = time.time()
    with _jobs_lock:
        expired_ids = []
        for jid, job in _jobs.items():
            completed_at = job.get('completed_at')
            if completed_at is None:
                continue
            if now - completed_at > JOB_RETENTION_SECONDS:
                expired_ids.append(jid)
        for jid in expired_ids:
            _jobs.pop(jid, None)

    try:
        threshold = now - JOB_RETENTION_SECONDS
        with sqlite3.connect(UPLOAD_JOBS_DB_PATH) as conn:
            cur = conn.cursor()
            cur.execute(
                "DELETE FROM upload_jobs WHERE completed_at IS NOT NULL AND completed_at < ?",
                (threshold,)
            )
            conn.commit()
    except Exception as exc:
        print(f"[WARN] Failed to cleanup expired upload jobs: {exc}")


def _ensure_upload_jobs_db():
    with sqlite3.connect(UPLOAD_JOBS_DB_PATH) as conn:
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS upload_jobs (
                job_id TEXT PRIMARY KEY,
                status TEXT NOT NULL,
                progress INTEGER NOT NULL DEFAULT 0,
                stage TEXT NOT NULL DEFAULT 'Queued',
                result_json TEXT,
                error TEXT,
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL,
                completed_at REAL
            )
        ''')
        conn.commit()


def _insert_upload_job(job_id):
    now_ts = time.time()
    with sqlite3.connect(UPLOAD_JOBS_DB_PATH) as conn:
        cur = conn.cursor()
        cur.execute('''
            INSERT OR REPLACE INTO upload_jobs
            (job_id, status, progress, stage, result_json, error, created_at, updated_at, completed_at)
            VALUES (?, 'running', 0, 'Queued', NULL, NULL, ?, ?, NULL)
        ''', (job_id, now_ts, now_ts))
        conn.commit()


def _fetch_upload_job(job_id):
    with sqlite3.connect(UPLOAD_JOBS_DB_PATH) as conn:
        cur = conn.cursor()
        cur.execute('''
            SELECT status, progress, stage, result_json, error
            FROM upload_jobs
            WHERE job_id = ?
        ''', (job_id,))
        row = cur.fetchone()
        if not row:
            return None
        result_json = row[3]
        result_payload = None
        if result_json:
            try:
                result_payload = json.loads(result_json)
            except Exception:
                result_payload = None
        return {
            'status': row[0],
            'progress': row[1],
            'stage': row[2],
            'result': result_payload,
            'error': row[4]
        }


_ensure_upload_jobs_db()


def _validate_upload_file(file_path, filename):
    validation = {
        'filename_valid': False,
        'required_sheets_present': False,
        'missing_sheets': [],
        'sheet_count': 0,
        'file_size_bytes': os.path.getsize(file_path) if file_path and os.path.exists(file_path) else 0,
        'warnings': [],
    }

    filename_match = re.search(r'FTP Input File (\w+) (\d{4})', filename or '')
    validation['filename_valid'] = bool(filename_match)
    if not validation['filename_valid']:
        validation['warnings'].append('Filename must be: FTP Input File Month Year.xlsx')
        return validation

    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
    except Exception as exc:
        validation['warnings'].append(f'Could not read workbook: {str(exc)}')
        return validation

    validation['sheet_count'] = len(sheet_names)
    missing_sheets = sorted(LOAN_SHEETS.difference(sheet_names))
    validation['missing_sheets'] = missing_sheets
    validation['required_sheets_present'] = not missing_sheets
    if missing_sheets:
        validation['warnings'].append(f'Missing required sheets: {", ".join(missing_sheets)}')

    if validation['file_size_bytes'] > int(INCLUDE_WORKINGS_MAX_UPLOAD_MB * 1024 * 1024):
        validation['warnings'].append('Workbook size exceeds the safe full-export threshold; results-only mode may be used.')

    return validation


def _run_scheduler_cycle():
    if not ENABLE_SCHEDULER:
        return 0

    _ensure_processed_reports_db()
    processed_files = 0
    with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
        connection.row_factory = sqlite3.Row
        jobs = connection.execute(
            'SELECT job_id, folder_path, include_workings, overwrite_existing FROM scheduler_jobs WHERE active = 1'
        ).fetchall()

        for job in jobs:
            folder_path = job['folder_path']
            if not os.path.isdir(folder_path):
                connection.execute(
                    'UPDATE scheduler_jobs SET updated_at = ?, last_status = ?, last_message = ? WHERE job_id = ?',
                    (time.time(), 'error', 'Folder not found', job['job_id'])
                )
                continue

            for entry in os.listdir(folder_path):
                file_path = os.path.join(folder_path, entry)
                if not os.path.isfile(file_path):
                    continue
                if not re.search(r'FTP Input File \w+ \d{4}.*\.(xlsx|xls)$', entry, flags=re.IGNORECASE):
                    continue

                archive_path = f'{file_path}.processed'
                if os.path.exists(archive_path):
                    continue

                validation = _validate_upload_file(file_path, entry)
                if not validation['filename_valid'] or not validation['required_sheets_present']:
                    connection.execute(
                        'UPDATE scheduler_jobs SET updated_at = ?, last_status = ?, last_message = ? WHERE job_id = ?',
                        (time.time(), 'error', '; '.join(validation['warnings']) or 'Validation failed', job['job_id'])
                    )
                    continue

                scheduler_job_id = f'scheduler-{uuid.uuid4()}'
                _insert_upload_job(scheduler_job_id)
                _run_ftp_job(
                    scheduler_job_id,
                    file_path,
                    entry,
                    bool(job['include_workings']),
                    bool(job['overwrite_existing']),
                    export_warning='Processed automatically by scheduler.'
                )
                shutil.copyfile(file_path, archive_path)
                connection.execute(
                    'UPDATE scheduler_jobs SET updated_at = ?, last_run_at = ?, last_status = ?, last_message = ? WHERE job_id = ?',
                    (time.time(), time.time(), 'success', f'Processed {entry}', job['job_id'])
                )
                processed_files += 1
                _audit_event('scheduler_processed', details={'folder_path': folder_path, 'filename': entry})

        connection.commit()

    return processed_files


def _run_ftp_job(job_id, upload_path, filename, include_non_loan_sheets, overwrite_existing=False, export_warning=None):
    """Background thread: process workbook and write results."""
    global latest_data

    def log_stage(stage_name, t0):
        print(f"[PERF] {stage_name}: {perf_counter() - t0:.3f}s")

    def progress(pct, stage):
        _update_job(job_id, progress=pct, stage=stage)

    try:
        upload_start_time = perf_counter()
        progress(5, 'Parsing filename')

        filename_match = re.search(r'FTP Input File (\w+) (\d{4})', filename)
        if not filename_match:
            _update_job(job_id, status='error', error='Filename must be: FTP Input File Month Year.xlsx')
            return

        month_name = filename_match.group(1)
        year = int(filename_match.group(2))

        month_map = {
            'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
            'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
        }
        month_num = month_map.get(month_name.lower())
        if not month_num:
            _update_job(job_id, status='error', error=f'Invalid month: {month_name}')
            return

        report_key_suffix = None if overwrite_existing else _next_report_version_suffix(year, month_num)
        report_key = _make_report_key(year, month_num, report_key_suffix)

        first_day = datetime(year, month_num, 1)
        last_day = (datetime(year + 1, 1, 1) if month_num == 12 else datetime(year, month_num + 1, 1)) - timedelta(days=1)

        progress(10, 'Reading workbook')
        read_start = perf_counter()
        excel_file = pd.ExcelFile(upload_path)
        sheet_names = excel_file.sheet_names
        log_stage('Read workbook metadata', read_start)
        print(f"Found sheets: {sheet_names}")

        output_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f"FTP_Results_{month_name}_{year}_{output_stamp}.xlsx"
        excel_output_path = os.path.join(PROCESSED_OUTPUTS_DIR, excel_filename)

        sheets_data = {}
        global_summaries = {'ZWG': {}, 'FX': {}}
        skipped_sheet_count = 0

        branch_sbu_map = load_branch_sbu_map()
        branch_sbu_lookup = {code: v.get('sbu', 'Unknown') for code, v in branch_sbu_map.items()}

        loan_sheet_names = [s for s in sheet_names if s in LOAN_SHEETS]
        total_loan_sheets = max(len(loan_sheet_names), 1)

        excel_write_start = perf_counter()
        if include_non_loan_sheets:
            # openpyxl is significantly more memory-hungry on large workbooks.
            # Use xlsxwriter without constant_memory for full-workbook exports.
            writer_engine = 'xlsxwriter'
            writer_kwargs = {}
        else:
            writer_engine = 'xlsxwriter'
            writer_kwargs = {'engine_kwargs': {'options': {'constant_memory': True}}}

        def process_sheets(writer):
            nonlocal skipped_sheet_count
            completed_loan = 0

            for sheet in sheet_names:
                print(f"Processing: {sheet}")
                sheet_start = perf_counter()

                if sheet not in LOAN_SHEETS and not include_non_loan_sheets:
                    skipped_sheet_count += 1
                    sheets_data[sheet] = {'columns': [], 'data': [], 'shape': [0, 0],
                                          'note': 'Skipped (memory optimisation)'}
                    log_stage(f'Sheet {sheet} (skipped)', sheet_start)
                    continue

                try:
                    df = excel_file.parse(sheet_name=sheet)
                except Exception as exc:
                    print(f"Error parsing {sheet}: {exc}")
                    continue

                if sheet in LOAN_SHEETS:
                    df_processed = df
                    del df

                    branch_col = next(
                        (c for c in ['Branch Code', 'BRANCHCODE', 'BRANCH_CODE'] if c in df_processed.columns),
                        None
                    )
                    if branch_col:
                        df_processed[branch_col] = df_processed[branch_col].astype(str).str.strip()
                        df_processed['SBU'] = df_processed[branch_col].map(branch_sbu_lookup).fillna('Unknown')
                    else:
                        df_processed['SBU'] = 'Unknown'

                    if 'BOOKING_DATE' in df_processed.columns:
                        df_processed['BOOKING_DATE'] = pd.to_datetime(df_processed['BOOKING_DATE'], errors='coerce').fillna(first_day)
                    if 'MATURITY_DATE' in df_processed.columns:
                        df_processed['MATURITY_DATE'] = pd.to_datetime(df_processed['MATURITY_DATE'], errors='coerce').fillna(first_day + timedelta(days=365))

                    if 'BOOKING_DATE' in df_processed.columns and 'MATURITY_DATE' in df_processed.columns:
                        df_processed['TENOR'] = (df_processed['MATURITY_DATE'] - df_processed['BOOKING_DATE']).dt.days.clip(lower=0)

                    first_day_ts = pd.Timestamp(first_day.date())
                    last_day_ts = pd.Timestamp(last_day.date())
                    bd = df_processed['BOOKING_DATE']
                    md = df_processed['MATURITY_DATE']

                    full_period = (last_day_ts - first_day_ts).days + 1
                    df_processed['DimDays'] = np.where(
                        (bd <= first_day_ts) & (md >= last_day_ts), full_period,
                        np.where((bd >= first_day_ts) & (md >= last_day_ts), (last_day_ts - bd).dt.days + 1,
                        np.where((bd >= first_day_ts) & (md <= last_day_ts), (md - bd).dt.days,
                                 (md - first_day_ts).dt.days))
                    ).astype(np.int32)

                    df_processed['DTM'] = np.where(md > last_day_ts, (md - last_day_ts).dt.days, 0).astype(np.int32)
                    df_processed['MTM'] = (df_processed['DTM'] / 30).round(1).astype(np.float32)

                    bucket_labels = ['<7days','7-14days','14-21days','21-30days','30-60days','60-90days',
                                     '90-180days','180-270days','270-360days','360-720days',
                                     '720-1080days','1080-1460days','1460-1800days','+1800days']
                    bin_edges = [0, 7, 14, 21, 30, 60, 90, 180, 270, 360, 720, 1080, 1460, 1800, float('inf')]

                    exposure = pd.to_numeric(
                        df_processed['Currency Exposure + Currency Accrued Reporting'], errors='coerce'
                    ).fillna(0).astype(np.float32)
                    mtm_days = (df_processed['MTM'] * 30).astype(np.float32)
                    bucket_idx = pd.cut(mtm_days, bins=bin_edges, labels=False, right=False, include_lowest=True)
                    bucket_idx = bucket_idx.fillna(len(bucket_labels) - 1).astype(int)

                    bv = np.zeros((len(df_processed), len(bucket_labels)), dtype=np.float32)
                    ev = exposure.to_numpy(dtype=np.float32, copy=False)
                    iv = bucket_idx.to_numpy(dtype=np.int16, copy=False)
                    pos = np.nonzero(ev > 0)[0]
                    bv[pos, iv[pos]] = ev[pos]
                    df_processed[bucket_labels] = bv
                    df_processed['Check'] = bv.sum(axis=1).astype(np.float32)

                    rates = zwg_rates if sheet == 'ZWG LOANS' else usd_rates
                    rv = np.array([(rates[i] if i < len(rates) else rates[-1]) / 100
                                   for i in range(len(bucket_labels))], dtype=np.float32)
                    df_processed['FTP Charge'] = (bv @ rv).astype(np.float32)

                    currency = 'ZWG' if sheet == 'ZWG LOANS' else 'FX'
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

                    preview = df_processed.head(100).copy()
                    for col in preview.select_dtypes(include=['datetime64']).columns:
                        preview[col] = preview[col].astype(str).replace('NaT', None)
                    sheets_data[sheet] = {
                        'columns': df_processed.columns.tolist(),
                        'data': preview.to_dict(orient='records'),
                        'shape': df_processed.shape
                    }

                    df_processed.to_excel(writer, sheet_name=sheet[:31], index=False)
                    del df_processed

                    completed_loan += 1
                    pct = 15 + int(75 * completed_loan / total_loan_sheets)
                    progress(pct, f'Processed {sheet}')

                else:
                    preview = df.head(100).copy()
                    sheets_data[sheet] = {
                        'columns': df.columns.tolist(),
                        'data': preview.to_dict(orient='records'),
                        'shape': df.shape
                    }
                    df.to_excel(writer, sheet_name=sheet[:31], index=False)
                    del df

                gc.collect()
                log_stage(f'Sheet {sheet}', sheet_start)
                print(f"Completed: {sheet}")

        try:
            with pd.ExcelWriter(excel_output_path, engine=writer_engine, **writer_kwargs) as writer:
                process_sheets(writer)
        except (ImportError, ModuleNotFoundError, ValueError):
            with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
                process_sheets(writer)

        log_stage('Write processed workbook', excel_write_start)
        progress(92, 'Saving results')

        snapshot_start_time = perf_counter()
        latest_data['filename'] = filename
        latest_data['sheets'] = sheets_data
        latest_data['summaries'] = global_summaries
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'month_number': month_num,
            'year': year,
            'report_key': report_key,
            'version_suffix': report_key_suffix
        }
        latest_data['excel_output_path'] = excel_output_path
        latest_data['excel_filename'] = excel_filename
        latest_data['excel_contains_workings'] = include_non_loan_sheets
        latest_data['skipped_sheet_count'] = skipped_sheet_count
        latest_data['export_warning'] = export_warning

        save_latest_data_snapshot()
        log_stage('Persist snapshot', snapshot_start_time)
        log_stage('Total upload pipeline', upload_start_time)
        print(f"✅ Stored summaries: {list(global_summaries.keys())}")

        _update_job(job_id,
                    status='done',
                    progress=100,
                    stage='Complete',
                    result={
                        'status': 'success',
                        'summary': global_summaries,
                        'period': latest_data['period'],
                        'excel_contains_workings': include_non_loan_sheets,
                        'skipped_sheet_count': skipped_sheet_count,
                        'export_warning': export_warning,
                        'report_key': report_key
                    })

    except Exception as exc:
        import traceback
        traceback.print_exc()
        _update_job(job_id, status='error', error=str(exc))
    finally:
        try:
            if upload_path and os.path.exists(upload_path):
                os.remove(upload_path)
        except OSError as exc:
            print(f"[WARN] Failed to remove temp upload {upload_path}: {exc}")


@app.route('/')
def index():
    from flask import send_from_directory
    return send_from_directory('.', 'index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    _cleanup_expired_jobs()

    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    filename = file.filename

    filename_match = re.search(r'FTP Input File (\w+) (\d{4})', filename)
    if not filename_match:
        return jsonify({'error': 'Filename must be: FTP Input File Month Year.xlsx'}), 400

    temp_upload_path = os.path.join(TEMP_UPLOADS_DIR, f'{uuid.uuid4()}_{os.path.basename(filename)}')
    try:
        file.save(temp_upload_path)
    except Exception as exc:
        return jsonify({'error': f'Failed to store upload temporarily: {str(exc)}'}), 500

    include_workings_raw = str(request.form.get('include_workings', '1')).strip().lower()
    include_workings_requested = include_workings_raw in {'1', 'true', 'yes', 'on'}
    include_non_loan_sheets = INCLUDE_NON_LOAN_SHEETS or include_workings_requested
    overwrite_existing_raw = str(request.form.get('overwrite_existing', '0')).strip().lower()
    overwrite_existing = overwrite_existing_raw in {'1', 'true', 'yes', 'on'}

    upload_size_mb = os.path.getsize(temp_upload_path) / (1024 * 1024)
    export_warning = None
    if include_non_loan_sheets and not INCLUDE_NON_LOAN_SHEETS and upload_size_mb > INCLUDE_WORKINGS_MAX_UPLOAD_MB:
        include_non_loan_sheets = False
        export_warning = (
            f'Include workings sheets was requested but automatically downgraded to results-only '
            f'because the upload size ({upload_size_mb:.1f} MB) exceeds the safe limit '
            f'({INCLUDE_WORKINGS_MAX_UPLOAD_MB:.1f} MB) for this hosting memory profile.'
        )

    job_id = str(uuid.uuid4())
    now = time.time()
    with _jobs_lock:
        _jobs[job_id] = {
            'status': 'running',
            'progress': 0,
            'stage': 'Queued',
            'result': None,
            'error': None,
            'created_at': now,
            'updated_at': now,
            'completed_at': None
        }
    _insert_upload_job(job_id)

    thread = threading.Thread(
        target=_run_ftp_job,
        args=(job_id, temp_upload_path, filename, include_non_loan_sheets, overwrite_existing, export_warning),
        daemon=True
    )
    thread.start()

    return jsonify({'job_id': job_id, 'status': 'running'})


@app.route('/upload/preflight', methods=['POST'])
def upload_preflight():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    temp_upload_path = os.path.join(TEMP_UPLOADS_DIR, f'{uuid.uuid4()}_{os.path.basename(file.filename)}')
    try:
        file.save(temp_upload_path)
        validation = _validate_upload_file(temp_upload_path, file.filename)
        return jsonify(validation)
    except Exception as exc:
        return jsonify({'error': f'Failed to validate workbook: {str(exc)}'}), 500
    finally:
        if os.path.exists(temp_upload_path):
            try:
                os.remove(temp_upload_path)
            except OSError:
                pass


@app.route('/upload/status/<job_id>', methods=['GET'])
def upload_status(job_id):
    _cleanup_expired_jobs()

    # Use SQLite as the authoritative store so status checks work across workers.
    job = _fetch_upload_job(job_id)
    if not job:
        return jsonify({'error': 'Unknown job'}), 404
    response = {
        'status': job['status'],
        'progress': job['progress'],
        'stage': job['stage'],
    }
    if job['status'] == 'done':
        response['result'] = job['result']
    elif job['status'] == 'error':
        response['error'] = job['error']
    return jsonify(response)


# ---------------------------------------------------------------------------
# Legacy synchronous path kept for local dev / non-Render environments.
# Remove once async path is proven in production.
# ---------------------------------------------------------------------------
def _upload_file_sync():
    global latest_data

    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    temp_upload_path = os.path.join(TEMP_UPLOADS_DIR, f'{uuid.uuid4()}_{os.path.basename(file.filename)}')

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

        report_key_suffix = None if overwrite_existing else _next_report_version_suffix(year, month_num)
        report_key = _make_report_key(year, month_num, report_key_suffix)

        try:
            file.save(temp_upload_path)
        except Exception as exc:
            return jsonify({'error': f'Failed to store upload temporarily: {str(exc)}'}), 500

        include_workings_raw = str(request.form.get('include_workings', '1')).strip().lower()
        include_workings_requested = include_workings_raw in {'1', 'true', 'yes', 'on'}
        include_non_loan_sheets = INCLUDE_NON_LOAN_SHEETS or include_workings_requested
        overwrite_existing_raw = str(request.form.get('overwrite_existing', '0')).strip().lower()
        overwrite_existing = overwrite_existing_raw in {'1', 'true', 'yes', 'on'}
        
        first_day = datetime(year, month_num, 1)
        if month_num == 12:
            last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(year, month_num + 1, 1) - timedelta(days=1)
        
        # Get sheet names
        read_start_time = perf_counter()
        excel_file = pd.ExcelFile(temp_upload_path)
        sheet_names = excel_file.sheet_names
        log_stage('Read workbook metadata', read_start_time)
        print(f"Found sheets: {sheet_names}")

        output_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f"FTP_Results_{month_name}_{year}_{output_stamp}.xlsx"
        excel_output_path = os.path.join(PROCESSED_OUTPUTS_DIR, excel_filename)
        
        sheets_data = {}
        global_summaries = {'ZWG': {}, 'FX': {}}
        skipped_sheet_count = 0
        
        branch_sbu_map = load_branch_sbu_map()
        branch_sbu_lookup = {code: value.get('sbu', 'Unknown') for code, value in branch_sbu_map.items()}
        
        excel_write_start_time = perf_counter()
        if include_non_loan_sheets:
            # openpyxl can exceed memory limits for large files; prefer xlsxwriter here.
            writer_engine = 'xlsxwriter'
            writer_kwargs = {}
        else:
            writer_engine = 'xlsxwriter'
            writer_kwargs = {'engine_kwargs': {'options': {'constant_memory': True}}}
        try:
            with pd.ExcelWriter(excel_output_path, engine=writer_engine, **writer_kwargs) as writer:
                for sheet in sheet_names:
                    print(f"Processing: {sheet}")
                    sheet_start_time = perf_counter()

                    if sheet not in LOAN_SHEETS and not include_non_loan_sheets:
                        skipped_sheet_count += 1
                        sheets_data[sheet] = {
                            'columns': [],
                            'data': [],
                            'shape': [0, 0],
                            'note': 'Skipped for memory optimization (set INCLUDE_NON_LOAN_SHEETS=1 to include)'
                        }
                        log_stage(f'Sheet {sheet} (skipped)', sheet_start_time)
                        print(f"Completed: {sheet} (skipped)")
                        continue

                    try:
                        df = excel_file.parse(sheet_name=sheet)
                    except Exception as e:
                        print(f"Error: {e}")
                        continue

                    if sheet in LOAN_SHEETS:
                        df_processed = df
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
                        df_processed['DimDays'] = dim_days.astype(np.int32)

                        # Calculate DTM and MTM
                        df_processed['DTM'] = np.where(
                            maturity_dates > last_day_ts,
                            (maturity_dates - last_day_ts).dt.days,
                            0
                        ).astype(np.int32)
                        df_processed['MTM'] = (df_processed['DTM'] / 30).round(1).astype(np.float32)

                        # Bucket columns
                        bucket_labels = ['<7days', '7-14days', '14-21days', '21-30days', '30-60days', '60-90days', '90-180days', '180-270days', '270-360days', '360-720days', '720-1080days', '1080-1460days', '1460-1800days', '+1800days']

                        # Allocate to buckets
                        exposure = pd.to_numeric(
                            df_processed['Currency Exposure + Currency Accrued Reporting'],
                            errors='coerce'
                        ).fillna(0).astype(np.float32)
                        mtm_days = (df_processed['MTM'] * 30).astype(np.float32)
                        tenors_list = [7,14,21,30,60,90,180,270,360,720,1080,1460,1800]
                        bin_edges = [0] + tenors_list + [float('inf')]
                        bucket_indices = pd.cut(mtm_days, bins=bin_edges, labels=False, right=False, include_lowest=True)
                        bucket_indices = bucket_indices.fillna(len(bucket_labels)-1).astype(int)

                        bucket_values = np.zeros((len(df_processed), len(bucket_labels)), dtype=np.float32)
                        exposure_values = exposure.to_numpy(dtype=np.float32, copy=False)
                        index_values = bucket_indices.to_numpy(dtype=np.int16, copy=False)
                        positive_exposure_rows = np.nonzero(exposure_values > 0)[0]
                        bucket_values[positive_exposure_rows, index_values[positive_exposure_rows]] = exposure_values[positive_exposure_rows]
                        df_processed[bucket_labels] = bucket_values

                        df_processed['Check'] = df_processed[bucket_labels].sum(axis=1).astype(np.float32)

                        # FTP Charge
                        rates = zwg_rates if sheet == "ZWG LOANS" else usd_rates
                        rate_vector = np.array(
                            [(rates[i] if i < len(rates) else rates[-1]) / 100 for i in range(len(bucket_labels))],
                            dtype=np.float32
                        )
                        df_processed['FTP Charge'] = (
                            df_processed[bucket_labels].to_numpy(dtype=np.float32, copy=False) @ rate_vector
                        ).astype(np.float32)

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
                        # Optional: include non-loan sheets in output when explicitly enabled.
                        df.to_excel(writer, sheet_name=sheet[:31], index=False)
                        del df

                    gc.collect()
                    log_stage(f'Sheet {sheet}', sheet_start_time)
                    print(f"Completed: {sheet}")
        except (ImportError, ModuleNotFoundError, ValueError):
            # Fallback when xlsxwriter is unavailable.
            with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
                for sheet in sheet_names:
                    print(f"Processing: {sheet}")
                    sheet_start_time = perf_counter()

                    if sheet not in LOAN_SHEETS and not include_non_loan_sheets:
                        skipped_sheet_count += 1
                        sheets_data[sheet] = {
                            'columns': [],
                            'data': [],
                            'shape': [0, 0],
                            'note': 'Skipped for memory optimization (set INCLUDE_NON_LOAN_SHEETS=1 to include)'
                        }
                        log_stage(f'Sheet {sheet} (skipped)', sheet_start_time)
                        print(f"Completed: {sheet} (skipped)")
                        continue

                    try:
                        df = excel_file.parse(sheet_name=sheet)
                    except Exception as e:
                        print(f"Error: {e}")
                        continue

                    if sheet in LOAN_SHEETS:
                        df_processed = df
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
                        df_processed['DimDays'] = dim_days.astype(np.int32)

                        # Calculate DTM and MTM
                        df_processed['DTM'] = np.where(
                            maturity_dates > last_day_ts,
                            (maturity_dates - last_day_ts).dt.days,
                            0
                        ).astype(np.int32)
                        df_processed['MTM'] = (df_processed['DTM'] / 30).round(1).astype(np.float32)

                        # Bucket columns
                        bucket_labels = ['<7days', '7-14days', '14-21days', '21-30days', '30-60days', '60-90days', '90-180days', '180-270days', '270-360days', '360-720days', '720-1080days', '1080-1460days', '1460-1800days', '+1800days']

                        # Allocate to buckets
                        exposure = pd.to_numeric(
                            df_processed['Currency Exposure + Currency Accrued Reporting'],
                            errors='coerce'
                        ).fillna(0).astype(np.float32)
                        mtm_days = (df_processed['MTM'] * 30).astype(np.float32)
                        tenors_list = [7,14,21,30,60,90,180,270,360,720,1080,1460,1800]
                        bin_edges = [0] + tenors_list + [float('inf')]
                        bucket_indices = pd.cut(mtm_days, bins=bin_edges, labels=False, right=False, include_lowest=True)
                        bucket_indices = bucket_indices.fillna(len(bucket_labels)-1).astype(int)

                        bucket_values = np.zeros((len(df_processed), len(bucket_labels)), dtype=np.float32)
                        exposure_values = exposure.to_numpy(dtype=np.float32, copy=False)
                        index_values = bucket_indices.to_numpy(dtype=np.int16, copy=False)
                        positive_exposure_rows = np.nonzero(exposure_values > 0)[0]
                        bucket_values[positive_exposure_rows, index_values[positive_exposure_rows]] = exposure_values[positive_exposure_rows]
                        df_processed[bucket_labels] = bucket_values

                        df_processed['Check'] = df_processed[bucket_labels].sum(axis=1).astype(np.float32)

                        # FTP Charge
                        rates = zwg_rates if sheet == "ZWG LOANS" else usd_rates
                        rate_vector = np.array(
                            [(rates[i] if i < len(rates) else rates[-1]) / 100 for i in range(len(bucket_labels))],
                            dtype=np.float32
                        )
                        df_processed['FTP Charge'] = (
                            df_processed[bucket_labels].to_numpy(dtype=np.float32, copy=False) @ rate_vector
                        ).astype(np.float32)

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
                        # Optional: include non-loan sheets in output when explicitly enabled.
                        df.to_excel(writer, sheet_name=sheet[:31], index=False)
                        del df

                    gc.collect()
                    log_stage(f'Sheet {sheet}', sheet_start_time)
                    print(f"Completed: {sheet}")
            log_stage('Write processed workbook', excel_write_start_time)

            # Store data
        latest_data['filename'] = file.filename
        latest_data['sheets'] = sheets_data
        latest_data['summaries'] = global_summaries
        latest_data['period'] = {
            'first_day': first_day.strftime('%d %B %Y'),
            'last_day': last_day.strftime('%d %B %Y'),
            'month': month_name,
            'month_number': month_num,
            'year': year,
            'report_key': report_key,
            'version_suffix': report_key_suffix
        }
        latest_data['excel_output_path'] = excel_output_path
        latest_data['excel_filename'] = excel_filename
        latest_data['excel_contains_workings'] = include_non_loan_sheets
        latest_data['skipped_sheet_count'] = skipped_sheet_count

        snapshot_start_time = perf_counter()
        save_latest_data_snapshot()
        log_stage('Persist snapshot', snapshot_start_time)
        log_stage('Total upload pipeline', upload_start_time)
        
        print(f"✅ Stored summaries: {list(global_summaries.keys())}")
        
        return jsonify({
            'status': 'success',
            'summary': global_summaries,
            'period': latest_data['period'],
            'excel_contains_workings': include_non_loan_sheets,
            'skipped_sheet_count': skipped_sheet_count,
            'report_key': report_key
        })
    
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        try:
            if os.path.exists(temp_upload_path):
                os.remove(temp_upload_path)
        except OSError as exc:
            print(f"[WARN] Failed to remove temp upload {temp_upload_path}: {exc}")

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
    report_key = request.args.get('report_key')
    month = request.args.get('month')
    year = request.args.get('year', type=int)
    if report_key:
        if load_latest_data_snapshot(report_key=report_key):
            return jsonify({'filename': latest_data['filename'], 'sheets': latest_data['sheets'], 'period': latest_data.get('period')})
        return jsonify({'message': 'No data found for that report version'}), 404
    if month and year:
        if load_latest_data_snapshot(month=month, year=year):
            return jsonify({'filename': latest_data['filename'], 'sheets': latest_data['sheets'], 'period': latest_data.get('period')})
        return jsonify({'message': 'No data found for that month and year'}), 404

    if latest_data['sheets'] or load_latest_data_snapshot():
        return jsonify({'filename': latest_data['filename'], 'sheets': latest_data['sheets']})
    return jsonify({'message': 'No data uploaded yet'}), 404


@app.route('/processed-reports', methods=['GET'])
def processed_reports():
    try:
        _ensure_processed_reports_db()
        month = request.args.get('month')
        year = request.args.get('year', type=int)
        with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
            connection.row_factory = sqlite3.Row
            if month and year:
                month_number = _normalize_month_number(month)
                rows = connection.execute(
                    '''
                    SELECT report_key, month_number, month_name, year, filename, excel_filename, updated_at, excel_contains_workings, skipped_sheet_count, export_warning
                    FROM processed_reports
                    WHERE month_number = ? AND year = ?
                    ORDER BY updated_at DESC
                    ''',
                    (month_number, int(year))
                ).fetchall()
            else:
                rows = connection.execute(
                    '''
                    SELECT report_key, month_number, month_name, year, filename, excel_filename, updated_at, excel_contains_workings, skipped_sheet_count, export_warning
                    FROM processed_reports
                    ORDER BY year DESC, month_number DESC, updated_at DESC
                    '''
                ).fetchall()

        return jsonify({
            'reports': [
                {
                    'report_key': row['report_key'],
                    'month': row['month_name'],
                    'month_number': row['month_number'],
                    'year': row['year'],
                    'filename': row['filename'],
                    'excel_filename': row['excel_filename'],
                    'updated_at': row['updated_at'],
                    'excel_contains_workings': bool(row['excel_contains_workings']),
                    'skipped_sheet_count': row['skipped_sheet_count'],
                    'export_warning': row['export_warning'],
                    'notes_count': len(_load_report_notes(row['report_key']))
                }
                for row in rows
            ]
        })
    except Exception as exc:
        return jsonify({'error': f'Failed to list processed reports: {str(exc)}'}), 500


@app.route('/processed-reports/<report_key>/metadata', methods=['GET'])
def processed_report_metadata(report_key):
    try:
        row = _fetch_report_metadata(report_key)
        if not row:
            return jsonify({'error': 'Report version not found'}), 404

        summaries = json.loads(row['summaries_json']) if row['summaries_json'] else {}
        sheets = json.loads(row['sheets_json']) if row['sheets_json'] else {}
        total_rows = sum((sheet_data.get('row_count') or 0) for currency in summaries.values() for sheet_data in currency.values())

        return jsonify({
            'report_key': row['report_key'],
            'month': row['month_name'],
            'month_number': row['month_number'],
            'year': row['year'],
            'filename': row['filename'],
            'excel_filename': row['excel_filename'],
            'excel_contains_workings': bool(row['excel_contains_workings']),
            'skipped_sheet_count': row['skipped_sheet_count'],
            'export_warning': row['export_warning'],
            'sheet_count': len(sheets),
            'total_rows': total_rows,
            'file_size_bytes': os.path.getsize(row['excel_output_path']) if row['excel_output_path'] and os.path.exists(row['excel_output_path']) else None,
            'notes': _load_report_notes(report_key),
        })
    except Exception as exc:
        return jsonify({'error': f'Failed to load report metadata: {str(exc)}'}), 500


@app.route('/processed-reports/<report_key>/notes', methods=['GET', 'POST'])
def processed_report_notes(report_key):
    try:
        if request.method == 'GET':
            return jsonify({'notes': _load_report_notes(report_key)})

        role_error = _require_role('admin', 'editor')
        if role_error:
            return role_error

        payload = request.get_json(silent=True) or {}
        note_text = str(payload.get('note', '')).strip()
        if not note_text:
            return jsonify({'error': 'note is required'}), 400

        _ensure_processed_reports_db()
        with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
            connection.execute(
                'INSERT INTO report_notes (report_key, note_text, created_by, created_at) VALUES (?, ?, ?, ?)',
                (report_key, note_text, _current_actor(), time.time())
            )
            connection.commit()

        _audit_event('note_added', report_key=report_key, details={'note': note_text})
        return jsonify({'status': 'success', 'notes': _load_report_notes(report_key)})
    except Exception as exc:
        return jsonify({'error': f'Failed to manage report notes: {str(exc)}'}), 500


@app.route('/processed-reports/compare', methods=['GET'])
def compare_processed_reports():
    report_key_a = request.args.get('report_key_a')
    report_key_b = request.args.get('report_key_b')
    if not report_key_a or not report_key_b:
        return jsonify({'error': 'report_key_a and report_key_b are required'}), 400

    row_a = _fetch_report_metadata(report_key_a)
    row_b = _fetch_report_metadata(report_key_b)
    if not row_a or not row_b:
        return jsonify({'error': 'One or both report versions were not found'}), 404

    summaries_a = json.loads(row_a['summaries_json']) if row_a['summaries_json'] else {}
    summaries_b = json.loads(row_b['summaries_json']) if row_b['summaries_json'] else {}

    comparison = []
    currencies = sorted(set(summaries_a.keys()) | set(summaries_b.keys()))
    for currency in currencies:
        sheets = sorted(set((summaries_a.get(currency) or {}).keys()) | set((summaries_b.get(currency) or {}).keys()))
        for sheet in sheets:
            left = (summaries_a.get(currency) or {}).get(sheet, {})
            right = (summaries_b.get(currency) or {}).get(sheet, {})
            comparison.append({
                'currency': currency,
                'sheet': sheet,
                'exposure_delta': float((right.get('total_exposure') or 0) - (left.get('total_exposure') or 0)),
                'ftp_delta': float((right.get('total_ftp_charge') or 0) - (left.get('total_ftp_charge') or 0)),
                'row_count_delta': int((right.get('row_count') or 0) - (left.get('row_count') or 0)),
            })

    _audit_event('report_compared', details={'report_key_a': report_key_a, 'report_key_b': report_key_b})
    return jsonify({'comparison': comparison})


@app.route('/audit-log', methods=['GET'])
def audit_log():
    try:
        limit = request.args.get('limit', default=100, type=int)
        with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
            connection.row_factory = sqlite3.Row
            rows = connection.execute(
                'SELECT event_id, event_type, report_key, actor, actor_role, details_json, created_at FROM audit_log ORDER BY created_at DESC LIMIT ?',
                (max(1, min(limit, 500)),)
            ).fetchall()

        return jsonify({
            'events': [
                {
                    'event_id': row['event_id'],
                    'event_type': row['event_type'],
                    'report_key': row['report_key'],
                    'actor': row['actor'],
                    'actor_role': row['actor_role'],
                    'details': json.loads(row['details_json']) if row['details_json'] else {},
                    'created_at': row['created_at'],
                }
                for row in rows
            ]
        })
    except Exception as exc:
        return jsonify({'error': f'Failed to load audit log: {str(exc)}'}), 500


@app.route('/notifications', methods=['GET'])
def notifications():
    try:
        with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
            connection.row_factory = sqlite3.Row
            rows = connection.execute(
                "SELECT event_id, report_key, details_json, created_at FROM audit_log WHERE event_type = 'notification' ORDER BY created_at DESC LIMIT 20"
            ).fetchall()

        return jsonify({
            'notifications': [
                {
                    'event_id': row['event_id'],
                    'report_key': row['report_key'],
                    'created_at': row['created_at'],
                    **(json.loads(row['details_json']) if row['details_json'] else {})
                }
                for row in rows
            ]
        })
    except Exception as exc:
        return jsonify({'error': f'Failed to load notifications: {str(exc)}'}), 500


@app.route('/settings/retention', methods=['GET'])
def retention_settings():
    return jsonify({
        'max_versions_per_month': REPORT_RETENTION_MAX_VERSIONS,
        'max_months': REPORT_RETENTION_MAX_MONTHS,
        'role_enforced': FTP_REQUIRE_ROLE,
        'default_role': FTP_DEFAULT_ROLE,
        'scheduler_enabled': ENABLE_SCHEDULER,
    })


@app.route('/scheduler-jobs', methods=['GET', 'POST'])
def scheduler_jobs():
    try:
        _ensure_processed_reports_db()
        if request.method == 'GET':
            with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
                connection.row_factory = sqlite3.Row
                rows = connection.execute(
                    'SELECT job_id, folder_path, include_workings, overwrite_existing, active, created_at, updated_at, last_run_at, last_status, last_message FROM scheduler_jobs ORDER BY updated_at DESC'
                ).fetchall()
            return jsonify({'jobs': [dict(row) for row in rows]})

        role_error = _require_role('admin')
        if role_error:
            return role_error

        payload = request.get_json(silent=True) or {}
        folder_path = str(payload.get('folder_path', '')).strip()
        if not folder_path:
            return jsonify({'error': 'folder_path is required'}), 400

        now_ts = time.time()
        with sqlite3.connect(PROCESSED_REPORTS_DB_PATH) as connection:
            connection.execute(
                '''
                INSERT INTO scheduler_jobs (folder_path, include_workings, overwrite_existing, active, created_at, updated_at)
                VALUES (?, ?, ?, 1, ?, ?)
                ''',
                (folder_path, int(bool(payload.get('include_workings'))), int(bool(payload.get('overwrite_existing'))), now_ts, now_ts)
            )
            connection.commit()

        _audit_event('scheduler_job_added', details={'folder_path': folder_path})
        return scheduler_jobs()
    except Exception as exc:
        return jsonify({'error': f'Failed to manage scheduler jobs: {str(exc)}'}), 500


@app.route('/scheduler-jobs/run', methods=['POST'])
def run_scheduler_jobs():
    try:
        role_error = _require_role('admin')
        if role_error:
            return role_error

        processed_files = _run_scheduler_cycle()
        return jsonify({'status': 'success', 'processed_files': processed_files})
    except Exception as exc:
        return jsonify({'error': f'Failed to run scheduler: {str(exc)}'}), 500


@app.route('/processed-reports/<report_key>', methods=['DELETE'])
def delete_processed_report_route(report_key):
    try:
        if not report_key:
            return jsonify({'error': 'report_key is required'}), 400

        deleted = delete_processed_report(report_key)
        if not deleted:
            return jsonify({'error': 'Report version not found'}), 404

        return jsonify({'status': 'success', 'report_key': report_key})
    except Exception as exc:
        return jsonify({'error': f'Failed to delete processed report: {str(exc)}'}), 500


_ensure_processed_reports_db()

if __name__ == '__main__':
    app.run(debug=True, port=5000)