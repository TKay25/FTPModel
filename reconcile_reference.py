"""Reconcile the model's processed FTP output against a CBZ reference workbook.

The reference model keeps loan-level results on the 'Assets' sheet of
"CBZ Bank _USD/ZWG New FTP Model - <Month> <Year>.xlsb". This script compares the
FTP Charge (and the DinM / DTM / MTM inputs) produced by the Flask model for the
same loan book, loan by loan.

Usage
-----
    python reconcile_reference.py --reference "CBZ Bank _USD New FTP Model - April 2026.xlsb" \
                                  --sheet "FX LOANS" \
                                  [--processed "processed_outputs/FTP_Results_April_2026_....xlsx"] \
                                  [--report reconciliation_april_usd.xlsx]

When --processed is omitted the newest processed output that contains the sheet
is used. Exit code is 0 only when every matched loan agrees to within 1e-6
(absolute) or 1e-9 (relative).

Rows are matched on branch code + booking date + maturity date + exposure, so
the model output and the reference can be in any order. Reference rows flagged as
Treasury (Treasury Marker populated) are treasury assets rather than loans and
are excluded from the comparison.
"""
import argparse
import glob
import os
import sys

import pandas as pd
from pyxlsb import open_workbook

EPOCH = pd.Timestamp('1899-12-30')
MATCH_TOLERANCE_ABS = 1e-6
MATCH_TOLERANCE_REL = 1e-9


def _serial(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return int((pd.Timestamp(value) - EPOCH).days)
    except (TypeError, ValueError):
        return None


def _contiguous_numbers(cells, start):
    values, index = [], start
    while isinstance(cells.get(index), (int, float)):
        values.append(float(cells[index]))
        index += 1
    return values


def read_reference(path):
    """Return (loans DataFrame, metadata) from a reference workbook Assets sheet."""
    rows, meta = [], {}
    with open_workbook(path) as workbook:
        with workbook.get_sheet('Assets') as sheet:
            for row in sheet.rows():
                cells = {cell.c: cell.v for cell in row if cell.v is not None}
                index = row[0].r if row else None
                if index is None:
                    continue
                if index == 0:
                    label = next((k for k, v in cells.items()
                                  if isinstance(v, str) and 'LoC FTP' in v), None)
                    if label is not None:
                        meta['loc_curve'] = _contiguous_numbers(cells, label + 1)
                    continue
                if index == 1:
                    label = next(k for k, v in cells.items()
                                 if isinstance(v, str) and 'Normal FTP Curve' in v)
                    meta['curve'] = _contiguous_numbers(cells, label + 1)
                    meta['first_day'] = float(cells.get(label - 6))
                    meta['last_day'] = float(cells.get(label - 5))
                    continue
                if index == 2:
                    label = next((k for k, v in cells.items()
                                  if isinstance(v, str) and 'Treasury FTP' in v), None)
                    if label is not None:
                        meta['treasury_curve'] = _contiguous_numbers(cells, label + 1)
                    continue
                if index == 4:
                    header = {k: v.strip() for k, v in cells.items() if isinstance(v, str)}

                    def find(predicate):
                        return next(k for k, v in header.items() if predicate(v.lower()))

                    columns = {
                        'branch': find(lambda v: v == 'branch code'),
                        'loan_ref': find(lambda v: v == 'loan ref'),
                        'exposure': find(lambda v: v.endswith('currency exposure+ currency accrued reporting')),
                        'booking': find(lambda v: v in ('booking date', 'booking_date')),
                        'maturity': find(lambda v: v == 'maturity date'),
                        'treasury_marker': find(lambda v: v == 'treasury marker'),
                        'sbu': find(lambda v: v == 'sbu'),
                        'dinm': find(lambda v: v == 'dinm'),
                        'dtm': find(lambda v: v == 'dtm'),
                        'mtm': find(lambda v: v == 'mtm'),
                        'charge': find(lambda v: v == 'ftp charge'),
                        'check': find(lambda v: v == 'check'),
                    }
                    columns['first_bucket'] = columns['mtm'] + 1
                    columns['bucket_count'] = (columns['check'] - columns['first_bucket']) // 2
                    meta['columns'] = columns
                    continue
                if index < 5 or 'columns' not in meta:
                    continue
                if cells.get(columns['loan_ref']) is None and cells.get(columns['exposure']) is None:
                    continue
                rows.append({
                    'branch': str(int(cells[columns['branch']])) if cells.get(columns['branch']) is not None else '',
                    'loan_ref': cells.get(columns['loan_ref']),
                    'exposure': float(cells.get(columns['exposure']) or 0.0),
                    'booking': cells.get(columns['booking']),
                    'maturity': cells.get(columns['maturity']),
                    'treasury': str(cells.get(columns['treasury_marker']) or '').strip(),
                    'sbu': str(cells.get(columns['sbu']) or '').strip(),
                    'dinm': cells.get(columns['dinm']),
                    'dtm': cells.get(columns['dtm']),
                    'mtm': cells.get(columns['mtm']),
                    'charge': float(cells.get(columns['charge']) or 0.0),
                })
    return pd.DataFrame(rows), meta


def read_processed(path, sheet_name):
    """Return the model output for one loan sheet, or None when the sheet is absent."""
    try:
        frame = pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        return None
    exposure_col = next((column for column in frame.columns
                         if str(column).strip().lower().startswith('currency exposure')), None)
    if exposure_col is None:
        return None
    frame = frame.rename(columns={exposure_col: 'exposure'})
    branch_col = next((column for column in frame.columns
                       if str(column).strip().lower() in ('branch code', 'branchcode')), None)
    frame['branch'] = frame[branch_col].map(
        lambda value: str(int(value)) if pd.notna(value) else '')
    frame['booking_serial'] = frame['BOOKING_DATE'].map(_serial)
    frame['maturity_serial'] = frame['MATURITY_DATE'].map(_serial)
    return frame


def reconcile(processed_path, reference_path, sheet_name, report_path=None):
    reference, meta = read_reference(reference_path)
    loans = reference[reference['treasury'] == ''].copy()
    model = read_processed(processed_path, sheet_name)
    if model is None:
        raise SystemExit(f'{sheet_name} not found in {processed_path}')

    def keys(frame, booking_col, maturity_col):
        return list(zip(frame['branch'], frame[booking_col], frame[maturity_col],
                        frame['exposure'].round(6)))

    model_keys = keys(model, 'booking_serial', 'maturity_serial')
    reference_keys = keys(loans, 'booking', 'maturity')

    model_charge = model['FTP Charge'].astype(float).tolist()
    model_dinm = model['DimDays'].tolist() if 'DimDays' in model.columns else [None] * len(model)
    model_dtm = model['DTM'].tolist() if 'DTM' in model.columns else [None] * len(model)
    model_mtm = model['MTM'].tolist() if 'MTM' in model.columns else [None] * len(model)

    # Group by key and compare the sorted charge multisets, so duplicate loans
    # (identical branch/date/exposure) never pair up out of order.
    model_groups = {}
    for position, key in enumerate(model_keys):
        model_groups.setdefault(key, []).append(
            (model_charge[position], model_dinm[position], model_dtm[position], model_mtm[position]))
    reference_groups = {}
    for key, charge, dinm, dtm, mtm in zip(reference_keys, loans['charge'],
                                           loans['dinm'], loans['dtm'], loans['mtm']):
        reference_groups.setdefault(key, []).append((float(charge), dinm, dtm, mtm))

    matched = mismatched = unmatched = 0
    model_total = reference_total = 0.0
    max_abs = max_rel = 0.0
    details = []
    for key, entries in model_groups.items():
        reference_entries = reference_groups.get(key)
        if not reference_entries:
            unmatched += len(entries)
            continue
        unmatched += abs(len(entries) - len(reference_entries))
        for model_entry, reference_entry in zip(sorted(entries), sorted(reference_entries)):
            charge, dinm, dtm, mtm = model_entry
            reference_charge, reference_dinm, reference_dtm, reference_mtm = reference_entry
            matched += 1
            model_total += charge
            reference_total += reference_charge
            delta = charge - reference_charge
            max_abs = max(max_abs, abs(delta))
            if reference_charge:
                max_rel = max(max_rel, abs(delta) / abs(reference_charge))
            if abs(delta) > max(MATCH_TOLERANCE_ABS, MATCH_TOLERANCE_REL * abs(reference_charge)):
                mismatched += 1
                details.append({
                    'Branch': key[0], 'Booking': key[1], 'Maturity': key[2], 'Exposure': key[3],
                    'Model FTP Charge': charge, 'Reference FTP Charge': reference_charge,
                    'Difference': delta,
                    'Model DimDays': dinm, 'Reference DinM': reference_dinm,
                    'Model DTM': dtm, 'Reference DTM': reference_dtm,
                    'Model MTM': mtm, 'Reference MTM': reference_mtm,
                })

    print(f'Sheet                : {sheet_name}')
    print(f'Reference workbook   : {os.path.basename(reference_path)}')
    print(f'Model output         : {os.path.basename(processed_path)}')
    print(f'Curve buckets        : {len(meta.get("curve") or [])} '
          f'(period {meta.get("first_day")}..{meta.get("last_day")})')
    print(f'Reference loans      : {len(loans):,}')
    print(f'Model rows           : {len(model):,}')
    print(f'Matched              : {matched:,}')
    print(f'Unmatched model rows : {unmatched:,}')
    print(f'Charge mismatches    : {mismatched:,}')
    print(f'Max absolute error   : {max_abs:.6e}')
    print(f'Max relative error   : {max_rel:.6e}')
    print(f'Reference charge     : {reference_total:,.2f}')
    print(f'Model charge         : {model_total:,.2f}')
    print(f'Difference           : {model_total - reference_total:,.2f}')

    if 'SBU' in model.columns:
        model_sbu = model.groupby('SBU')['FTP Charge'].sum().round(2)
        reference_sbu = loans.groupby('sbu')['charge'].sum().round(2)
        print('\nFTP charge by SBU')
        for sbu in sorted(set(model_sbu.index) | set(reference_sbu.index)):
            print(f'  {sbu or "<blank>":<24} reference {reference_sbu.get(sbu, 0.0):>16,.2f}'
                  f'   model {model_sbu.get(sbu, 0.0):>16,.2f}'
                  f'   diff {model_sbu.get(sbu, 0.0) - reference_sbu.get(sbu, 0.0):>14,.2f}')

    if report_path and details:
        pd.DataFrame(details).to_excel(report_path, index=False)
        print(f'\nMismatch detail written to {report_path}')

    return mismatched == 0 and unmatched == 0


def newest_processed_output(sheet_name):
    candidates = sorted(glob.glob(os.path.join('processed_outputs', 'FTP_Results_*.xlsx')))
    for candidate in reversed(candidates):
        try:
            if sheet_name in pd.ExcelFile(candidate).sheet_names:
                return candidate
        except (OSError, ValueError):
            continue
    return None


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--reference', required=True,
                        help='Reference CBZ workbook (.xlsb) with the Assets sheet')
    parser.add_argument('--sheet', default='FX LOANS', choices=['FX LOANS', 'ZWG LOANS'])
    parser.add_argument('--processed', help='Model output workbook (defaults to the newest in processed_outputs)')
    parser.add_argument('--report', help='Optional Excel file for the mismatch detail')
    args = parser.parse_args()

    processed = args.processed or newest_processed_output(args.sheet)
    if not processed:
        raise SystemExit('No processed output found - pass --processed')

    ok = reconcile(processed, args.reference, args.sheet, args.report)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
