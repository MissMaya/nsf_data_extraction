"""
Extract the contents of DECM_16C_Final_Gazetteer.zip into a JSON.
"""

# ---------------------------------------------------------------------------
# Import the required modules
# ---------------------------------------------------------------------------
from __future__ import annotations
import argparse
import csv
import io
import json
import re
import sys
import unicodedata
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from openpyxl import load_workbook
from datetime import date, time

# ---------------------------------------------------------------------------
# Set up the paths
# ---------------------------------------------------------------------------
SCRIPT_PATH = Path(__file__).resolve()
EXTRACTORS_DIR = SCRIPT_PATH.parent
BASE_DIR = EXTRACTORS_DIR.parent

INGESTION_DIR = BASE_DIR / "ingestion"
RAW_DIR = INGESTION_DIR / "raw"
CANONICAL_OUTPUT_DIR = INGESTION_DIR / "canonical_output"

ZIP_PATH = RAW_DIR / "DECM_16C_Final_Gazetteer.zip"
OUTPUT_JSON = CANONICAL_OUTPUT_DIR / "decm_16c_toponyms.canonical.json"

# ---------------------------------------------------------------------------
# Helpers to read CSV and Excel source files
# ---------------------------------------------------------------------------

def excel_value(cell):
    """Return a JSON-compatible cell value."""
    value = cell.value
    if value is None:
        return None
    if cell.data_type == 'f':
        return {'formula': value, 'cached_value': None}
    if cell.data_type == 'e':
        return {'excel_error': value}
    if isinstance(value, (datetime, date, time)):
        # openpyxl has already interpreted the Excel date formatting.
        return value.isoformat()
    return value


def xlsx_tables(data):
    """Read each worksheet with openpyxl."""
    workbook = load_workbook(io.BytesIO(data), read_only = False, data_only = False)
    try:
        for sheet in workbook.worksheets:
            rows = []
            for row in sheet.iter_rows():
                cells = {cell.column - 1: excel_value(cell)
                         for cell in row if cell.value is not None}
                if cells:
                    rows.append((row[0].row, cells))
            yield sheet.title, rows
    finally:
        workbook.close()

def csv_tables(data):
    # Try common CSV encodings without silently replacing unreadable characters
    decoded = None
    for encoding in ('utf-8-sig', 'utf-16', 'cp1252'):
        try:
            candidate = data.decode(encoding)
            if '\x00' in candidate:
                continue
            decoded = candidate
            break
        except UnicodeError:
            pass
    if decoded is None:
        raise ValueError('Cannot decode CSV as UTF-8, UTF-16 or Windows-1252')
    sample = decoded[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
    except csv.Error:
        dialect = csv.excel
    reader = csv.reader(io.StringIO(decoded, newline = ''), dialect)
    yield None, [(reader.line_num, {i: v for i, v in enumerate(row)}) for row in reader]

# ---------------------------------------------------------------------------
# Classify tables whilst recording folder name for provenance
# ---------------------------------------------------------------------------

def classify(filename, sheet):
    """Classify by worksheet name first; filename only for CSV and gazetteer sheets."""
    def fold(value):
        value = unicodedata.normalize('NFKD', value.lower())
        return ''.join(c for c in value if not unicodedata.combining(c))
    name, tab = fold(filename), fold(sheet or '')
    if sheet is not None:
        if 'indice de documentos' in tab:
            return 'document_index', None
        if 'indice de toponimos' in tab:
            return 'toponym_index', None
        if 'gazetteer' in tab or 'gazetter' in tab or (name.startswith('1_2_') and tab == 'hoja1'):
            return 'places', 'unspecified'
        return 'additional_sheets', None
    if 'polygons' in name or 'poligono' in name:
        return 'places', 'polygons'
    if 'localidades' in name or 'gazetteer' in name or 'gazetter' in name or name.startswith('1_guatemala'):
        return 'places', 'localidades' if 'localidades' in name else 'unspecified'
    return 'additional_sheets', None

def split_folder(folder):
    # Identify the numbered prefix so folders can be sorted numerically.
    match = re.match(r'^(\d+(?:_\d+)*)_(.+)$', folder)
    return (match.group(1), match.group(2)) if match else (None, folder)


def archive_sort_key(item):
    """Sort numbered source folders numerically, then files by name."""
    parts = item.filename.split('/')
    folder = parts[1] if len(parts) > 1 else ''
    folder_id, _ = split_folder(folder)
    numbers = tuple(int(part) for part in folder_id.split('_')) if folder_id else (float('inf'),)
    return numbers, item.filename

# ---------------------------------------------------------------------------
# Convert source rows to canonical records
# ---------------------------------------------------------------------------

def convert_table(rows):
    """Keep source headings and create readable, uniquely named row fields."""
    if not rows:
        return [], {'columns': [], 'field_keys': [], 'header_row_number': None,
                    'record_count': 0}

    # Treat the first non-empty row as the header.
    header_pos = next((i for i, (_, cells) in enumerate(rows)
                       if any(v is not None and str(v).strip() for v in cells.values())), None)
    if header_pos is None:
        return [], {'columns': [], 'field_keys': [], 'header_row_number': None,
                    'record_count': 0}

    header_row, header_cells = rows[header_pos]
    data_rows = [(number, cells) for number, cells in rows[header_pos + 1:]
                 if any(v is not None and str(v).strip() for v in cells.values())]

    # Include data columns beyond the header; retain original headings separately.
    width = max((max(cells, default=-1) + 1 for _, cells in
                 [(header_row, header_cells), *data_rows]), default=0)
    columns = [header_cells.get(i) for i in range(width)]

    # Use readable keys; suffix duplicate headings and name blank headings by position.
    # Reserve original headings first, so generated keys cannot hide a source heading.
    original_keys = {str(h) for h in columns if h is not None and str(h).strip()}
    used = set()
    field_keys = []
    for i, heading in enumerate(columns, start=1):
        base = str(heading) if heading is not None and str(heading).strip() else f'column_{i}'
        key = base
        suffix = 2
        while key in used or (key != base and key in original_keys) or (key == base and base not in original_keys and key in original_keys):
            key = f'{base}__{suffix}'
            suffix += 1
        used.add(key)
        field_keys.append(key)

    records = [{'row_number': number,
                'fields': {key: cells.get(i) for i, key in enumerate(field_keys)}}
               for number, cells in data_rows]
    return records, {'columns': columns, 'field_keys': field_keys,
                     'header_row_number': header_row, 'record_count': len(records)}

# ---------------------------------------------------------------------------
# Extract the complete archive and write out the JSON
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--zip', type = Path, default = ZIP_PATH,
                        help = 'Input DECM gazetteer ZIP (default: ingestion/raw)')
    parser.add_argument('--output', type = Path, default = OUTPUT_JSON,
                        help = 'Canonical JSON output (default: ingestion/canonical_output)')
    args = parser.parse_args()
    result = {
        'schema_name': 'toponyms_canonical',
        'schema_version': '1.0',
        'source_dataset': {
            'archive_filename': args.zip.name,
            'description': 'DECM 16C Final Gazetteer',
            'extraction_timestamp_utc': datetime.now(timezone.utc).isoformat(),
        },
        'mapping_notes': [
            'This JSON is an intermediate extraction intended to support the later population of a database.',
            'source_tables contains one entry per CSV file or Excel worksheet, ordered by numbered source folder.',
            'archive_member_path identifies the original file inside the ZIP; sheet_name identifies the worksheet (this is null for CSVs).',
            'table_type is a descriptive label derived from the source filename or worksheet name.',
            'columns retains original headings in column order; field_keys lists the corresponding unique JSON keys in that same order.',
            'Each record contains its original row_number and readable fields as key-value pairs.',
            'Duplicate headings receive suffixes such as My_FID__2; blank headings receive keys such as column_3.', 
            'Use the columns and field_keys that appear at the top of the file to recover the original heading and position.',
            'The first non-empty row is treated as the header and entirely empty rows are skipped.',
            'CSV values remain text, including empty strings. Missing Excel cells are treated as null; Excel numbers remain numbers and Excel date objects are ISO strings.',
        ],
        'source_tables': [],
        'issues': [],
    }
    with zipfile.ZipFile(args.zip) as archive:
        entries = sorted((item for item in archive.infolist() if not item.is_dir()
                          and not item.filename.startswith('__MACOSX/')
                          and not any(part.startswith('.') for part in item.filename.split('/'))),
                         key = archive_sort_key)
        for item in entries:
            parts = item.filename.split('/')
            if len(parts) < 3:
                result['issues'].append({'path': item.filename, 'issue': 'unexpected_path'})
                continue
            folder, filename = parts[1], parts[-1]
            if not filename.lower().endswith(('.csv', '.xlsx')):
                result['issues'].append({'path': item.filename, 'issue': 'unsupported_file'})
                continue
            data = archive.read(item)
            try:
                tables = csv_tables(data) if filename.lower().endswith('.csv') else xlsx_tables(data)
                for sheet, rows in tables:
                    collection, subtype = classify(filename, sheet)
                    table_type = (subtype if subtype != 'unspecified' else 'gazetteer') if collection == 'places' else (collection if collection != 'additional_sheets' else 'additional_sheet')
                    records, info = convert_table(rows)
                    result['source_tables'].append({
                        'archive_member_path': item.filename,
                        'sheet_name': sheet,
                        'table_type': table_type,
                        **info,
                        'records': records,
                    })
            except Exception as exc:
                result['issues'].append({'path': item.filename, 'issue': 'extraction_error',
                                         'detail': str(exc)})
    if result['issues']:
        for issue in result['issues']:
            print('WARNING:', issue, file = sys.stderr)
        print('Extraction reported issues; output was not overwritten.', file = sys.stderr)
        return 1
    args.output.parent.mkdir(parents = True, exist_ok = True)
    with args.output.open('w', encoding = 'utf-8') as handle:
        json.dump(result, handle, ensure_ascii = False, indent = 2)
        handle.write('\n')
    print('Output:', args.output)
    print('Source tables:', len(result['source_tables']), 'Issues:', len(result['issues']))
    print('Extracted records:', sum(table['record_count'] for table in result['source_tables']))
    return 0

if __name__ == '__main__':
    sys.exit(main())
