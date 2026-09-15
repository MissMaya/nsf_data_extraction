"""
Extract abbreviation data from the Misión Abreviatura NSF workbook (renamed as 
mision_abreviatura.xlsx due to unicode issues).

The source workbook contains four worksheets:

    Procesal
    Encadenada
    Itálica cursiva
    Redonda

Each worksheet uses the following source columns:

    Documento
    Caligrafía
    Paleográfico
    Desatado

Confirmed mappings:

- `Paleográfico` is the historical abbreviation form.
- `Desatado` is the modern Spanish expansion.
- `Documento` identifies the source document in which the abbreviation occurs.
- `Caligrafía` identifies the calligraphy type.
- Worksheet and Excel row information are retained under `provenance`.

The workbook contains formatted blank rows beyond the actual data. These are
ignored.

The extractor:

1. reads all four source worksheets;
2. ignores blank rows;
3. converts each source row into canonical schema version 1.0.0;
4. writes one JSON file representing the complete Misión Abreviatura NSF
   source.

This script does not insert data into a database.
"""

# ---------------------------------------------------------------------------
# Import the required modules
# ---------------------------------------------------------------------------

import json
from pathlib import Path
from typing import Any
from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Paths and source configuration
# ---------------------------------------------------------------------------

# Script location:
#
#     abbreviations_db/
#         extractors/
#             extract_mision_nsf.py
#
# Therefore the parent of the extractors directory is abbreviations_db.

SCRIPT_PATH = Path(__file__).resolve()
EXTRACTORS_DIR = SCRIPT_PATH.parent
BASE_DIR = EXTRACTORS_DIR.parent

INGESTION_DIR = BASE_DIR / "ingestion"
RAW_DIR = INGESTION_DIR / "raw"
CANONICAL_OUTPUT_DIR = INGESTION_DIR / "canonical_output"

INPUT_XLSX = (
    RAW_DIR
    / "mision_abreviatura_nsf.xlsx"
)

OUTPUT_JSON = (
    CANONICAL_OUTPUT_DIR
    / "mision_abreviatura_nsf.canonical.json"
)

SCHEMA_NAME = "abbreviations_canonical"
SCHEMA_VERSION = "1.0"

SOURCE_DATASET_ID = "mision_abreviatura_nsf"
SOURCE_DATASET_NAME = "Misión Abreviatura NSF"

SOURCE_SHEETS = [
    "Procesal",
    "Encadenada",
    "Itálica cursiva",
    "Redonda",
]

EXPECTED_COLUMNS = {
    "Documento",
    "Caligrafía",
    "Paleográfico",
    "Desatado",
}


# ---------------------------------------------------------------------------
# Clean values
# ---------------------------------------------------------------------------

def clean(value: Any) -> str | None:
    """
    Convert a workbook value into a clean JSON-compatible string.

    Missing and empty values become None. All other values are converted to
    stripped strings.
    """
    if value is None:
        return None

    cleaned_value = str(value).strip()

    if not cleaned_value:
        return None

    return cleaned_value


# ---------------------------------------------------------------------------
# Workbook handling
# ---------------------------------------------------------------------------

def get_header_map(worksheet) -> dict[str, int]:
    """
    Map each column heading in Excel row 1 to its one-based column number.
    """
    header_map = {}

    for cell in worksheet[1]:
        heading = clean(cell.value)

        if heading:
            header_map[heading] = cell.column

    return header_map


def check_required_columns(
    worksheet_name: str,
    header_map: dict[str, int],
) -> None:
    """
    Confirm that all expected source columns exist in a worksheet.
    """
    missing_columns = sorted(
        EXPECTED_COLUMNS - set(header_map)
    )

    if missing_columns:
        raise ValueError(
            f"Worksheet '{worksheet_name}' is missing expected columns: "
            + ", ".join(missing_columns)
        )


def get_cell_value(
    worksheet,
    excel_row_number: int,
    column_name: str,
    header_map: dict[str, int],
) -> str | None:
    """
    Read and clean a cell value using its source column heading.
    """
    column_number = header_map[column_name]

    value = worksheet.cell(
        row=excel_row_number,
        column=column_number,
    ).value

    return clean(value)


# ---------------------------------------------------------------------------
# Construct canonical record
# ---------------------------------------------------------------------------

def make_record(
    worksheet,
    worksheet_name: str,
    excel_row_number: int,
    header_map: dict[str, int],
) -> dict[str, Any]:
    """
    Transform one Misión Abreviatura NSF row into a canonical record.
    """
    document_identifier = get_cell_value(
        worksheet,
        excel_row_number,
        "Documento",
        header_map,
    )

    calligraphy_type = get_cell_value(
        worksheet,
        excel_row_number,
        "Caligrafía",
        header_map,
    )

    abbreviation = get_cell_value(
        worksheet,
        excel_row_number,
        "Paleográfico",
        header_map,
    )

    modern_expansion = get_cell_value(
        worksheet,
        excel_row_number,
        "Desatado",
        header_map,
    )

    # Include worksheet and Excel row because row numbers restart on each
    # worksheet.
    source_record_id = (
        f"{SOURCE_DATASET_ID}__"
        f"{worksheet_name.replace(' ', '_').lower()}__"
        f"row_{excel_row_number}"
    )

    modern_expansions = []

    if modern_expansion:
        modern_expansions.append(
            {
                "value": modern_expansion,
                "source_label": "Desatado",
            }
        )

    return {
        "source_record_id": source_record_id,

        "abbreviation": {
            "value": abbreviation,
            "form_type": "historical",
        },

        # This source supplies modern Spanish expansions only.
        "expansions": {
            "historical": [],
            "modern": modern_expansions,
        },

        "occurrence": {
            # No separate collection information is supplied by the workbook.
            "collection": {
                "source_identifier": None,
                "name": None,
                "collection_type": None,
                "collection_status": None,
                "archival_institution": None,
                "path": None,
                "url": None,
            },

            # Preserve the source document identifier exactly as supplied.
            "document": {
                "source_identifier": document_identifier,
                "reference": document_identifier,
                "name": None,
                "page": None,
                "path": None,
                "url": None,
                "document_status": None,
            },

            # Unlike several of the other sources, this workbook explicitly
            # supplies calligraphy information.
            "calligraphy_type": calligraphy_type,

            "location": {
                "context_text": None,
                "region": None,
                "line": None,
                "word": None,
                "offset": None,
                "length": None,
            },
        },

        # This workbook does not provide image data.
        "images": {
            "source_document_images": [],
            "abbreviation_images": [],
        },

        "provenance": {
            "source_file": {
                "archive_name": None,
                "dataset_folder": None,
                "relative_file_path": str(
                    INPUT_XLSX.relative_to(BASE_DIR)
                ),
                "filename": INPUT_XLSX.name,
                "sheet_name": worksheet_name,
                "excel_row_number": excel_row_number,
            }
        },

        "source_specific_metadata": {},

        "notes": [],
    }


# ---------------------------------------------------------------------------
# Main extraction process
# ---------------------------------------------------------------------------

def extract() -> None:
    """
    Run the complete Misión Abreviatura NSF extraction process.
    """
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(
            f"Misión Abreviatura NSF source workbook not found: "
            f"{INPUT_XLSX}"
        )

    workbook = load_workbook(
        INPUT_XLSX,
        read_only=False,
        data_only=False,
    )

    missing_sheets = [
        sheet_name
        for sheet_name in SOURCE_SHEETS
        if sheet_name not in workbook.sheetnames
    ]

    if missing_sheets:
        raise ValueError(
            "The workbook is missing expected worksheets: "
            + ", ".join(missing_sheets)
        )

    records = []

    rows_inspected = 0
    blank_rows_skipped = 0
    records_without_abbreviation = 0
    records_without_modern_expansion = 0
    records_without_document = 0
    records_without_calligraphy = 0

    records_by_sheet = {}

    for worksheet_name in SOURCE_SHEETS:
        worksheet = workbook[worksheet_name]

        header_map = get_header_map(worksheet)

        check_required_columns(
            worksheet_name,
            header_map,
        )

        sheet_record_count = 0

        for excel_row_number in range(
            2,
            worksheet.max_row + 1,
        ):
            rows_inspected += 1

            document_identifier = get_cell_value(
                worksheet,
                excel_row_number,
                "Documento",
                header_map,
            )

            calligraphy_type = get_cell_value(
                worksheet,
                excel_row_number,
                "Caligrafía",
                header_map,
            )

            abbreviation = get_cell_value(
                worksheet,
                excel_row_number,
                "Paleográfico",
                header_map,
            )

            modern_expansion = get_cell_value(
                worksheet,
                excel_row_number,
                "Desatado",
                header_map,
            )

            # The worksheets contain many formatted blank rows so we ignore rows
            # where all four meaningful source fields are empty.
            if not any(
                [
                    document_identifier,
                    calligraphy_type,
                    abbreviation,
                    modern_expansion,
                ]
            ):
                blank_rows_skipped += 1
                continue

            # Preserve incomplete source records rather than silently dropping
            # them when one of the lexical fields is absent.
            if not abbreviation:
                records_without_abbreviation += 1

            if not modern_expansion:
                records_without_modern_expansion += 1

            if not document_identifier:
                records_without_document += 1

            if not calligraphy_type:
                records_without_calligraphy += 1

            record = make_record(
                worksheet = worksheet,
                worksheet_name = worksheet_name,
                excel_row_number = excel_row_number,
                header_map = header_map
            )

            records.append(record)
            sheet_record_count += 1

        records_by_sheet[worksheet_name] = sheet_record_count

    canonical_output = {
        "schema_name": SCHEMA_NAME,
        "schema_version": SCHEMA_VERSION,

        "source_dataset": {
            "dataset_id": SOURCE_DATASET_ID,
            "dataset_name": SOURCE_DATASET_NAME,

            "field_mappings": {
                "Paleográfico": {
                    "canonical_field": "abbreviation.value",
                    "meaning": "Historical abbreviation form",
                },
                "Desatado": {
                    "canonical_field": (
                        "expansions.modern[].value"
                    ),
                    "meaning": "Modern Spanish expansion",
                },
                "Documento": {
                    "canonical_field": (
                        "occurrence.document.source_identifier"
                    ),
                    "meaning": (
                        "Identifier of the source document containing "
                        "the abbreviation"
                    ),
                },
                "Caligrafía": {
                    "canonical_field": (
                        "occurrence.calligraphy_type"
                    ),
                    "meaning": "Calligraphy type",
                },
            },

            "mapping_notes": [
                (
                    "The source supplies modern Spanish expansions but "
                    "does not supply separate historical expansions."
                ),
                (
                    "The workbook is divided into four worksheets by "
                    "calligraphy type."
                ),
                (
                    "The Caligrafía value is retained from the source "
                    "rather than inferred from the worksheet name."
                ),
                (
                    "Rows containing no source data are ignored."
                    "Some worksheets contain empty rows that Excel" 
                    "still considers part of the used worksheet range."
                ),
                (
                    "Workbook, worksheet and Excel row information is "
                    "retained under provenance."
                ),
            ],
        },

        "records": records,
    }

    OUTPUT_JSON.parent.mkdir(
        parents = True,
        exist_ok = True,
    )

    with OUTPUT_JSON.open(
        "w",
        encoding = "utf-8",
    ) as output_file:
        json.dump(
            canonical_output,
            output_file,
            ensure_ascii = False,
            indent = 2,
        )

    workbook.close()

    print("\nExtraction complete")
    print("-------------------")

    for worksheet_name in SOURCE_SHEETS:
        print(
            f"{worksheet_name:<20} "
            f"{records_by_sheet[worksheet_name]:>5} records"
        )

    print("-------------------")
    print(f"Rows inspected:                     {rows_inspected}")
    print(f"Blank rows skipped:                 {blank_rows_skipped}")
    print(f"Canonical records:                  {len(records)}")
    print(
        "Records without abbreviation:      "
        f"{records_without_abbreviation}"
    )
    print(
        "Records without modern expansion:  "
        f"{records_without_modern_expansion}"
    )
    print(
        "Records without document:          "
        f"{records_without_document}"
    )
    print(
        "Records without calligraphy type:  "
        f"{records_without_calligraphy}"
    )
    print(f"Output:                             {OUTPUT_JSON}")


if __name__ == "__main__":
    extract()