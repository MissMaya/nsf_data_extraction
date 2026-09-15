"""
Extract abbreviation data from the Abreviaturas paleografía workbook.

Expected source workbook columns:

    Abreviatura
    Palabra completa
    Ejemplo
    Ubicación

Confirmed mappings:

- `Abreviatura` is the historical abbreviation form.
- `Palabra completa` is the modern Spanish expansion.
- `Ubicación` identifies the source document or location associated with
  the abbreviation.
- `Ejemplo` is retained as source-specific metadata. In the supplied
  workbook, the populated cells contain the text `Picture`, but the workbook
  contains no embedded images or image hyperlinks.
- Workbook and Excel row information are retained under `provenance`.

The extractor:

1. reads the source workbook;
2. converts each source row into canonical schema version 1.0.0;
3. writes one JSON file representing the complete source dataset.

This script does not insert data into a database.
"""

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
#             extract_paleografia.py
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
    / "Abreviaturas_paleografia(in).xlsx"
)

OUTPUT_JSON = (
    CANONICAL_OUTPUT_DIR
    / "abreviaturas_paleografia.canonical.json"
)

SCHEMA_NAME = "abbreviations_canonical"
SCHEMA_VERSION = "1.0"

SOURCE_DATASET_ID = "abreviaturas_paleografia"
SOURCE_DATASET_NAME = "Abreviaturas paleografía"

SOURCE_SHEET = "Abreviaturas_paleografia(in)"

EXPECTED_COLUMNS = {
    "Abreviatura",
    "Palabra completa",
    "Ejemplo",
    "Ubicación",
}


# ---------------------------------------------------------------------------
# Clean values
# ---------------------------------------------------------------------------

def clean(value: Any) -> str | None:
    """
    Convert a workbook value into a clean JSON-compatible string.

    Missing and empty values become None so that they are written to JSON as
    null. All other values are converted to stripped strings.
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


def check_required_columns(header_map: dict[str, int]) -> None:
    """
    Confirm that all expected source columns exist.
    """
    missing_columns = sorted(
        EXPECTED_COLUMNS - set(header_map)
    )

    if missing_columns:
        raise ValueError(
            "The Paleografía workbook is missing expected columns: "
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
    excel_row_number: int,
    header_map: dict[str, int],
) -> dict[str, Any]:
    """
    Transform one Paleografía workbook row into a canonical record.
    """
    abbreviation = get_cell_value(
        worksheet,
        excel_row_number,
        "Abreviatura",
        header_map,
    )

    modern_expansion = get_cell_value(
        worksheet,
        excel_row_number,
        "Palabra completa",
        header_map,
    )

    example_value = get_cell_value(
        worksheet,
        excel_row_number,
        "Ejemplo",
        header_map,
    )

    location_identifier = get_cell_value(
        worksheet,
        excel_row_number,
        "Ubicación",
        header_map,
    )

    source_record_id = (
        f"{SOURCE_DATASET_ID}__row_{excel_row_number}"
    )

    modern_expansions = []

    if modern_expansion:
        modern_expansions.append(
            {
                "value": modern_expansion,
                "source_label": "Palabra completa",
            }
        )

    return {
        "source_record_id": source_record_id,

        "abbreviation": {
            "value": abbreviation,
            "form_type": "historical",
        },

        # This source supplies only modern Spanish expansions.
        "expansions": {
            "historical": [],
            "modern": modern_expansions,
        },

        "occurrence": {
            # The source does not supply a separate collection field.
            "collection": {
                "source_identifier": None,
                "name": None,
                "collection_type": None,
                "collection_status": None,
                "archival_institution": None,
                "path": None,
                "url": None,
            },

            # Preserve `Ubicación` as the supplied source-document or
            # occurrence identifier without trying to parse or reinterpret it.
            "document": {
                "source_identifier": location_identifier,
                "reference": location_identifier,
                "name": None,
                "page": None,
                "path": None,
                "url": None,
                "document_status": None,
            },

            "calligraphy_type": None,

            "location": {
                "context_text": None,
                "region": None,
                "line": None,
                "word": None,
                "offset": None,
                "length": None,
            },
        },

        # The workbook contains no embedded images or image hyperlinks.
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
                "sheet_name": SOURCE_SHEET,
                "excel_row_number": excel_row_number,
            }
        },

        # Retain the source's `Ejemplo` value, including the four cells that
        # contain the text `Picture`. Do not treat these as image records
        # because the workbook contains no actual images or hyperlinks.
        "source_specific_metadata": {
            "example": example_value,
        },

        "notes": [],
    }


# ---------------------------------------------------------------------------
# Main extraction process
# ---------------------------------------------------------------------------

def extract() -> None:
    """
    Run the complete Paleografía extraction process.
    """
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(
            f"Paleografía source workbook not found: {INPUT_XLSX}"
        )

    workbook = load_workbook(
        INPUT_XLSX,
        read_only=False,
        data_only=False,
    )

    if SOURCE_SHEET not in workbook.sheetnames:
        raise ValueError(
            f"Worksheet '{SOURCE_SHEET}' not found in {INPUT_XLSX.name}. "
            f"Available worksheets: {workbook.sheetnames}"
        )

    worksheet = workbook[SOURCE_SHEET]

    header_map = get_header_map(worksheet)
    check_required_columns(header_map)

    records = []

    rows_seen = 0
    rows_skipped = 0
    records_without_abbreviation = 0
    records_without_modern_expansion = 0
    records_without_location = 0
    records_with_example_marker = 0

    # Excel row 1 contains the column headings.
    for excel_row_number in range(2, worksheet.max_row + 1):
        rows_seen += 1

        abbreviation = get_cell_value(
            worksheet,
            excel_row_number,
            "Abreviatura",
            header_map,
        )

        modern_expansion = get_cell_value(
            worksheet,
            excel_row_number,
            "Palabra completa",
            header_map,
        )

        location_identifier = get_cell_value(
            worksheet,
            excel_row_number,
            "Ubicación",
            header_map,
        )

        example_value = get_cell_value(
            worksheet,
            excel_row_number,
            "Ejemplo",
            header_map,
        )

        # Skip only rows where both core lexical fields are empty.
        if not abbreviation and not modern_expansion:
            rows_skipped += 1
            continue

        if not abbreviation:
            records_without_abbreviation += 1

        if not modern_expansion:
            records_without_modern_expansion += 1

        if not location_identifier:
            records_without_location += 1

        if example_value:
            records_with_example_marker += 1

        record = make_record(
            worksheet=worksheet,
            excel_row_number=excel_row_number,
            header_map=header_map,
        )

        records.append(record)

    canonical_output = {
        "schema_name": SCHEMA_NAME,
        "schema_version": SCHEMA_VERSION,

        "source_dataset": {
            "dataset_id": SOURCE_DATASET_ID,
            "dataset_name": SOURCE_DATASET_NAME,

            "field_mappings": {
                "Abreviatura": {
                    "canonical_field": "abbreviation.value",
                    "meaning": "Historical abbreviation form",
                },
                "Palabra completa": {
                    "canonical_field": (
                        "expansions.modern[].value"
                    ),
                    "meaning": "Modern Spanish expansion",
                },
                "Ubicación": {
                    "canonical_field": (
                        "occurrence.document.source_identifier"
                    ),
                    "meaning": (
                        "Source-document or occurrence identifier "
                        "supplied by the workbook"
                    ),
                },
                "Ejemplo": {
                    "canonical_field": (
                        "source_specific_metadata.example"
                    ),
                    "meaning": (
                        "Source example marker retained as supplied"
                    ),
                },
            },

            "mapping_notes": [
                (
                    "The source supplies modern Spanish expansions but "
                    "does not supply separate historical expansions."
                ),
                (
                    "Ubicación is retained without parsing or inferring "
                    "its internal archive and document components."
                ),
                (
                    "Four Ejemplo cells contain the text `Picture`, but "
                    "the workbook contains no embedded images or image "
                    "hyperlinks."
                ),
                (
                    "Workbook and Excel row information is retained "
                    "under provenance."
                ),
            ],
        },

        "records": records,
    }

    OUTPUT_JSON.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with OUTPUT_JSON.open(
        "w",
        encoding="utf-8",
    ) as output_file:
        json.dump(
            canonical_output,
            output_file,
            ensure_ascii=False,
            indent=2,
        )

    workbook.close()

    print("\nExtraction complete")
    print("-------------------")
    print(f"Source rows inspected:             {rows_seen}")
    print(f"Rows skipped:                      {rows_skipped}")
    print(f"Canonical records:                 {len(records)}")
    print(
        "Records without abbreviation:     "
        f"{records_without_abbreviation}"
    )
    print(
        "Records without modern expansion: "
        f"{records_without_modern_expansion}"
    )
    print(
        "Records without location:         "
        f"{records_without_location}"
    )
    print(
        "Records with Ejemplo value:       "
        f"{records_with_example_marker}"
    )
    print(f"Output:                            {OUTPUT_JSON}")


if __name__ == "__main__":
    extract()