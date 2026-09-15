"""
Extract abbreviation data from the UNAM DicabeNovo scrape workbook.

Expected source workbook columns:

    abbr
    letter
    source
    old_spanish
    modern_spanish
    image_1
    ...
    image_14

Confirmed mappings:

- `abbr` is the historical abbreviation form.
- `old_spanish` is the historical Spanish expansion.
- `modern_spanish` is the modern Spanish expansion.
- `image_1` to `image_14` contain hyperlinks to images depicting the
  abbreviation itself.
- `letter` records the alphabetical section under which the entry appeared.
- Workbook and Excel row information are retained under `provenance`.

The extractor:

1. reads the source workbook;
2. extracts cell values and image hyperlinks;
3. converts each source row into canonical schema version 1.0.0;
4. writes one JSON file representing the complete UNAM source.

This script does not insert data into a database.
"""

import json
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell


# ---------------------------------------------------------------------------
# Paths and source configuration
# ---------------------------------------------------------------------------

# Script location:
#
#     abbreviations_db/
#         extractors/
#             extract_unam.py
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
    / "unam_dicabenovo_scrape_with_source.xlsx"
)

OUTPUT_JSON = (
    CANONICAL_OUTPUT_DIR
    / "unam_dicabenovo.canonical.json"
)

SCHEMA_NAME = "abbreviations_canonical"
SCHEMA_VERSION = "1.0"

SOURCE_DATASET_ID = "unam_dicabenovo"
SOURCE_DATASET_NAME = "UNAM DicabeNovo"

SOURCE_SHEET = "Sheet1"

CORE_COLUMNS = {
    "abbr",
    "letter",
    "source",
    "old_spanish",
    "modern_spanish",
}

IMAGE_COLUMNS = [
    f"image_{number}"
    for number in range(1, 15)
]


# ---------------------------------------------------------------------------
# Clean values
# ---------------------------------------------------------------------------

def clean(value: Any) -> str | None:
    """
    Convert a workbook value into a clean JSON-compatible string.

    Missing and empty values become None so that they are written to JSON
    as null. All other values are converted to stripped strings.
    """
    if value is None:
        return None

    cleaned_value = str(value).strip()

    if not cleaned_value:
        return None

    return cleaned_value


def filename_from_url(url: str | None) -> str | None:
    """
    Extract and decode the filename component of a URL.
    """
    if not url:
        return None

    parsed_url = urlparse(url)
    filename = Path(unquote(parsed_url.path)).name

    return filename or None


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
    Confirm that the expected core and image columns exist.
    """
    expected_columns = CORE_COLUMNS | set(IMAGE_COLUMNS)
    missing_columns = sorted(expected_columns - set(header_map))

    if missing_columns:
        raise ValueError(
            "The UNAM workbook is missing expected columns: "
            + ", ".join(missing_columns)
        )


def get_cell(
    worksheet,
    row_number: int,
    column_name: str,
    header_map: dict[str, int],
) -> Cell:
    """
    Return a worksheet cell using a source column name.
    """
    column_number = header_map[column_name]

    return worksheet.cell(
        row=row_number,
        column=column_number,
    )


def get_cell_value(
    worksheet,
    row_number: int,
    column_name: str,
    header_map: dict[str, int],
) -> str | None:
    """
    Read and clean a cell value using a source column name.
    """
    cell = get_cell(
        worksheet=worksheet,
        row_number=row_number,
        column_name=column_name,
        header_map=header_map,
    )

    return clean(cell.value)


def get_hyperlink(cell: Cell) -> str | None:
    """
    Return the target of an Excel hyperlink, where present.
    """
    if cell.hyperlink is None:
        return None

    return clean(cell.hyperlink.target)


# ---------------------------------------------------------------------------
# Construct canonical record
# ---------------------------------------------------------------------------

def extract_abbreviation_images(
    worksheet,
    excel_row_number: int,
    header_map: dict[str, int],
) -> list[dict[str, Any]]:
    """
    Extract all abbreviation-image hyperlinks from one workbook row.

    The visible cell value is usually the word `image`; the useful source
    information is the hyperlink target.
    """
    abbreviation_images = []

    for image_column in IMAGE_COLUMNS:
        cell = get_cell(
            worksheet=worksheet,
            row_number=excel_row_number,
            column_name=image_column,
            header_map=header_map,
        )

        hyperlink = get_hyperlink(cell)
        visible_value = clean(cell.value)

        # Skip genuinely empty image cells.
        if not hyperlink and not visible_value:
            continue

        abbreviation_images.append(
            {
                "source_identifier": hyperlink,
                "filename": filename_from_url(hyperlink),
                "path": None,
                "url": hyperlink,
                "calligraphy_type": None,
                "language": None,
                "source_label": image_column,
            }
        )

    return abbreviation_images


def make_record(
    worksheet,
    excel_row_number: int,
    header_map: dict[str, int],
) -> dict[str, Any]:
    """
    Transform one UNAM workbook row into a canonical record.
    """
    abbreviation = get_cell_value(
        worksheet,
        excel_row_number,
        "abbr",
        header_map,
    )

    historical_expansion = get_cell_value(
        worksheet,
        excel_row_number,
        "old_spanish",
        header_map,
    )

    modern_expansion = get_cell_value(
        worksheet,
        excel_row_number,
        "modern_spanish",
        header_map,
    )

    letter = get_cell_value(
        worksheet,
        excel_row_number,
        "letter",
        header_map,
    )

    source_label = get_cell_value(
        worksheet,
        excel_row_number,
        "source",
        header_map,
    )

    source_record_id = (
        f"{SOURCE_DATASET_ID}__row_{excel_row_number}"
    )

    historical_expansions = []

    if historical_expansion:
        historical_expansions.append(
            {
                "value": historical_expansion,
                "source_label": "old_spanish",
            }
        )

    modern_expansions = []

    if modern_expansion:
        modern_expansions.append(
            {
                "value": modern_expansion,
                "source_label": "modern_spanish",
            }
        )

    abbreviation_images = extract_abbreviation_images(
        worksheet=worksheet,
        excel_row_number=excel_row_number,
        header_map=header_map,
    )

    return {
        "source_record_id": source_record_id,

        "abbreviation": {
            "value": abbreviation,
            "form_type": "historical",
        },

        "expansions": {
            "historical": historical_expansions,
            "modern": modern_expansions,
        },

        # The scraped UNAM workbook does not provide a source-document
        # occurrence equivalent to the Transkribus collection/document
        # hierarchy.
        "occurrence": {
            "collection": {
                "source_identifier": source_label,
                "name": SOURCE_DATASET_NAME,
                "collection_type": None,
                "collection_status": None,
                "archival_institution": (
                    "Universidad Nacional Autónoma de México"
                ),
                "path": None,
                "url": None,
            },

            "document": {
                "source_identifier": None,
                "reference": None,
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

        "images": {
            # The workbook contains images of individual abbreviations rather
            # than general source-document or page images.
            "source_document_images": [],
            "abbreviation_images": abbreviation_images,
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

        # Preserve source fields that do not currently correspond to a common
        # canonical concept.
        "source_specific_metadata": {
            "alphabetical_section": letter,
            "source_label": source_label,
        },

        "notes": [],
    }


# ---------------------------------------------------------------------------
# Main extraction process
# ---------------------------------------------------------------------------

def extract() -> None:
    """
    Run the complete UNAM extraction process.
    """
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(
            f"UNAM source workbook not found: {INPUT_XLSX}"
        )

    # read_only=False is necessary because openpyxl does not expose cell
    # hyperlinks reliably in read-only mode.
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
    records_without_historical_expansion = 0
    records_without_modern_expansion = 0
    records_without_images = 0
    total_abbreviation_images = 0

    # Excel row 1 contains column headings.
    for excel_row_number in range(2, worksheet.max_row + 1):
        rows_seen += 1

        abbreviation = get_cell_value(
            worksheet,
            excel_row_number,
            "abbr",
            header_map,
        )

        historical_expansion = get_cell_value(
            worksheet,
            excel_row_number,
            "old_spanish",
            header_map,
        )

        modern_expansion = get_cell_value(
            worksheet,
            excel_row_number,
            "modern_spanish",
            header_map,
        )

        # Skip only completely empty core records.
        if (
            not abbreviation
            and not historical_expansion
            and not modern_expansion
        ):
            rows_skipped += 1
            continue

        record = make_record(
            worksheet=worksheet,
            excel_row_number=excel_row_number,
            header_map=header_map,
        )

        if not abbreviation:
            records_without_abbreviation += 1

        if not historical_expansion:
            records_without_historical_expansion += 1

        if not modern_expansion:
            records_without_modern_expansion += 1

        image_count = len(
            record["images"]["abbreviation_images"]
        )

        total_abbreviation_images += image_count

        if image_count == 0:
            records_without_images += 1

        records.append(record)

    canonical_output = {
        "schema_name": SCHEMA_NAME,
        "schema_version": SCHEMA_VERSION,

        "source_dataset": {
            "dataset_id": SOURCE_DATASET_ID,
            "dataset_name": SOURCE_DATASET_NAME,

            "field_mappings": {
                "abbr": {
                    "canonical_field": "abbreviation.value",
                    "meaning": "Historical abbreviation form",
                },
                "old_spanish": {
                    "canonical_field": (
                        "expansions.historical[].value"
                    ),
                    "meaning": "Historical Spanish expansion",
                },
                "modern_spanish": {
                    "canonical_field": (
                        "expansions.modern[].value"
                    ),
                    "meaning": "Modern Spanish expansion",
                },
                "image_1 ... image_14": {
                    "canonical_field": (
                        "images.abbreviation_images[]"
                    ),
                    "meaning": (
                        "Hyperlinks to images depicting the "
                        "abbreviation"
                    ),
                },
                "letter": {
                    "canonical_field": (
                        "source_specific_metadata."
                        "alphabetical_section"
                    ),
                    "meaning": (
                        "Alphabetical section of the DicabeNovo source"
                    ),
                },
                "source": {
                    "canonical_field": (
                        "source_specific_metadata.source_label"
                    ),
                    "meaning": "Source label supplied by the scrape",
                },
            },

            "mapping_notes": [
                (
                    "The image columns contain Excel hyperlinks. "
                    "The hyperlink target is retained as the image URL."
                ),
                (
                    "UNAM is the only current source that supplies both "
                    "historical and modern expansions."
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
    print(f"Source rows inspected:                 {rows_seen}")
    print(f"Rows skipped:                          {rows_skipped}")
    print(f"Canonical records:                     {len(records)}")
    print(
        "Records without abbreviation:         "
        f"{records_without_abbreviation}"
    )
    print(
        "Records without historical expansion: "
        f"{records_without_historical_expansion}"
    )
    print(
        "Records without modern expansion:     "
        f"{records_without_modern_expansion}"
    )
    print(
        "Records without abbreviation images:  "
        f"{records_without_images}"
    )
    print(
        "Total abbreviation-image links:        "
        f"{total_abbreviation_images}"
    )
    print(f"Output:                                {OUTPUT_JSON}")


if __name__ == "__main__":
    extract()