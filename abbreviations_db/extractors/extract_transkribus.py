"""
Extract abbreviation data from the nested Transkribus dataset.

The source dataset is contained in a ZIP archive with the following structure:

    abreviaturas_Transkribus.zip
    └── <collection_folder>/
        └── <7_digit_document_folder>/
            └── <workbook>.xlsx

Individual workbooks may be empty. Workbooks containing abbreviation data should 
have an `abbrev` worksheet.

This extractor:

1. unzips the source archive into a working directory;
2. recursively discovers all `.xlsx` files;
3. reads the `abbrev` worksheet where present;
4. skips empty workbooks and rows without abbreviation or expansion data;
5. converts each source row into canonical schema version 1.0;
6. writes one JSON file representing the complete Transkribus source.

This script does not insert data into a database.
"""

# ---------------------------------------------------------------------------
# Required modules
# ---------------------------------------------------------------------------

import json
import shutil
import zipfile
import pandas as pd
from pathlib import Path
from typing import Any


# ---------------------------------------------------------------------------
# Paths and source configuration
# ---------------------------------------------------------------------------

# Expected path to script:
#
#     abbreviations_db/
#         extractors/
#             extract_transkribus.py
#
BASE_DIR = Path(__file__).resolve().parents[1]

INGESTION_DIR = BASE_DIR / "ingestion"
RAW_DIR = INGESTION_DIR / "raw"
WORKING_DIR = INGESTION_DIR / "working"
CANONICAL_OUTPUT_DIR = INGESTION_DIR / "canonical_output"

ZIP_PATH = RAW_DIR / "abreviaturas_Transkribus.zip"

EXTRACT_DIR = (
    WORKING_DIR
    / "abreviaturas_Transkribus"
)

OUTPUT_JSON = (
    CANONICAL_OUTPUT_DIR
    / "transkribus_nested.canonical.json"
)

SCHEMA_NAME = "abbreviations_canonical"
SCHEMA_VERSION = "1.0"

SOURCE_DATASET_ID = "transkribus_nested"
SOURCE_DATASET_NAME = "abreviaturas_Transkribus"

# Worksheet containing the already-separated abbreviation data.
ABBREV_SHEET = "abbrev"


# ---------------------------------------------------------------------------
# Clean values
# ---------------------------------------------------------------------------

def clean(value: Any) -> str | None:
    """
    Convert a spreadsheet value into a clean JSON-compatible string.

    Missing values, including pandas NaN values and empty strings, become
    None so that they are written to JSON as null.

    All other values are converted to stripped strings.
    """
    if pd.isna(value):
        return None

    cleaned_value = str(value).strip()

    if not cleaned_value:
        return None

    return cleaned_value


def clean_int(value: Any) -> int | str | None:
    """
    Convert a spreadsheet value to an integer where possible.

    Fields such as `offset` and `length` should normally be numeric. If a
    value cannot be converted safely, it is preserved as cleaned text rather than
    being discarded.

    Numeric spreadsheet values are converted safely to integers.
    """
    if pd.isna(value):
        return None

    try:
        numeric_value = float(value)

        if numeric_value.is_integer():
            return int(numeric_value)

    except (TypeError, ValueError):
        pass

    return clean(value)


# ---------------------------------------------------------------------------
# Archive handling
# ---------------------------------------------------------------------------

def unzip_source(refresh: bool = False) -> None:
    """
    Extract the Transkribus ZIP archive into the working directory.

    By default, reuse an existing extracted directory. Set `refresh = True` to
    delete and recreate the extracted directory from the current ZIP archive.
    """
    if not ZIP_PATH.exists():
        raise FileNotFoundError(
            f"Transkribus source archive not found: {ZIP_PATH}"
        )

    if refresh and EXTRACT_DIR.exists():
        shutil.rmtree(EXTRACT_DIR)

    if EXTRACT_DIR.exists():
        print(f"Using existing extracted data: {EXTRACT_DIR}")
        return

    EXTRACT_DIR.mkdir(
        parents = True,
        exist_ok = True,
    )

    with zipfile.ZipFile(ZIP_PATH, "r") as zip_ref:
        zip_ref.extractall(EXTRACT_DIR)

    print(f"Extracted source archive to: {EXTRACT_DIR}")


# ---------------------------------------------------------------------------
# Source structure
# ---------------------------------------------------------------------------

def get_folder_metadata(xlsx_path: Path) -> dict[str, str | None]:
    """
    Get collection, document and provenance metadata from an XLSX file path.

    Expected structure relative to EXTRACT_DIR:

        <collection_folder>/
            <7_digit_document_folder>/
                <file>.xlsx

    Depending on how the ZIP archive was created, an additional top-level
    dataset folder may also be present:

        abreviaturas_Transkribus/
            <collection_folder>/
                <7_digit_document_folder>/
                    <file>.xlsx

    Collection and document folder names are retained because they provide
    source identifiers that may later be mapped to database entities.
    """
    relative_path = xlsx_path.relative_to(EXTRACT_DIR)
    parts = relative_path.parts

    if parts and parts[0] == SOURCE_DATASET_NAME:
        dataset_folder = parts[0]
        collection_folder = parts[1] if len(parts) > 1 else None
        document_folder = parts[2] if len(parts) > 2 else None
    else:
        dataset_folder = SOURCE_DATASET_NAME
        collection_folder = parts[0] if len(parts) > 0 else None
        document_folder = parts[1] if len(parts) > 1 else None

    collection_path = None

    if collection_folder:
        collection_path = str(
            Path(dataset_folder)
            / collection_folder
        )

    document_path = None

    if collection_folder and document_folder:
        document_path = str(
            Path(dataset_folder)
            / collection_folder
            / document_folder
        )

    return {
        "relative_xlsx_path": str(relative_path),
        "dataset_folder": dataset_folder,
        "collection_folder": collection_folder,
        "document_folder": document_folder,
        "collection_path": collection_path,
        "document_path": document_path,
    }


# ---------------------------------------------------------------------------
# Read workbook
# ---------------------------------------------------------------------------

def read_abbrev_sheet(xlsx_path: Path) -> pd.DataFrame | None:
    """
    Read the `abbrev` worksheet from a workbook.

    Workbooks without an `abbrev` worksheet are skipped as some source texts 
    contain no abbreviations.
    """
    try:
        return pd.read_excel(
            xlsx_path,
            sheet_name = ABBREV_SHEET,
        )

    except ValueError:
        print(
            f"Skipping workbook without '{ABBREV_SHEET}' sheet: "
            f"{xlsx_path}"
        )
        return None

    except Exception as error:
        print(f"Could not read workbook {xlsx_path}: {error}")
        return None


def workbook_has_data(df: pd.DataFrame | None) -> bool:
    """
    Check whether a worksheet contains usable abbreviation data.

    A worksheet is considered usable when:

    - it exists;
    - it contains at least one row;
    - it contains either `Value` or `expansion`;
    - at least one row contains a non-empty abbreviation or expansion.

    Incomplete source rows are preserved for later inspection.
    """
    if df is None or df.empty:
        return False

    core_columns = [
        column
        for column in ("Value", "expansion")
        if column in df.columns
    ]

    if not core_columns:
        return False

    for column in core_columns:
        cleaned_values = df[column].apply(clean)

        if not cleaned_values.dropna().empty:
            return True

    return False


# ---------------------------------------------------------------------------
# Construct canonical record
# ---------------------------------------------------------------------------

def make_record(
    row: pd.Series,
    excel_row_number: int,
    xlsx_path: Path,
) -> dict[str, Any]:
    """
    Transform one Transkribus abbreviation row into a canonical record.

    Mappings:

    - `Value` is the historical abbreviation form.
    - `expansion` is the modern Spanish expansion.
    - `Imagename` identifies the source document image from which the
      abbreviation was taken.
    - `Doc` and the seven-digit document folder are distinct source values.
    - Transkribus-specific positional metadata is retained under
      `occurrence.location`.
    - Workbook and row information is retained under `provenance`.

    """
    folder_meta = get_folder_metadata(xlsx_path)

    abbreviation = clean(row.get("Value"))
    modern_expansion = clean(row.get("expansion"))

    collection_identifier = folder_meta["collection_folder"]
    document_folder = folder_meta["document_folder"]
    document_reference = clean(row.get("Doc"))
    source_image_filename = clean(row.get("Imagename"))
    page = clean(row.get("Page"))

    # Stable and traceable identifier for this extracted source row.
    # Not intended to become a database primary key.

    source_record_id = (
        f"{SOURCE_DATASET_ID}__"
        f"{collection_identifier or 'unknown_collection'}__"
        f"{document_folder or 'unknown_document'}__"
        f"{xlsx_path.stem}__"
        f"row_{excel_row_number}"
    )

    historical_expansions: list[dict[str, Any]] = []
    modern_expansions: list[dict[str, Any]] = []

    if modern_expansion:
        modern_expansions.append(
            {
                "value": modern_expansion,
                "source_label": "expansion",
            }
        )

    source_document_images: list[dict[str, Any]] = []

    if source_image_filename:
        source_document_images.append(
            {
                "source_identifier": source_image_filename,
                "filename": source_image_filename,
                "path": None,
                "url": None,
                "calligraphy_type": None,
                "language": None,
            }
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

        "occurrence": {
            "collection": {
                "source_identifier": collection_identifier,
                "name": collection_identifier,
                "collection_type": None,
                "collection_status": None,
                "archival_institution": None,
                "path": folder_meta["collection_path"],
                "url": None,
            },

            "document": {
                "source_identifier": document_folder,
                "reference": document_reference,
                "name": None,
                "page": page,
                "path": folder_meta["document_path"],
                "url": None,
                "document_status": None,
            },

            "calligraphy_type": None,

            "location": {
                "context_text": clean(row.get("Context")),
                "region": clean(row.get("Region")),
                "line": clean(row.get("Line")),
                "word": clean(row.get("Word")),
                "offset": clean_int(row.get("offset")),
                "length": clean_int(row.get("length")),
            },
        },

        "images": {
            "source_document_images": source_document_images,
            "abbreviation_images": [],
        },

        "provenance": {
            "source_file": {
                "archive_name": ZIP_PATH.name,
                "dataset_folder": folder_meta["dataset_folder"],
                "relative_file_path": folder_meta[
                    "relative_xlsx_path"
                ],
                "filename": xlsx_path.name,
                "sheet_name": ABBREV_SHEET,
                "excel_row_number": excel_row_number,
            }
        },

        "source_specific_metadata": {},

        "notes": [],
    }


# ---------------------------------------------------------------------------
# Main extraction process
# ---------------------------------------------------------------------------

def extract(refresh_extracted_data: bool = False) -> None:
    """
    Run the complete Transkribus extraction process.

    Recursively process every XLSX workbook in the extracted directory and
    write all canonical records to one versioned JSON file.

    Rows are skipped only when both the abbreviation and modern expansion are
    empty so that we can still inspect incomplete records later if necessary. 
    """
    unzip_source(refresh = refresh_extracted_data)

    records: list[dict[str, Any]] = []

    files_seen = 0
    files_with_data = 0
    files_without_data = 0
    rows_seen = 0
    rows_skipped = 0
    records_without_abbreviation = 0
    records_without_modern_expansion = 0

    xlsx_paths = sorted(EXTRACT_DIR.rglob("*.xlsx"))

    for xlsx_path in xlsx_paths:
        files_seen += 1

        df = read_abbrev_sheet(xlsx_path)

        if not workbook_has_data(df):
            files_without_data += 1
            continue

        files_with_data += 1

        # The type checker cannot infer that df is no longer None after
        # workbook_has_data(), so assert it explicitly.
        assert df is not None

        for index, row in df.iterrows():
            rows_seen += 1

            abbreviation = clean(row.get("Value"))
            modern_expansion = clean(row.get("expansion"))

            if not abbreviation and not modern_expansion:
                rows_skipped += 1
                continue

            if not abbreviation:
                records_without_abbreviation += 1

            if not modern_expansion:
                records_without_modern_expansion += 1

            # Convert the zero-based DataFrame index to the original Excel row
            # number. Excel row 1 contains the column headings.
            excel_row_number = int(index) + 2

            record = make_record(
                row=row,
                excel_row_number=excel_row_number,
                xlsx_path=xlsx_path,
            )

            records.append(record)

    canonical_output = {
        "schema_name": SCHEMA_NAME,
        "schema_version": SCHEMA_VERSION,

        "source_dataset": {
            "dataset_id": SOURCE_DATASET_ID,
            "dataset_name": SOURCE_DATASET_NAME,
        },

        "records": records,
    }

    OUTPUT_JSON.parent.mkdir(
        parents = True,
        exist_ok = True,
    )

    with OUTPUT_JSON.open(
        "w",
        encoding="utf-8",
    ) as output_file:
        json.dump(
            canonical_output,
            output_file,
            ensure_ascii = False,
            indent = 2,
        )

    print("\nExtraction complete")
    print("-------------------")
    print(f"XLSX files found:                  {files_seen}")
    print(f"XLSX files with data:              {files_with_data}")
    print(f"XLSX files without usable data:    {files_without_data}")
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
    print(f"Output:                            {OUTPUT_JSON}")


if __name__ == "__main__":
    extract()