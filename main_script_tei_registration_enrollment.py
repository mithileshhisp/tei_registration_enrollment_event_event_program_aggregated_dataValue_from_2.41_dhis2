##pip install pandas openpyxl requests python-dotenv
import os
import logging
from datetime import datetime, date

import pandas as pd
import requests
from dotenv import load_dotenv

# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

DHIS2_BASE_URL = os.getenv("DHIS2_BASE_URL", "").rstrip("/")
DHIS2_USERNAME = os.getenv("DHIS2_USERNAME", "")
DHIS2_PASSWORD = os.getenv("DHIS2_PASSWORD", "")

TRACKED_ENTITY_TYPE_UID = os.getenv(
    "TRACKED_ENTITY_TYPE_UID", ""
)

EXCEL_FILE = "tei_registration_enrollment_done.xlsx"
SHEET_NAME = "registration_enrollment"

REPORT_FILE = "dhis2_import_report.xlsx"

'''
VERIFY_SSL = (
    os.getenv("DHIS2_VERIFY_SSL", "true").lower() == "true"
)
'''

# Use CREATE to avoid overwriting existing tracked entities.
# Do not change to CREATE_AND_UPDATE unless you intend
# to update existing tracked entities.
IMPORT_STRATEGY = "CREATE"

# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(
            "dhis2_tracker_import.log",
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

# ============================================================
# VALIDATION
# ============================================================

if not DHIS2_BASE_URL:
    raise ValueError("DHIS2_BASE_URL is missing in .env")

if not DHIS2_USERNAME or not DHIS2_PASSWORD:
    raise ValueError("DHIS2 credentials are missing in .env")

if not TRACKED_ENTITY_TYPE_UID:
    raise ValueError(
        "TRACKED_ENTITY_TYPE_UID is missing in .env"
    )

# ============================================================
# DHIS2 SESSION
# ============================================================

session = requests.Session()

session.auth = (
    DHIS2_USERNAME,
    DHIS2_PASSWORD
)

session.headers.update({
    "Content-Type": "application/json",
    "Accept": "application/json"
})

#session.verify = VERIFY_SSL

TRACKER_IMPORT_URL = (
    f"{DHIS2_BASE_URL}/api/tracker"
)

# ============================================================
# EXCEL HELPERS
# ============================================================

METADATA_COLUMNS = {
    "tei_uid",
    "program_uid",
    "enrollmentDate",
    "ou",
    "orgunit_uid"
}


def clean_value(value):
    """
    Convert Excel values to JSON-compatible values.
    Return None for empty cells.
    """

    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, bool):
        return value

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))

        return str(value)

    return str(value).strip()


def format_date(value):
    """
    Convert Excel dates to DHIS2 YYYY-MM-DD format.
    """

    if pd.isna(value):
        raise ValueError("Enrollment date is empty")

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")

    value = str(value).strip()

    parsed_date = pd.to_datetime(
        value,
        errors="raise"
    )

    return parsed_date.strftime("%Y-%m-%d")


# ============================================================
# BUILD TRACKED ENTITY PAYLOAD
# ============================================================

def build_tracker_payload(row):

    tei_uid = clean_value(row["tei_uid"])
    program_uid = clean_value(row["program_uid"])
    orgunit_uid = clean_value(row["orgunit_uid"])

    if not tei_uid:
        raise ValueError("TEI UID is empty")

    if not program_uid:
        raise ValueError("Program UID is empty")

    if not orgunit_uid:
        raise ValueError("Organization unit UID is empty")

    enrollment_date = format_date(
        row["enrollmentDate"]
    )

    attributes = []

    # Every non-metadata column is treated as
    # a DHIS2 tracked entity attribute UID.
    for column in row.index:

        if column in METADATA_COLUMNS:
            continue

        attribute_uid = str(column).strip()

        value = clean_value(row[column])

        # Skip empty attribute values.
        if value is None or value == "":
            continue

        attributes.append({
            "attribute": attribute_uid,
            "value": value
        })

    # Create the tracked entity and its enrollment.
    tracker_payload = {
        "trackedEntities": [
            {
                "trackedEntity": tei_uid,

                "trackedEntityType": (
                    TRACKED_ENTITY_TYPE_UID
                ),

                "orgUnit": orgunit_uid,

                "attributes": attributes,

                "enrollments": [
                    {
                        "program": program_uid,

                        "orgUnit": orgunit_uid,

                        "enrolledAt": enrollment_date,

                        "status": "ACTIVE"
                    }
                ]
            }
        ]
    }

    return tracker_payload


# ============================================================
# IMPORT INTO DHIS2
# ============================================================

def import_tracker(payload):

    response = session.post(
        TRACKER_IMPORT_URL,
        params={
            "async": "false",
            "importStrategy": IMPORT_STRATEGY,
            "reportMode": "FULL",
            "validationMode": "FULL"
        },
        json=payload,
        timeout=120
    )

    if not response.ok:
        raise RuntimeError(
            f"HTTP {response.status_code}: "
            f"{response.text}"
        )

    return response.json()


# ============================================================
# EXTRACT IMPORT RESULT
# ============================================================

def extract_import_result(result):

    status = result.get("status", "UNKNOWN")

    stats = result.get("stats", {})

    if not stats:
        stats = result.get("bundleReport", {}).get(
            "stats", {}
        )

    return {
        "status": status,
        "created": stats.get("created", 0),
        "updated": stats.get("updated", 0),
        "deleted": stats.get("deleted", 0),
        "ignored": stats.get("ignored", 0),
        "total": stats.get("total", 0)
    }


# ============================================================
# MAIN PROCESS
# ============================================================

def main():

    logging.info("Starting DHIS2 Tracker import")

    # Read Excel.
    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        dtype=object
    )

    df.columns = df.columns.str.strip()

    required_columns = {
        "tei_uid",
        "program_uid",
        "enrollmentDate",
        "orgunit_uid"
    }

    missing_columns = (
        required_columns - set(df.columns)
    )

    if missing_columns:
        raise ValueError(
            f"Missing Excel columns: {missing_columns}"
        )

    # Prevent accidental duplicate UIDs in the workbook.
    df = df.dropna(subset=["tei_uid"])

    duplicate_mask = df["tei_uid"].duplicated(
        keep=False
    )

    if duplicate_mask.any():
        duplicates = (
            df.loc[duplicate_mask, "tei_uid"]
            .astype(str)
            .tolist()
        )

        raise ValueError(
            f"Duplicate TEI UIDs in Excel: {duplicates}"
        )

    report_rows = []

    total_rows = len(df)

    logging.info(
        "Total records to import: %s",
        total_rows
    )

    for index, row in df.iterrows():

        tei_uid = clean_value(row["tei_uid"])

        program_uid = clean_value(row["program_uid"])

        logging.info(
            "Processing %s/%s | TEI=%s",
            len(report_rows) + 1,
            total_rows,
            tei_uid
        )

        report = {
            "tei_uid": tei_uid,
            "program_uid": program_uid,
            "orgunit_uid": clean_value(
                row["orgunit_uid"]
            ),
            "status": "FAILED",
            "message": "",
            "created": 0,
            "updated": 0,
            "ignored": 0
        }

        try:

            payload = build_tracker_payload(row)

            result = import_tracker(payload)

            summary = extract_import_result(
                result
            )

            report.update(summary)

            report["message"] = str(result)

            if summary["status"] == "OK":
                report["status"] = "SUCCESS"

                logging.info(
                    "Imported TEI=%s successfully",
                    tei_uid
                )

            else:
                report["status"] = "FAILED"

                logging.error(
                    "DHIS2 import failed for TEI=%s: %s",
                    tei_uid,
                    result
                )

        except Exception as exc:

            report["status"] = "FAILED"
            report["message"] = str(exc)

            logging.exception(
                "Error importing TEI=%s",
                tei_uid
            )

        report_rows.append(report)

    # Save a report even if some records failed.
    report_df = pd.DataFrame(report_rows)

    report_df.to_excel(
        REPORT_FILE,
        index=False
    )

    success_count = (
        report_df["status"] == "SUCCESS"
    ).sum()

    failure_count = (
        report_df["status"] == "FAILED"
    ).sum()

    logging.info("Import completed")
    logging.info("Successful: %s", success_count)
    logging.info("Failed: %s", failure_count)
    logging.info("Report: %s", REPORT_FILE)


if __name__ == "__main__":
    main()