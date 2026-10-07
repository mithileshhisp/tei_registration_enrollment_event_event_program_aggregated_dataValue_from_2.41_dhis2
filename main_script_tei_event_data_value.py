
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

DHIS2_BASE_URL = os.getenv(
    "DHIS2_BASE_URL", ""
).rstrip("/")

DHIS2_USERNAME = os.getenv("DHIS2_USERNAME", "")
DHIS2_PASSWORD = os.getenv("DHIS2_PASSWORD", "")

VERIFY_SSL = (
    os.getenv("DHIS2_VERIFY_SSL", "true").lower()
    == "true"
)

EXCEL_FILE = "tei_event_with_data_value_stage8.xlsx"
SHEET_NAME = "event_with_data_value"

REPORT_FILE = "dhis2_event_import_report.xlsx"

# Set False to test payloads without sending them.
DRY_RUN = False

# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(
            "dhis2_event_import.log",
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

# ============================================================
# VALIDATION
# ============================================================

if not DHIS2_BASE_URL:
    raise ValueError("DHIS2_BASE_URL is missing")

if not DHIS2_USERNAME or not DHIS2_PASSWORD:
    raise ValueError("DHIS2 credentials are missing")

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

session.verify = VERIFY_SSL

TRACKER_URL = f"{DHIS2_BASE_URL}/api/tracker"

# ============================================================
# EXCEL COLUMN CONFIGURATION
# ============================================================

METADATA_COLUMNS = {
    "event_uid",
    "tei_uid",
    "orgunit_uid",
    "program_uid",
    "program_stage_uid",
    "eventDate",
    "GPS Latitude",
    "GPS Longitude"
}

REQUIRED_COLUMNS = {
    "event_uid",
    "tei_uid",
    "orgunit_uid",
    "program_uid",
    "program_stage_uid",
    "eventDate"
}


# ============================================================
# date conversion function
# ============================================================
def format_data_element_date(value):

    if is_empty(value):
        return None

    parsed_date = pd.to_datetime(
        value,
        errors="raise"
    )

    return parsed_date.strftime("%Y-%m-%d")

# ============================================================
# retrieve the enrollment UID
# ============================================================
def get_enrollment_uid(tei_uid, program_uid):

    url = (
        f"{DHIS2_BASE_URL}/api/tracker/"
        f"trackedEntities/{tei_uid}"
    )

    params = {
        "fields": (
            "trackedEntity,"
            "enrollments[enrollment,program,orgUnit,status]"
        )
    }

    response = session.get(
        url,
        params=params,
        timeout=60
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed to retrieve TEI {tei_uid}: "
            f"HTTP {response.status_code}: "
            f"{response.text}"
        )

    result = response.json()

    enrollments = result.get("enrollments", [])

    matching_enrollments = [
        enrollment
        for enrollment in enrollments
        if enrollment.get("program") == program_uid
    ]

    if not matching_enrollments:
        raise ValueError(
            f"No enrollment found for TEI={tei_uid}, "
            f"program={program_uid}"
        )

    if len(matching_enrollments) > 1:
        raise ValueError(
            f"Multiple enrollments found for "
            f"TEI={tei_uid}, program={program_uid}. "
            "Please specify the correct enrollment UID."
        )

    enrollment_uid = matching_enrollments[0].get(
        "enrollment"
    )

    if not enrollment_uid:
        raise ValueError(
            f"Enrollment UID is missing for TEI={tei_uid}"
        )

    return enrollment_uid



# ============================================================
# VALUE HELPERS
# ============================================================

def is_empty(value):
    if value is None:
        return True

    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def clean_value(value):
    """
    Convert Excel values into DHIS2-compatible strings.
    """

    if is_empty(value):
        return None

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, bool):
        return str(value).lower()

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))

        return str(value)

    return str(value).strip()


def format_event_date(value):
    """
    DHIS2 event date format: YYYY-MM-DD.
    """

    if is_empty(value):
        raise ValueError("eventDate is empty")

    parsed = pd.to_datetime(
        value,
        errors="raise"
    )

    return parsed.strftime("%Y-%m-%d")


# ============================================================
# READ EXCEL WITH ORIGINAL HEADERS
# ============================================================

def read_excel_file():

    # Read original Excel headers using openpyxl.
    # This preserves repeated data element UIDs.

    from openpyxl import load_workbook

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True
    )

    worksheet = workbook[SHEET_NAME]

    headers = [
        str(cell.value).strip()
        if cell.value is not None
        else ""
        for cell in worksheet[1]
    ]

    workbook.close()

    # Pandas reads data rows.
    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=1,
        dtype=object
    )

    if len(headers) != len(df.columns):
        raise ValueError(
            "Excel header count does not match data columns"
        )

    df.columns = headers

    missing = REQUIRED_COLUMNS - set(df.columns)

    if missing:
        raise ValueError(
            f"Missing required columns: {missing}"
        )

    return df


# ============================================================
# BUILD DATA VALUES
# ============================================================


'''
def build_data_values(row, headers):

    data_values = {}

    for column in headers:

        # Skip metadata and GPS columns.
        if column in METADATA_COLUMNS:
            continue

        data_element_uid = str(column).strip()

        if not data_element_uid:
            continue

        value = clean_value(row[column])

        # Skip empty Excel cells.
        if value is None or value == "":
            continue

        # Duplicate data element UIDs are combined.
        # If duplicate columns contain different values,
        # stop rather than silently overwrite data.
        if data_element_uid in data_values:

            existing_value = data_values[
                data_element_uid
            ]

            if existing_value != value:
                raise ValueError(
                    f"Conflicting duplicate data element "
                    f"{data_element_uid}: "
                    f"{existing_value} vs {value}"
                )

            continue

        data_values[data_element_uid] = value

    return [
        {
            "dataElement": uid,
            "value": value
        }
        for uid, value in data_values.items()
    ]

'''


def build_data_values(row, headers):

    data_values = {}

    for position, column in enumerate(headers):

        if column in METADATA_COLUMNS:
            continue

        data_element_uid = str(column).strip()

        if not data_element_uid:
            continue

        value = row.iloc[position]

        if is_empty(value):
            continue

        # Date-type data element oHIf2bbxU7Z, RrrUkNy8mJv zt0vGb2dIRr, jh08Ocg1Lab, L5mWX8sQnVR, WVAOQENG6Um, yRU8RVm4TZY
        if data_element_uid == "oHIf2bbxU7Z":

            value = format_data_element_date(value)

        else:
            value = clean_value(value)

        if value is None or value == "":
            continue

        if data_element_uid in data_values:

            existing_value = data_values[
                data_element_uid
            ]

            if existing_value != value:
                raise ValueError(
                    f"Conflicting duplicate data element "
                    f"{data_element_uid}: "
                    f"{existing_value} vs {value}"
                )

            continue

        data_values[data_element_uid] = value

    return [
        {
            "dataElement": uid,
            "value": value
        }
        for uid, value in data_values.items()
    ]

# ============================================================
# BUILD EVENT PAYLOAD
# ============================================================

'''
def build_event_payload(row, headers):

    tei_uid = clean_value(row["tei_uid"])
    orgunit_uid = clean_value(row["orgunit_uid"])
    program_uid = clean_value(row["program_uid"])
    program_stage_uid = clean_value(
        row["program_stage_uid"]
    )

    if not tei_uid:
        raise ValueError("tei_uid is empty")

    if not orgunit_uid:
        raise ValueError("orgunit_uid is empty")

    if not program_uid:
        raise ValueError("program_uid is empty")

    if not program_stage_uid:
        raise ValueError("program_stage_uid is empty")

    event_date = format_event_date(
        row["eventDate"]
    )

    data_values = build_data_values(
        row,
        headers
    )

    if not data_values:
        raise ValueError(
            "No non-empty data element values found"
        )

    event = {
        "trackedEntity": tei_uid,
        "program": program_uid,
        "programStage": program_stage_uid,
        "orgUnit": orgunit_uid,
        "occurredAt": event_date,
        "status": "COMPLETED",
        "dataValues": data_values
    }

    return {
        "events": [event]
    }
'''


def build_event_payload(row, headers):

    event_uid = clean_value(row["event_uid"])
    tei_uid = clean_value(row["tei_uid"])
    orgunit_uid = clean_value(row["orgunit_uid"])
    program_uid = clean_value(row["program_uid"])
    enrollment_uid = get_enrollment_uid(
        tei_uid,
        program_uid
    )
    program_stage_uid = clean_value(
        row["program_stage_uid"]
    )

    # --------------------------------------------------------
    # Validate required fields
    # --------------------------------------------------------

    required_values = {
        "event_uid": event_uid,
        "tei_uid": tei_uid,
        "orgunit_uid": orgunit_uid,
        "program_uid": program_uid,
        "program_stage_uid": program_stage_uid
    }

    for field, value in required_values.items():

        if not value:
            raise ValueError(
                f"{field} is empty"
            )

    # DHIS2 UID: 11 alphanumeric characters
    '''
    import re
    if not re.fullmatch(
        r"[A-Za-z][A-Za-z0-9]{10}",
        event_uid
    ):
        raise ValueError(
            f"Invalid event UID: {event_uid}"
        )
    '''

    # --------------------------------------------------------
    # Event date
    # --------------------------------------------------------

    event_date = format_event_date(
        row["eventDate"]
    )

    # --------------------------------------------------------
    # Event data values
    # --------------------------------------------------------

    data_values = build_data_values(
        row,
        headers
    )

    if not data_values:
        raise ValueError(
            "No non-empty data element values found"
        )

    # --------------------------------------------------------
    # Event coordinates
    # --------------------------------------------------------

    latitude = clean_value(
        row["GPS Latitude"]
    )

    longitude = clean_value(
        row["GPS Longitude"]
    )

    geometry = None

    # Both coordinates must be present.
    if latitude is not None and longitude is not None:

        try:
            latitude = float(latitude)
            longitude = float(longitude)

        except (ValueError, TypeError):
            raise ValueError(
                "Invalid GPS Latitude or GPS Longitude"
            )

        if not -90 <= latitude <= 90:
            raise ValueError(
                f"Invalid latitude: {latitude}"
            )

        if not -180 <= longitude <= 180:
            raise ValueError(
                f"Invalid longitude: {longitude}"
            )

        # GeoJSON uses longitude first, latitude second.
        geometry = {
            "type": "Point",
            "coordinates": [
                longitude,
                latitude
            ]
        }

    elif latitude is not None or longitude is not None:

        raise ValueError(
            "Both GPS Latitude and GPS Longitude "
            "must be provided together"
        )

    # --------------------------------------------------------
    # Build Tracker event
    # --------------------------------------------------------

    event = {
        "event": event_uid,
        "trackedEntity": tei_uid,
        "enrollment": enrollment_uid,
        "program": program_uid,
        "programStage": program_stage_uid,
        "orgUnit": orgunit_uid,
        "occurredAt": event_date,
        "status": "ACTIVE",
        #"featureType": "POINT",
        "featureType": "NO",
        "dataValues": data_values
    }

    # Add geometry only when coordinates are available.
    if geometry is not None:
        event["geometry"] = geometry

    return {
        "events": [event]
    }


# ============================================================
# IMPORT EVENT INTO DHIS2
# ============================================================

def import_event(payload):

    response = session.post(
        TRACKER_URL,
        params={
            "async": "false",
            "importStrategy": "CREATE",
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
# EXTRACT DHIS2 IMPORT RESULT
# ============================================================

def extract_result(result):

    stats = result.get("stats", {})

    if not stats:
        stats = result.get(
            "bundleReport", {}
        ).get("stats", {})

    return {
        "dhis2_status": result.get(
            "status", "UNKNOWN"
        ),
        "created": stats.get("created", 0),
        "updated": stats.get("updated", 0),
        "ignored": stats.get("ignored", 0),
        "deleted": stats.get("deleted", 0),
        "total": stats.get("total", 0)
    }


# ============================================================
# MAIN
# ============================================================

def main():

    logging.info("Starting DHIS2 event import")

    df = read_excel_file()

    headers = list(df.columns)

    df = df.dropna(subset=["tei_uid"])

    logging.info(
        "Total events to process: %s",
        len(df)
    )

    report_rows = []

    for index, row in df.iterrows():

        tei_uid = clean_value(row["tei_uid"])

        report = {
            "excel_row": index + 2,
            "tei_uid": tei_uid,
            "program_uid": clean_value(
                row["program_uid"]
            ),
            "program_stage_uid": clean_value(
                row["program_stage_uid"]
            ),
            "eventDate": clean_value(
                row["eventDate"]
            ),
            "status": "FAILED",
            "message": ""
        }

        try:

            payload = build_event_payload(
                row,
                headers
            )

            if DRY_RUN:

                report["status"] = "DRY_RUN"
                report["message"] = str(payload)

                logging.info(
                    "DRY RUN | TEI=%s",
                    tei_uid
                )

            else:

                result = import_event(payload)

                summary = extract_result(result)

                report.update(summary)

                report["message"] = str(result)

                if summary["dhis2_status"] == "OK":
                    report["status"] = "SUCCESS"

                    logging.info(
                        "Event imported | TEI=%s",
                        tei_uid
                    )

                else:
                    report["status"] = "FAILED"

                    logging.error(
                        "Import failed | TEI=%s | %s",
                        tei_uid,
                        result
                    )

        except Exception as exc:

            report["status"] = "FAILED"
            report["message"] = str(exc)

            logging.exception(
                "Error importing event for TEI=%s",
                tei_uid
            )

        report_rows.append(report)

    # Save results.
    report_df = pd.DataFrame(report_rows)

    report_df.to_excel(
        REPORT_FILE,
        index=False
    )

    logging.info("Import completed")

    logging.info(
        "Successful: %s",
        (report_df["status"] == "SUCCESS").sum()
    )

    logging.info(
        "Failed: %s",
        (report_df["status"] == "FAILED").sum()
    )

    logging.info(
        "Report saved: %s",
        REPORT_FILE
    )


if __name__ == "__main__":
    main()