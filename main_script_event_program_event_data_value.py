
import os
import logging
import math
from datetime import datetime, date

import pandas as pd
import requests
from dotenv import load_dotenv

from constants import LOG_FILE_EVENT_PROGRAM
from utils import ( configure_logging )

# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

DHIS2_BASE_URL = os.getenv("DHIS2_BASE_URL", "").rstrip("/")
DHIS2_USERNAME = os.getenv("DHIS2_USERNAME")
DHIS2_PASSWORD = os.getenv("DHIS2_PASSWORD")
EXCEL_FILE_EVENT_PROGRAM = os.getenv("EXCEL_FILE_EVENT_PROGRAM")
#EXCEL_FILE_PATH = f"files/{EXCEL_FILE_EVENT_PROGRAM}"
EXCEL_FILE_PATH = os.path.join("files", EXCEL_FILE_EVENT_PROGRAM)

#LOG_FILE = "dhis2_event_program_import.log"

# Existing events will be updated using their event UID.
#TRACKER_URL = f"{DHIS2_BASE_URL}/api/tracker"

# IMPORT_STRATEGY = "CREATE" "UPDATE"
EVENT_IMPORT_URL = (
    f"{DHIS2_BASE_URL}/api/tracker"
    "?async=false&importStrategy=CREATE"
)

## for update
'''
EVENT_IMPORT_URL = (
    f"{DHIS2_BASE_URL}/api/tracker"
    "?async=false&importStrategy=UPDATE"
)
'''

'''
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
'''

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def is_empty(value):
    """Check whether an Excel value is empty."""
    if value is None:
        return True

    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def convert_value(value):
    """Convert Excel values to DHIS2-compatible strings."""

    if is_empty(value):
        return None

    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, bool):
        return str(value).lower()

    if isinstance(value, float):
        if math.isfinite(value) and value.is_integer():
            return str(int(value))

    return str(value).strip()


def convert_event_date(value):
    """Convert Excel event date to YYYY-MM-DD."""

    if is_empty(value):
        raise ValueError("Event date is empty")

    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")

    parsed = pd.to_datetime(value, errors="raise")

    return parsed.strftime("%Y-%m-%d")


def get_column(df, column_name):
    """Find a column ignoring case and surrounding spaces."""

    for column in df.columns:
        if str(column).strip().lower() == column_name.lower():
            return column

    raise ValueError(
        f"Required Excel column not found: {column_name}"
    )


# ============================================================
# BUILD EVENT FROM EXCEL ROW
# ============================================================

def build_event(row, event_uid, program_uid, program_stage_uid, orgunit_uid, eventDate):

    # Read metadata using Excel column names.
    event_uid = convert_value(row[event_uid])
    program_uid = convert_value(row[program_uid])
    program_stage_uid = convert_value(row[program_stage_uid])
    orgunit_uid = convert_value(row[orgunit_uid])

    if not event_uid:
        raise ValueError("Event UID is empty")

    if not program_uid:
        raise ValueError("Program UID is empty")

    if not program_stage_uid:
        raise ValueError("Program Stage UID is empty")
    
    if not orgunit_uid:
        raise ValueError("Organisation unit UID is empty")

    event_date = convert_event_date(row[eventDate])
    #https://lllmis.org/dhis/api/programs/Hm1Cwpei6i1.json?fields=id,name,programStages[id,name,sortOrder]
    event = {
        "event": event_uid,
        "program": program_uid,
        "programStage": program_stage_uid,
        "orgUnit": orgunit_uid,
        "occurredAt": event_date,
        #"featureType": "NONE",
        #"status": "ACTIVE",
        "status": "COMPLETED",
        "dataValues": []
    }

    # These are Excel column names, not their values.
    '''
    metadata_columns = {
        event_col,
        program_col,
        orgunit_col,
        date_col
    }
    '''
    METADATA_COLUMNS = {
        "event_uid",
        "orgunit_uid",
        "program_uid",
        "program_stage_uid",
        "eventDate"
    }



    # Remaining columns are data element UIDs.
    for column in row.index:

        if column in METADATA_COLUMNS:
            continue

        data_element_uid = str(column).strip()

        value = convert_value(row[column])

        if value is None or value == "":
            continue

        event["dataValues"].append({
            "dataElement": data_element_uid,
            "value": value
        })

    return event


# ============================================================
# IMPORT EVENT TO DHIS2
# ============================================================
'''
def import_event(session, event):

    payload = {
        "events": [event]
    }

    response = session.post(
        EVENT_IMPORT_URL,
        json=payload,
        timeout=120
    )

    if not response.ok:
        #raise RuntimeError(f"HTTP {response.status_code}: {response.text}")
        #print(f"Event {event['event']}")
        #print("Event %s event["event"] )
        print("Event %s" % event["event"])

    result = response.json()

    logging.info(
        "Event %s response: %s",
        event["event"],
        result
    )
    print(
        "Event %s response: %s",
        event["event"],
        result
    )

    # DHIS2 can return HTTP 200 even when the import
    # contains validation errors.

    status = result.get("status", "").upper()

    if status == "ERROR":
        raise RuntimeError(
            f"DHIS2 import error: {result}"
        )

    stats = result.get("stats", {})

    return result, stats

'''


def import_event(session, event):

    payload = {
        "events": [event]
    }

    try:
        response = session.post(
            EVENT_IMPORT_URL,
            json=payload,
            timeout=120
        )
        
        logging.info(
            "Event %s HTTP status: %s",
            event["event"],
            response.status_code
        )

        # Log HTTP errors before raising.
        if not response.ok:

            error_message = (
                f"HTTP {response.status_code}: "
                f"{response.text}"
            )

            '''
            logging.error(
                "Event %s failed: %s",
                event["event"],
                error_message
            )

            print(
                "Event %s failed: %s",
                event["event"],
                error_message
            )
            '''
            raise RuntimeError(error_message)

        result = response.json()

        '''
        logging.info(
            "Event %s response: %s",
            event["event"],
            result
        )
        '''

        '''
        print(
            "Event %s response: %s",
            event["event"],
            result
        )
        '''
        #print(f"Event:  {event['event']} response: {result}")

        #print("Event %s response: %s" % (event["event"], result))

        status = result.get("status", "").upper()

        if status == "ERROR":

            error_message = (
                f"DHIS2 import error: {result}"
            )

            logging.error(
                "Event %s failed: %s",
                event["event"],
                error_message
            )

            print(
                "Event %s failed: %s",
                event["event"],
                error_message
            )
            raise RuntimeError(error_message)

        return result, result.get("stats", {})

    except requests.RequestException:

        '''
        logging.exception(
            "HTTP request failed for event %s",
            event["event"]
        )
        '''
        logging.error(
            "HTTP request failed for event %s",
            event["event"]
        )

        print(
            "HTTP request failed for event %s",
            event["event"]
        )
        raise


# ============================================================
# MAIN
# ============================================================

def main():

    #log_path = configure_logging(LOG_FILE_EVENT_PROGRAM)
    #print(f"Log file location: {log_path}")

    configure_logging( LOG_FILE_EVENT_PROGRAM )

    #current_time_start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    current_time_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print( f"Import Event Program DataValue Start . { current_time_start }" )
    logging.info(f"Import Event Program DataValue Start . { current_time_start }")

    print( f"EXCEL_FILE_PATH_NAME . { EXCEL_FILE_PATH }" )
    logging.info(f"EXCEL_FILE_PATH_NAME . { EXCEL_FILE_PATH }")

    if not all([
        DHIS2_BASE_URL,
        DHIS2_USERNAME,
        DHIS2_PASSWORD
    ]):
        raise ValueError(
            "Please configure DHIS2_API_URL, "
            "DHIS2_USERNAME and DHIS2_PASSWORD in .env"
        )

    if not os.path.exists(EXCEL_FILE_PATH):
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE_PATH}"
        )

    df = pd.read_excel(
        EXCEL_FILE_PATH,
        dtype=object
    )


    print( f"Event list size . { len(df) }" )
    logging.info( f"Event list size . { len(df) }" )

    print("-" * 100)
    logging.info("-" * 100)

    # Normalize Excel column names.
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    event_uid = get_column(df, "event_uid")
    program_uid = get_column(df, "program_uid")
    program_stage_uid = get_column(df, "program_stage_uid")
    orgunit_uid = get_column(df, "orgunit_uid")
    eventDate = get_column(df, "eventDate")

    # Optional enrollment column.
    #enrollment_col = None
    '''
    for column in df.columns:
        if column.lower() == "enrollment":
            enrollment_col = column
            break
     '''
           
    session = requests.Session()

    session.auth = (
        DHIS2_USERNAME,
        DHIS2_PASSWORD
    )

    session.headers.update({
        "Content-Type": "application/json",
        "Accept": "application/json"
    })

    total = len(df)
    success = 0
    failed = 0
    
    for index, row in df.iterrows():

        current_event_uid = convert_value(row[event_uid])

        try:
            event = build_event(
                row,
                event_uid,
                program_uid,
                program_stage_uid,
                orgunit_uid,
                eventDate
            )

            result, stats = import_event(
                session,
                event
            )

            success += 1

            print(
                f"Row {index + 2}/{total + 1}: "
                f"SUCCESS - Event {current_event_uid}"
            )
            #flush=True tells Python to send the printed text to the output immediately,

            print(f"Event: {event['event']} response: {result}", flush=True)
            print(f"Event: {event['event']} stats: {stats}", flush=True)

            logging.info(
                f"Row {index + 2}/{total + 1}: "
                f"SUCCESS - Event {current_event_uid}"
            )
            logging.info(
                "Event: %s response: %s",
                event["event"],
                result
            )
            logging.info(
                "Event: %s response: %s",
                event["event"],
                stats
            )

            #logging.info(f"Import result: {stats}")
            #print(f"Import result: {stats}")


        except Exception as error:

            failed += 1
            '''
            logging.exception(
                "Failed importing event %s",
                current_event_uid
            )
            '''

            error_message = (
                f"Row {index + 2}/{total + 1}: "
                f"FAILED - Event {current_event_uid}: {error}"
            )

            logging.error(error_message)

            print(error_message)
            '''
            print(
                f"Row {index + 2}/{total + 1}: "
                f"FAILED - Event {current_event_uid}: {error}"
            )
            '''
        # ========================================================
        # SEPARATOR AFTER EVERY ROW
        # ========================================================

        print("-" * 100)
        logging.info("-" * 100)


    print("\n========== IMPORT SUMMARY ==========")
    logging.info("\n========== IMPORT SUMMARY ==========")

    print(f"Total events : {total}")
    logging.info(f"Total events : {total}")

    print(f"Successful   : {success}")
    logging.info(f"Successful   : {success}")

    print(f"Failed       : {failed}")
    logging.info(f"Failed       : {failed}")

    print(f"Log file     : {LOG_FILE_EVENT_PROGRAM}")
    logging.info(f"Log file     : {LOG_FILE_EVENT_PROGRAM}")


if __name__ == "__main__":

    main()

    current_time_end = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print( f"Import Event Program DataValue End . { current_time_end }", flush=True )
    logging.info(f"Import Event Program DataValue End . { current_time_end }" )