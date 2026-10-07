
import os
import json
import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime, date

from dotenv import load_dotenv

from constants import LOG_FILE_AGGREGATE_DATA_VALUE
from utils import ( configure_logging )

# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

DHIS2_BASE_URL = os.getenv("DHIS2_BASE_URL", "").rstrip("/")
DHIS2_USERNAME = os.getenv("DHIS2_USERNAME")
DHIS2_PASSWORD = os.getenv("DHIS2_PASSWORD")
EXCEL_FILE_AGGREGATE_DATA_VALUE = os.getenv("EXCEL_FILE_AGGREGATE_DATA_VALUE")
#EXCEL_FILE_PATH = f"files/{EXCEL_FILE_EVENT_PROGRAM}"
EXCEL_FILE_PATH = os.path.join("files", EXCEL_FILE_AGGREGATE_DATA_VALUE)


# ============================================================
# DHIS2 CONFIGURATION
# ============================================================

DATA_VALUE_SET_ENDPOINT = (
    f"{DHIS2_BASE_URL}/api/dataValueSets"
)

# ============================================================
# LOGGING
# ============================================================

'''
logging.basicConfig(
    filename="dataValueSet_import.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
'''
# ============================================================
# DHIS2 SESSION
# ============================================================

session_get_post = requests.Session()

session_get_post.auth = HTTPBasicAuth(
    DHIS2_USERNAME,
    DHIS2_PASSWORD
)

session_get_post.headers.update({
    "Content-Type": "application/json",
    "Accept": "application/json"
})


# ============================================================
# PUSH DATAVALUESET TO DHIS2
# ============================================================

def push_dataValueSet_in_dhis2(dataValueSet_payload):

    try:

        response = session_get_post.post(
            DATA_VALUE_SET_ENDPOINT,
            json=dataValueSet_payload,
            timeout=120
        )

        if response.status_code == 200:

            response_json = response.json()

            response_data = response_json.get("response", {})

            conflicts = response_data.get("conflicts", [])

            description = response_data.get(
                "description",
                ""
            )

            import_count = response_data.get(
                "importCount",
                {}
            )

            imported = import_count.get(
                "imported",
                0
            )

            updated = import_count.get(
                "updated",
                0
            )

            ignored = import_count.get(
                "ignored",
                0
            )

            deleted = import_count.get(
                "deleted",
                0
            )

            print(
                f"DataValueSet imported successfully. "
                f"Imported: {imported}, "
                f"Updated: {updated}, "
                f"Ignored: {ignored}, "
                f"Deleted: {deleted}"
            )

            logging.info(
                f"DataValueSet imported successfully. "
                f"Imported: {imported}, "
                f"Updated: {updated}, "
                f"Ignored: {ignored}, "
                f"Deleted: {deleted}, "
                f"Description: {description}"
            )

            if conflicts:

                print(
                    f"Conflicts found: {len(conflicts)}"
                )

                logging.error(
                    f"DHIS2 conflicts: "
                    f"{json.dumps(conflicts)}"
                )

            return True

        else:

            print(
                f"Failed to import DataValueSet. "
                f"HTTP Status: {response.status_code}"
            )

            print(
                response.text
            )

            logging.error(
                f"Failed to import DataValueSet. "
                f"HTTP Status: {response.status_code}. "
                f"Response: {response.text}"
            )

            return False

    except Exception as e:

        print(
            f"Exception while importing DataValueSet: {e}"
        )

        logging.exception(
            "Exception while importing DataValueSet"
        )

        return False


# ============================================================
# READ EXCEL AND CREATE DHIS2 DATAVALUESET PAYLOAD
# ============================================================

def create_dataValueSet_payload(excel_file):

    print(
        f"Reading Excel file: {excel_file}"
    )

    logging.info(
        f"Reading Excel file: {excel_file}"
    )

    dataValueSet = pd.read_excel(
        excel_file
    )

    print(
        f"Total Excel rows: {len(dataValueSet)}"
    )

    logging.info(
        f"Total Excel rows: {len(dataValueSet)}"
    )


    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

    required_columns = [
        "orgunit_uid",
        "dataset_uid",
        "coc_uid",
        "attribute_uid",
        "period"
    ]


    for column in required_columns:

        if column not in dataValueSet.columns:

            raise Exception(
                f"Required column not found in Excel: {column}"
            )


    # --------------------------------------------------------
    # All columns after period are Data Element UIDs
    # --------------------------------------------------------

    period_index = dataValueSet.columns.get_loc(
        "period"
    )

    data_element_columns = list(
        dataValueSet.columns[
            period_index + 1:
        ]
    )


    print(
        f"Total Data Elements: "
        f"{len(data_element_columns)}"
    )

    logging.info(
        f"Total Data Elements: "
        f"{len(data_element_columns)}"
    )


    #tempDataValues = []

    tempDataValues = list()

    skipped_values = 0


    # ========================================================
    # PROCESS EACH EXCEL ROW
    # ========================================================

    for row_index, row in dataValueSet.iterrows():

        org_unit = str(
            row["orgunit_uid"]
        ).strip()

        data_set = str(
            row["dataset_uid"]
        ).strip()

        category_option_combo = str(
            row["coc_uid"]
        ).strip()

        attribute_option_combo = str(
            row["attribute_uid"]
        ).strip()

        period = str(
            row["period"]
        ).strip()


        # ----------------------------------------------------
        # Process every Data Element column
        # ----------------------------------------------------

        for data_element_uid in data_element_columns:

            cell_value = row[
                data_element_uid
            ]


            # ----------------------------------------------
            # Skip empty Excel cells
            # ----------------------------------------------

            if pd.isna(cell_value):

                skipped_values += 1

                continue


            # ----------------------------------------------
            # Convert value to string
            # ----------------------------------------------

            value = str(
                cell_value
            ).strip()


            # ----------------------------------------------
            # Remove commas from numeric values
            #
            # Example:
            # 2,592,073.72
            #
            # becomes:
            # 2592073.72
            # ----------------------------------------------

            value = value.replace(
                ",",
                ""
            )


            # ----------------------------------------------
            # Skip blank values
            # ----------------------------------------------

            if value == "":

                skipped_values += 1

                continue


            # ----------------------------------------------
            # Create DHIS2 DataValue
            # ----------------------------------------------

            dataValue = {

                "dataElement": str(
                    data_element_uid
                ).strip(),
            
                "categoryOptionCombo":
                    category_option_combo,
                
                "attributeOptionCombo":
                    attribute_option_combo,

                "orgUnit":
                    org_unit,

                "period":
                    period,

                "value":
                    value
            }


            tempDataValues.append(
                dataValue
            )


    # ========================================================
    # CREATE FINAL PAYLOAD
    # ========================================================

    dataValueSet_payload = {

        "dataSet": data_set,
        "dataValues": tempDataValues
    }


    print(
        f"Total DataValues generated: "
        f"{len(tempDataValues)}"
    )
    
    #print(f"DataValueSet Payload: "f"{dataValueSet_payload}")

    print(
        f"Empty cells skipped: "
        f"{skipped_values}"
    )


    logging.info(
        f"Total DataValues generated: "
        f"{len(tempDataValues)}"
    )

    logging.info(
        f"Empty cells skipped: "
        f"{skipped_values}"
    )


    return dataValueSet_payload


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    configure_logging( LOG_FILE_AGGREGATE_DATA_VALUE )

    #current_time_start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    current_time_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print( f"Pushing Aggregated DataValue Start . { current_time_start }" )
    logging.info(f"Pushing Aggregated DataValue Start . { current_time_start }")

    print("-" * 100)
    logging.info("-" * 100)

    print( f"EXCEL_FILE_PATH_NAME . { EXCEL_FILE_PATH }" )
    logging.info(f"EXCEL_FILE_PATH_NAME . { EXCEL_FILE_PATH }")
    
    try:
        
        # ----------------------------------------------------
        # Create payload from Excel
        # ----------------------------------------------------
        print("-" * 100)
        logging.info("-" * 100)
        
        dataValueSet_payload = (
            create_dataValueSet_payload(
                EXCEL_FILE_PATH
            )
        )


        # ----------------------------------------------------
        # Optional: Save payload for checking
        # ----------------------------------------------------
        '''
        with open(
            "dataValueSet_payload.json",
            "w",
            encoding="utf-8"
        ) as json_file:

            json.dump(
                dataValueSet_payload,
                json_file,
                indent=4,
                ensure_ascii=False
            )


        print(
            "Payload saved to "
            "dataValueSet_payload.json"
        )
        '''

        # ----------------------------------------------------
        # Push to DHIS2
        # ----------------------------------------------------
        print("-" * 100)
        logging.info("-" * 100)   

        success = (
            push_dataValueSet_in_dhis2(
                dataValueSet_payload
            )
        )

        print("-" * 100)
        logging.info("-" * 100)

        if success:

            print(
                "DataValueSet import completed."
            )

            logging.info(f"DataValueSet import completed.")
        else:

            print(
                "DataValueSet import failed."
            )
            logging.info(f"DataValueSet import failed.")

    except Exception as e:

        print(
            f"Process failed: {e}"
        )

        logging.exception(
            "DataValueSet import process failed"
        )

    current_time_end = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print( f"Pushing Aggregated DataValue End . { current_time_end }", flush=True )
    logging.info(f"Pushing Aggregated DataValue End . { current_time_end }" )        