# constants.py

from datetime import datetime

#LOG_FILE = datetime.now().strftime("%Y-%m-%d") + "_dss_child_health_integration_missing_patient.log"


LOG_FILE_EVENT_PROGRAM = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + "_dhis2_event_program_import.log"

LOG_FILE_AGGREGATE_DATA_VALUE = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + "_dhis2_aggregated_datavalue_import.log"

LOG_FILE_TEI_ATTRIBUTE_VALUE_ERROR_LOG = datetime.now().strftime("%Y-%m-%d") + "_tei_update_error_log.txt"

LOG_FILE_ADD_UPDATE_BANK = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + "_oracle_netsuite_add_update_bank.log"