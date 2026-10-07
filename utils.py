# utils.py

import requests
from requests_oauthlib import OAuth1
import logging

import certifi  ## for post data in hmis production certificate issue


import json
import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders
from urllib.parse import quote

## for nepali date
#import nepali_datetime
from datetime import datetime, timedelta, date

#from datetime import timedelta

from dotenv import load_dotenv
import os
import glob
import base64
from dotenv import load_dotenv


load_dotenv()  # this loads .env file



from constants import LOG_FILE_TEI_ATTRIBUTE_VALUE_ERROR_LOG

#event program and its stage list
#https://stage1.pmnpis.org.ph/pmnpis-uat/api/programs/XboqpHaNn91.json?fields=id,name,programStages[id,name,sortOrder]
#https://lllmis.org/dhis/api/programs/Hm1Cwpei6i1.json?fields=id,name,programStages[id,name,sortOrder]


def configure_logging( log_file_name ):

    #Optional (Advanced, but useful)


    LOG_DIR = "logs"
    #os.makedirs(LOG_DIR, exist_ok=True)

    os.makedirs(LOG_DIR, exist_ok=True)
    assert LOG_DIR != "/" and LOG_DIR != "" #### Never delete outside log folder.

    # Create unique log filename
    #log_filename = f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
    log_filename = log_file_name
    #log_filename = f"{LOG_FILE}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
    log_path = os.path.join(LOG_DIR, log_filename)

    #logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    
    logging.basicConfig(filename=log_path, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def log_info(message):
    logging.info(message)

def log_error(message):
    logging.error(message)




'''
def configure_logging(log_file_name):

    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(log_dir, log_file_name)

    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        force=True
    )

    logging.info("Logging started")
    logging.info("Log file: %s", log_path)

    return log_path    
'''
