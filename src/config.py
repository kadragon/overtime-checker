"""
Handles loading configuration settings from environment variables.

This module uses the python-dotenv library to load variables from a .env file
and makes them available as Python constants.
"""
from dotenv import load_dotenv
import os

load_dotenv()

DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR")
WORK_DIR = os.getenv("WORK_DIR")
MEAL_FEE = int(os.getenv("MEAL_FEE"))
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE")

EXCLUDED_NAMES_CHECK_OVERTIME_STR = os.getenv("EXCLUDED_NAMES_CHECK_OVERTIME", "")
EXCLUDED_NAMES_CHECK_OVERTIME = [name.strip() for name in EXCLUDED_NAMES_CHECK_OVERTIME_STR.split(',') if name.strip()]

OFFICIAL_DATA_NAMES_STR = os.getenv("OFFICIAL_DATA_NAMES", "")
OFFICIAL_DATA_NAMES = [name.strip() for name in OFFICIAL_DATA_NAMES_STR.split(',') if name.strip()]
