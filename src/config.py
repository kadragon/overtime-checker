from dotenv import load_dotenv
import os

load_dotenv()

DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR")
WORK_DIR = os.getenv("WORK_DIR")
MEAL_FEE = int(os.getenv("MEAL_FEE"))
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE")
