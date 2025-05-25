"""
Handles loading configuration settings from environment variables.

This module uses the python-dotenv library to load variables from a .env file
and makes them available as Python constants.
"""
from dotenv import load_dotenv
import os
from typing import Optional


def require_env_var(name: str, default: Optional[str] = None) -> str:
    value = os.getenv(name, default)
    if value is None or value == "":
        raise EnvironmentError(f"환경변수 '{name}'가 설정되어 있지 않습니다.")
    return value


load_dotenv()

DOWNLOAD_DIR = require_env_var("DOWNLOAD_DIR")
WORK_DIR = require_env_var("WORK_DIR")

_meal_fee_str = require_env_var("MEAL_FEE")
try:
    MEAL_FEE = int(_meal_fee_str)
except ValueError:
    raise EnvironmentError(
        f"환경변수 'MEAL_FEE'는 정수여야 합니다. 제공된 값: '{_meal_fee_str}'")

SPREADSHEET_ID = require_env_var("SPREADSHEET_ID")
GOOGLE_CREDENTIALS_FILE = require_env_var("GOOGLE_CREDENTIALS_FILE")

EXCLUDED_NAMES_CHECK_OVERTIME_STR = os.getenv(
    "EXCLUDED_NAMES_CHECK_OVERTIME", "")
EXCLUDED_NAMES_CHECK_OVERTIME = [name.strip(
) for name in EXCLUDED_NAMES_CHECK_OVERTIME_STR.split(',') if name.strip()]

OFFICIAL_DATA_NAMES_STR = os.getenv("OFFICIAL_DATA_NAMES", "")
OFFICIAL_DATA_NAMES = [
    name.strip() for name in OFFICIAL_DATA_NAMES_STR.split(',') if name.strip()]
