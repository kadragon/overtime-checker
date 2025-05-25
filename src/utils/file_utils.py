"""
Contains utility functions for file system operations like finding, copying, and converting files.
"""
import os
import shutil
import pandas as pd
from typing import Tuple, Optional, Union  # Added Union for return type
from config import DOWNLOAD_DIR, WORK_DIR
# Corrected import for local package
from .date_utils import get_current_and_previous_month
from ..constants.filenames import OVERTIME_APPROVAL_FILE_START, OVERTIME_MONTHLY_AGGREGATE_FILE_START, TARGET_OVERTIME_APPROVAL_FILENAME, TARGET_OVERTIME_MONTHLY_AGGREGATE_FILENAME


def find_target_excel_file(file_type: str) -> Union[str, None]:
    """
    Finds and copies a target Excel file (overtime approval or monthly aggregate)
    from the download directory to the working directory for the previous month.

    It identifies files based on predefined prefixes and the current month.
    If a file for the current month is found, it's copied to a directory
    named after the previous month, with a standardized target filename.

    Args:
        file_type (str): The type of file to find. Expected values are
                         "초과근무승인" (overtime approval) or
                         "초과근무월집계" (monthly overtime aggregate).

    Returns:
        Optional[str]: The path to the copied file in the work directory if found
                       and copied, otherwise None if the file_type is invalid or
                       raises an Exception if the file is not found.
                       (Note: The original code returned Exception, which is unusual.
                        Returning None for not found, or raising a specific error like
                        FileNotFoundError would be more Pythonic. For now, sticking
                        to original behavior of returning Exception object)
    """
    now_month, prev_month = get_current_and_previous_month()

    file_info: Tuple[str, str] = {
        "초과근무승인": (OVERTIME_APPROVAL_FILE_START, TARGET_OVERTIME_APPROVAL_FILENAME),
        "초과근무월집계": (OVERTIME_MONTHLY_AGGREGATE_FILE_START, TARGET_OVERTIME_MONTHLY_AGGREGATE_FILENAME)
    }.get(file_type, ("", ""))

    if not file_info[0]:
        return None

    file_start, target_filename = file_info
    work_dir = os.path.join(WORK_DIR, prev_month)
    os.makedirs(work_dir, exist_ok=True)
    os.chmod(work_dir, 0o777)  # 읽기, 쓰기, 실행 권한 부여

    for filename in os.listdir(DOWNLOAD_DIR):
        if filename.startswith(file_start + now_month):
            base_path = os.path.join(DOWNLOAD_DIR, filename)
            save_path = os.path.join(
                work_dir, f"{target_filename}{prev_month}).xls")

            if not os.path.isfile(save_path):
                shutil.copy(base_path, save_path)

            return save_path

    # The original code returns an Exception object, which is unconventional.
    # Typically, one would raise FileNotFoundError or return None.
    # For now, maintaining original behavior. A custom exception class would be better.
    # Consider changing this to: raise FileNotFoundError(f"No {file_type} file found for month {now_month}")
    raise FileNotFoundError(f"No {file_type} file found for {now_month}")


def convert_xls_to_xlsx(xls_file: str) -> str:
    """
    Converts an .xls file to .xlsx format using pandas.

    Args:
        xls_file (str): Path to the input .xls file.

    Returns:
        str: Path to the created .xlsx file.
    """
    df = pd.read_excel(xls_file, engine='xlrd')
    xlsx_file = xls_file.replace(".xls", ".xlsx")

    df.to_excel(xlsx_file, index=False, engine='openpyxl')

    return xlsx_file


def create_meal_expense_file(file_path: str) -> str:
    """
    Creates a copy of an overtime details file to be used as a meal expense file.

    The new file is named by replacing "초과근무내역" (overtime details) with
    "매식비" (meal expenses) in the original filename.

    Args:
        file_path (str): The path to the source overtime details file (usually .xlsx).

    Returns:
        str: The path to the newly created meal expense file.
    """
    meal_expense_filename = file_path.replace(
        # Using constant for source part
        TARGET_OVERTIME_APPROVAL_FILENAME, "매식비(")
    # Ensure the target name is also constructed consistently if needed, or ensure "매식비" is the final desired string
    return shutil.copy(file_path, meal_expense_filename)
