import os
import shutil
from typing import Tuple, Optional
import win32com.client as win32
from config import DOWNLOAD_DIR, WORK_DIR
from utils.date_utils import get_current_and_previous_month


def find_target_excel_file(file_type: str) -> Optional[str]:
    """
    파일 목록에서 초과근무승인 또는 초과근무월집계 파일을 찾아 복사한다.
    """
    now_month, prev_month = get_current_and_previous_month()

    file_info: Tuple[str, str] = {
        "초과근무승인": ("초과근무승인(서무용)_", "초과근무내역("),
        "초과근무월집계": ("초과근무월집계_", "초과근무월집계(")
    }.get(file_type, ("", ""))

    if not file_info[0]:
        return None

    file_start, target_filename = file_info
    work_dir = os.path.join(WORK_DIR, prev_month)
    os.makedirs(work_dir, exist_ok=True)

    for filename in os.listdir(DOWNLOAD_DIR):
        if filename.startswith(file_start + now_month):
            base_path = os.path.join(DOWNLOAD_DIR, filename)
            save_path = os.path.join(
                work_dir, f"{target_filename}{prev_month}).xls")

            if not os.path.isfile(save_path):
                shutil.copy(base_path, save_path)

            return save_path

    return Exception(f"No {file_type} file found")


def convert_xls_to_xlsx(filename: str) -> str:
    try:
        xls = win32.gencache.EnsureDispatch('Excel.Application')
        wb = xls.Workbooks.Open(filename)

        xlsx_filename = filename + "x"
        wb.SaveAs(xlsx_filename, FileFormat=51)
        wb.Close()
        xls.Application.Quit()

        os.remove(filename)
        return xlsx_filename
    except Exception as e:
        print(f"Error converting {filename} to xlsx: {e}")
        return filename


def create_meal_expense_file(file_path: str) -> str:
    """
    초과근무내역 파일을 복사하여 매식비 파일을 생성합니다.

    Args:
        file_path (str): 초과근무내역 파일 경로

    Returns:
        str: 생성된 매식비 파일 경로
    """
    meal_expense_filename = file_path.replace("초과근무내역", "매식비")
    return shutil.copy(file_path, meal_expense_filename)
