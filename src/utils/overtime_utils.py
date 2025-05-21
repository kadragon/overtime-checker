"""
Includes functions for processing overtime data, such as checking pay conditions,
counting overtime instances, and preparing data for official reports.
"""
import openpyxl
from typing import Dict # For type hinting

from openpyxl.worksheet.dimensions import ColumnDimension # Keep if still needed, remove if not
# from openpyxl.utils import get_column_letter # This might be in excel_utils now

from ..config import MEAL_FEE, EXCLUDED_NAMES_CHECK_OVERTIME, OFFICIAL_DATA_NAMES
from .excel_utils import apply_default_report_styles


def check_overtime_pay(file_path: str) -> None:
    """
    Filters rows in an overtime Excel sheet based on specific conditions
    and user input for pre-approved overtime.

    Conditions for row deletion include:
    - Column 'U' value is 'X'.
    - Name in column 'F' is in the EXCLUDED_NAMES_CHECK_OVERTIME list.
    - For specific days (Wednesday, Saturday, Sunday), prompts user for
      confirmation if overtime was pre-approved; deletes row if not.

    Args:
        file_path (str): Path to the overtime Excel file.
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb[wb.sheetnames[0]]

    row_len = len(ws['A'])

    for i in range(row_len-1, 0, -1):
        if ws['U'][i].value == 'X' or ws['F'][i].value in EXCLUDED_NAMES_CHECK_OVERTIME:
            ws.delete_rows(i+1, 1)
        else:
            if ws['I'][i].value in ['수요일', '토요일', '일요일']:
                if input(ws['F'][i].value + " | " + ws['H'][i].value.strftime("%Y-%m-%d") + " | " + ws['I'][i].value + " | 사전 보고 확인? :").upper() == 'N':
                    ws.delete_rows(i+1, 1)

    wb.save(file_path)


def overtimeCnt(filename: str) -> Dict[str, int]:
    """
    Processes an overtime file to count overtime instances per person and
    generate a summary sheet ('매식비 통계') with meal expenses.

    The summary sheet includes overtime dates, personnel count per date,
    unit meal fee, total meal expenses per date, and overall totals.
    It also applies default styling to the new sheet.

    Args:
        filename (str): Path to the overtime Excel file.

    Returns:
        Dict[str, int]: A dictionary mapping names to their overtime counts.
    """
    overtimeNameCnt: Dict[str, int] = {}

    wb = openpyxl.load_workbook(filename)
    ws = wb[wb.sheetnames[0]]

    dateCnt = {}
    maxCnt = 0

    for row_data in ws.iter_rows(2):
        if row_data[7].value.strftime("%Y-%m-%d") in dateCnt:
            dateCnt[row_data[7].value.strftime("%Y-%m-%d")] += 1
        else:
            dateCnt[row_data[7].value.strftime("%Y-%m-%d")] = 1

        if row_data[5].value in overtimeNameCnt:
            overtimeNameCnt[row_data[5].value] += 1
        else:
            overtimeNameCnt[row_data[5].value] = 1

        maxCnt += 1

    dateCnt = sorted(dateCnt.items())

    ws2 = wb.create_sheet("매식비 통계", 0)
    ColumnDimension(ws2, bestFit=True)

    # 데이터 채우기
    ws2['B2'] = "초과근무일자"
    ws2['C2'] = '인원'
    ws2['D2'] = '단가'
    ws2['E2'] = '금액'
    ws2['F2'] = '비고'

    for i in range(0, len(dateCnt)):
        j = str(i+3)
        (date, cnt) = dateCnt[i]
        ws2['B'+j] = date
        ws2['C'+j] = cnt
        ws2['D'+j] = MEAL_FEE
        ws2['E'+j] = cnt*MEAL_FEE

    lastRow = str(len(dateCnt)+3)
    ws2['B'+lastRow] = '합계'
    ws2['C'+lastRow] = maxCnt
    ws2['D'+lastRow] = ''
    ws2['E'+lastRow] = maxCnt*MEAL_FEE

    apply_default_report_styles(ws2) # Call the new styling function

    wb.save(filename)

    return overtimeNameCnt


def officialDataMaker(filename: str, overtimeNameCnt: Dict[str, int]) -> None:
    """
    Reads an overtime monthly aggregate file and combines it with overtime counts
    to print a summary for specific individuals (defined in OFFICIAL_DATA_NAMES).

    The summary includes name, a value from column 'K' (presumably hours),
    a value from column 'AC' (presumably another count or amount), and
    the overtime count from overtimeNameCnt.

    Args:
        filename (str): Path to the overtime monthly aggregate Excel file.
        overtimeNameCnt (Dict[str, int]): Dictionary mapping names to overtime counts.
    """
    wb = openpyxl.load_workbook(filename)
    ws = wb[wb.sheetnames[0]]

    data = {}

    row_len = len(ws['A'])
    for i in range(2, row_len-1):
        data[ws['I'][i].value] = [
            int(ws['K'][i].value.split(':')[0]), int(ws['AC'][i].value)]

    for name in OFFICIAL_DATA_NAMES:
        try:
            overtimeNameCnt[name]
        except KeyError:
            overtimeNameCnt[name] = 0

        print("%s | %2d | %2d | %2d" %
              (name, data[name][0], data[name][1], overtimeNameCnt[name]))
