"""
Includes functions for processing overtime data, such as checking pay conditions,
counting overtime instances, and preparing data for official reports.
"""
import os
import logging
from typing import Dict
import datetime

import openpyxl
from openpyxl.styles import Alignment

from utils.excel_utils import apply_default_report_styles

logging.basicConfig(level=logging.WARNING)


MEAL_FEE = int(os.getenv("MEAL_FEE", 5500))

OFFICIAL_DATA_NAMES_STR = os.getenv("OFFICIAL_DATA_NAMES_STR", "")
OFFICIAL_DATA_NAMES = [
    name.strip() for name in OFFICIAL_DATA_NAMES_STR.split(',') if name.strip()]


def check_overtime_pay(file_path: str) -> None:
    """
    '매식비'라는 헤더의 컬럼에서 값이 'X'인 행을 삭제한다.
    기타 기존 조건도 유지.
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb[wb.sheetnames[0]]

    # 1. 헤더(1행)에서 '매식비' 컬럼 인덱스 찾기
    header_row = ws[1]
    meal_col_idx = None

    for idx, cell in enumerate(header_row, 1):  # 1-based index
        if cell.value == '매식비':
            meal_col_idx = idx
            break

    if meal_col_idx is None:
        print("❌ '매식비'라는 헤더가 없습니다.")
        return

    row_len = ws.max_row

    # 2. 아래에서 위로 데이터 행 반복
    for i in range(row_len, 1, -1):  # 2행부터 시작, 1행(헤더)는 제외
        if ws.cell(row=i, column=meal_col_idx).value == 'X':
            ws.delete_rows(i, 1)
            continue  # 삭제 시, 다음 라인으로

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
        value = row_data[7].value
        if isinstance(value, datetime.datetime):
            if value.strftime("%Y-%m-%d") in dateCnt:
                dateCnt[value.strftime("%Y-%m-%d")] += 1
            else:
                dateCnt[value.strftime("%Y-%m-%d")] = 1
        else:
            # 날짜가 None이거나 str도 아니면 경고 로그 출력
            logging.warning(
                f"[overtimeCnt] 잘못된 날짜 데이터: {value} (타입: {type(value)}) "
                f"엑셀 파일: {filename}"
            )

        value = row_data[5].value
        if isinstance(value, str):
            if value in overtimeNameCnt:
                overtimeNameCnt[value] += 1
            else:
                overtimeNameCnt[value] = 1
        else:
            # 이름이 None이거나 str이 아니면 경고 로그 출력
            logging.warning(
                f"[overtimeCnt] 잘못된 이름 데이터: {value} (타입: {type(value)}) "
                f"엑셀 파일: {filename}"
            )

        maxCnt += 1

    dateCnt = sorted(dateCnt.items())

    # ws2 = wb.create_sheet("매식비 통계", 0)

    # # 데이터 채우기
    # ws2['B1'] = "초과근무일자"
    # ws2['C1'] = '인원'
    # ws2['D1'] = '단가'
    # ws2['E1'] = '금액'
    # ws2['F1'] = '비고'

    # for i in range(0, len(dateCnt)):
    #     j = str(i+2)
    #     (date, cnt) = dateCnt[i]
    #     ws2['B'+j] = date
    #     ws2['C'+j] = cnt
    #     ws2['D'+j] = MEAL_FEE
    #     ws2['E'+j] = cnt*MEAL_FEE

    # lastRow = str(len(dateCnt)+3)
    # ws2['B'+lastRow] = '합계'
    # ws2['C'+lastRow] = maxCnt
    # ws2['D'+lastRow] = ''
    # ws2['E'+lastRow] = maxCnt*MEAL_FEE

    # apply_default_report_styles(ws2, center_columns=['인원'],
    #                             number_columns=['금액', '합계'],
    #                             column_style_map={
    #     '합계': {'align': Alignment(horizontal="center", vertical="center"), 'format': '#,##0'}
    # })  # Call the new styling function

    # wb.save(filename)

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

    print("\n%s | %s | %s | %s" %
          ("성명", "초과", "출근", "매식비"))

    for name in OFFICIAL_DATA_NAMES:
        try:
            overtimeNameCnt[name]
        except KeyError:
            overtimeNameCnt[name] = 0
        print("%s | %2d | %2d | %2d" %
              (name, data[name][0], data[name][1], overtimeNameCnt[name]))

    print("\n")
