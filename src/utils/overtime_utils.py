"""
Includes functions for processing overtime data, such as checking pay conditions,
counting overtime instances, and preparing data for official reports.
"""
import logging
from typing import Dict, List
import datetime

import openpyxl
from openpyxl.styles import Alignment

from utils.excel_utils import apply_default_report_styles

logging.basicConfig(level=logging.INFO)




def check_overtime_pay(file_path: str) -> None:
    """
    '매식비'라는 헤더의 컬럼에서 값이 'X'인 행을 삭제한다.
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

    return overtimeNameCnt


def officialDataMaker(
    filename: str, overtimeNameCnt: Dict[str, int], official_data_names: List[str]
) -> None:
    """
    Reads an overtime monthly aggregate file and combines it with overtime counts
    to print a summary for specific individuals provided via ``official_data_names``.

    Args:
        filename (str): Path to the overtime monthly aggregate Excel file.
        overtimeNameCnt (Dict[str, int]): Dictionary mapping names to overtime counts.
        official_data_names (List[str]): Names to include in the summary.
    """
    wb = openpyxl.load_workbook(filename)
    ws = wb[wb.sheetnames[0]]

    data = {}

    # 헤더에서 필요한 컬럼 인덱스 찾기
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        header_value = ws.cell(row=1, column=col_idx).value
        if header_value in ['성명', '초과근무인정시간', '출근근무일수']:
            header_map[header_value] = col_idx

    if '성명' not in header_map:
        logging.warning("'성명' 헤더를 찾을 수 없습니다.")
        return
    if '초과근무인정시간' not in header_map:
        logging.warning("'초과근무인정시간' 헤더를 찾을 수 없습니다.")
        return
    if '출근근무일수' not in header_map:
        logging.warning("'출근근무일수' 헤더를 찾을 수 없습니다.")
        return

    name_col = header_map['성명']
    overtime_col = header_map['초과근무인정시간']
    attendance_col = header_map['출근근무일수']

    # 데이터는 행 3부터 시작 (행 1: 헤더, 행 2: 서브헤더)
    row_len = ws.max_row
    for i in range(3, row_len + 1):  # 합계 행 제외
        name = ws.cell(row=i, column=name_col).value
        if not name or name == "합계" or name == "총":
            continue

        # 초과근무인정시간
        overtime_value = ws.cell(row=i, column=overtime_col).value
        # 출근근무일수
        attendance_value = ws.cell(row=i, column=attendance_col).value
        
        # 초과근무시간 파싱
        overtime_hours = 0
        if overtime_value:
            try:
                # Handles "34" and "0034 : 01" formats
                overtime_str = str(overtime_value).split(':')[0].strip()
                if overtime_str:
                    overtime_hours = int(overtime_str)
            except ValueError:
                logging.warning(f"Could not parse overtime value '{overtime_value}' for name '{name}'.")

        # 출근근무일수 파싱
        attendance_days = 0
        if attendance_value:
            try:
                attendance_days = int(attendance_value)
            except (ValueError, TypeError):
                logging.warning(f"Could not parse attendance days value '{attendance_value}' for name '{name}'.")
        
        data[name] = [overtime_hours, attendance_days]

    print("\n%s | %s | %s | %s" %
          ("성명", "초과", "출근", "매식비"))

    for name in official_data_names:
        overtime_count = overtimeNameCnt.get(name, 0)
        overtime_hours, attendance_days = data.get(name, [0, 0])
        print("%s | %2d | %2d | %2d" %
              (name, overtime_hours, attendance_days, overtime_count))

    print("\n")
