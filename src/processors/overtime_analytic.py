"""
Orchestrates the process of creating various overtime-related files.

This module handles the workflow from finding raw overtime approval files
to generating processed reports and data for official documentation.
"""
from utils.file_utils import find_target_excel_file, convert_xls_to_xlsx, create_meal_expense_file
from utils.overtime_utils import check_overtime_pay, overtimeCnt, officialDataMaker


def create_overtime_file():
    """
    Manages the workflow of finding, converting, and processing overtime files.

    This includes:
    1. Finding the '초과근무승인' (overtime approval) file.
    2. Converting it from .xls to .xlsx.
    3. Creating a copy for meal expense calculations.
    4. Checking and filtering overtime records based on specific criteria.
    5. Generating a meal expense summary ('매식비 통계').
    6. Finding the '초과근무월집계' (monthly overtime aggregate) file.
    7. Converting it to .xlsx.
    8. Generating data for official documentation using the aggregate file and counts.
    """
    # 초과근무승인 파일 찾기
    overtimeListFileName = find_target_excel_file("초과근무승인")
    # print('파일 찾기 완료')

    # 엑셀 파일 변환
    convertedExcelFileName = convert_xls_to_xlsx(overtimeListFileName)
    # print('초과근무 승인 내역 엑셀 변환 완료')

    # 초과근무 검토 파일 생성
    overtimeFileName = create_meal_expense_file(convertedExcelFileName)
    # print('초과근무 검토 완료')

    # 초과근무 조교 / 사전보고 등 검토
    check_overtime_pay(overtimeFileName)

    # 초과근무 월집계 파일 생성
    overtimeNameCnt = overtimeCnt(overtimeFileName)

    # 월 집계 파일 찾기
    overtimeListFileName = find_target_excel_file("초과근무월집계")

    # 엑셀 파일 변환
    convertedExcelFileName = convert_xls_to_xlsx(overtimeListFileName)

    # 공문 데이터 생성
    officialDataMaker(convertedExcelFileName, overtimeNameCnt)
