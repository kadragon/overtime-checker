"""
Main entry point for processing overtime files.
"""
from utils.file_utils import find_target_excel_file, convert_xls_to_xlsx, create_meal_expense_file
from utils.overtime_utils import check_overtime_pay, overtimeCnt, officialDataMaker

from dotenv import load_dotenv


def main():
    """
    Directly processes the overtime files.
    """
    load_dotenv()

    print("초과근무승인 파일 처리를 시작합니다.")
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
    print('파일 찾기 완료')

    # 엑셀 파일 변환
    convertedExcelFileName = convert_xls_to_xlsx(overtimeListFileName)
    print('초과근무 승인 내역 엑셀 변환 완료')

    # 초과근무 검토 파일 생성
    overtimeFileName = create_meal_expense_file(convertedExcelFileName)
    print('초과근무 검토 완료')

    # 초과근무 조교 / 사전보고 등 검토
    check_overtime_pay(overtimeFileName)
    print('초과근무 조교 / 사전보고 등 검토 완료')

    # 초과근무 월집계 파일 생성
    overtimeNameCnt = overtimeCnt(overtimeFileName)
    print('초과근무 월집계 완료')

    # 월 집계 파일 찾기
    overtimeListFileName = find_target_excel_file("초과근무월집계")
    print("월 집계 파일 찾기 완료")

    # 엑셀 파일 변환
    convertedExcelFileName = convert_xls_to_xlsx(overtimeListFileName)
    print("엑셀 파일 변환 완료")

    # 공문 데이터 생성
    officialDataMaker(convertedExcelFileName, overtimeNameCnt)
    print("공문 데이터 생성 완료")


if __name__ == "__main__":
    main()
