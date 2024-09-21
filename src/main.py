from utils.file_utils import find_target_excel_file, convert_xls_to_xlsx, create_meal_expense_file
from utils.overtime_utils import check_overtime_pay, overtimeCnt, officialDataMaker


def main():
    # 초과근무승인 파일 찾기
    overtimeListFileName = find_target_excel_file("초과근무승인")

    # 엑셀 파일 변환
    convertedExcelFileName = convert_xls_to_xlsx(overtimeListFileName)

    # 초과근무 검토 파일 생성
    overtimeFileName = create_meal_expense_file(convertedExcelFileName)

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


if __name__ == "__main__":
    main()
