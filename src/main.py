from utils.file_utils import find_target_excel_file, convert_xls_to_xlsx, create_meal_expense_file
from utils.overtime_utils import check_overtime_pay, overtimeCnt, officialDataMaker
from dotenv import load_dotenv


def find_and_convert_excel(keyword):
    original = find_target_excel_file(keyword)
    print(f"{keyword} 파일 찾기 완료")

    converted = convert_xls_to_xlsx(original)
    print(f"{keyword} 엑셀 변환 완료")

    return converted


def process_overtime_approval(overtime_xlsx):
    review_file = create_meal_expense_file(overtime_xlsx)
    print("초과근무 검토 완료")

    check_overtime_pay(review_file)
    print("초과근무 조교 / 사전보고 등 검토 완료")

    count_data = overtimeCnt(review_file)
    print("초과근무 월집계 파일 생성 완료")

    return count_data


def generate_official_data(monthly_xlsx, count_data):
    officialDataMaker(monthly_xlsx, count_data)
    print("공문 데이터 생성 완료")


def main():
    load_dotenv()
    print("초과근무승인 파일 처리를 시작합니다.")

    # 1. 초과근무승인 파일 찾기 → 변환
    approval_xlsx = find_and_convert_excel("초과근무승인")

    # 2. 초과근무 승인 파일 처리 (검토 → 월집계)
    count_data = process_overtime_approval(approval_xlsx)

    # 3. 초과근무월집계 파일 찾기 → 변환
    monthly_xlsx = find_and_convert_excel("초과근무월집계")

    # 4. 공문 데이터 생성
    generate_official_data(monthly_xlsx, count_data)


if __name__ == "__main__":
    main()
