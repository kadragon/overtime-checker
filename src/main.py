from utils.overtime_utils import check_overtime_pay, overtimeCnt, officialDataMaker
from utils.file_utils import (
    find_target_excel_file,
    convert_xls_to_xlsx,
    create_meal_expense_file,
)
from typing import List

import os


def find_and_convert_excel(keyword: str, download_dir: str, work_dir: str) -> str:
    original = find_target_excel_file(keyword, download_dir, work_dir)
    print(f"{keyword} 파일 찾기 완료")

    converted = convert_xls_to_xlsx(original)
    print(f"{keyword} 엑셀 변환 완료")

    return converted


def process_overtime_approval(overtime_xlsx: str) -> dict:
    review_file = create_meal_expense_file(overtime_xlsx)
    print("초과근무 검토 완료")

    check_overtime_pay(review_file)
    print("초과근무 조교 / 사전보고 등 검토 완료")

    count_data = overtimeCnt(review_file)
    print("초과근무 월집계 파일 생성 완료")

    return count_data


def generate_official_data(monthly_xlsx: str, count_data: dict, names: List[str]) -> None:
    officialDataMaker(monthly_xlsx, count_data, names)
    print("공문 데이터 생성 완료")


def main(
    download_dir: str,
    work_dir: str,
    meal_fee: str,
    official_data_names_str: str,
) -> None:
    print("초과근무승인 파일 처리를 시작합니다.")

    # 1. 초과근무승인 파일 찾기 → 변환
    approval_xlsx = find_and_convert_excel("초과근무승인", download_dir, work_dir)

    # 2. 초과근무 승인 파일 처리 (검토 → 월집계)
    count_data = process_overtime_approval(approval_xlsx)

    # 3. 초과근무월집계 파일 찾기 → 변환
    monthly_xlsx = find_and_convert_excel("초과근무월집계", download_dir, work_dir)

    # 4. 공문 데이터 생성
    names = [name.strip() for name in official_data_names_str.split(',') if name.strip()]
    generate_official_data(monthly_xlsx, count_data, names)


if __name__ == "__main__":
    raise RuntimeError(
        "This script should be run via an external interface that supplies arguments.")
