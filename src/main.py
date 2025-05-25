"""
Main entry point for processing overtime files.
"""
from processors.overtime_analytic import create_overtime_file


def main():
    """
    Directly processes the overtime files.
    """
    print("초과근무승인 파일 처리를 시작합니다.")
    create_overtime_file()


if __name__ == "__main__":
    main()
