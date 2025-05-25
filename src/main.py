"""
Main entry point for the application.

Provides a command-line interface to trigger different processing tasks
such as creating weekly work plans or processing overtime files.
"""
from processors.weekpage_maker import create_weekly_sheet
from processors.overtime_maker import create_overtime_file


def main():
    """
    Handles user input to run different processing tasks.

    Presents a menu to the user and executes the selected functionality,
    either creating a weekly sheet or processing overtime files.
    """
    print('''
          [1] 주간업무계획 생성
          [2] 초과근무승인 파일 처리
          ''')
    print("번호를 입력하세요: ")
    num = input()

    if num == "1":
        create_weekly_sheet(1)
    elif num == "2":
        create_overtime_file()


if __name__ == "__main__":
    main()
