from processors.weekpage_maker import create_weekly_sheet
from processors.overtime_maker import create_overtime_file


def main():
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
