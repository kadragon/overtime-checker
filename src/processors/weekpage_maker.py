import gspread
from contextlib import contextmanager
import datetime
from config import SPREADSHEET_ID, GOOGLE_CREDENTIALS_FILE


@contextmanager
def google_sheets():
    gc = gspread.service_account(GOOGLE_CREDENTIALS_FILE)
    yield gc


def create_weekly_sheet(len: int):
    with google_sheets() as gc:
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        copied_sheet = spreadsheet.get_worksheet(0).copy_to(SPREADSHEET_ID)
        copied_sheet = spreadsheet.worksheet(copied_sheet['title'])

        print("## 복사 완료...")

        # Update copied sheet title
        title = today_date()
        copied_sheet.update_title(title)

        # Set date
        copied_sheet.update_acell('A2', make_date(len))

        # index update
        copied_sheet.update_index(0)

        print("## 내용 정리 시작...")

        # Clear contents
        copied_sheet.batch_clear(['C4:C35'])
        copied_sheet.batch_clear(['E4:I35'])
        copied_sheet.batch_clear(['A37:A39'])

        print("## 작업 완료...")


def make_date(len: int) -> str:
    """
    주간업무회의 제목 날짜를 생성한다.
        Args:
            len (int): 1: 이번주, 2: 다음주
        Returns:
            str: ex)"2022.09.30. ~ 2022.10.05."
    """
    today_date_info = datetime.date.today()
    start_date = today_date_info - \
        datetime.timedelta(days=today_date_info.weekday())
    endDT = start_date + datetime.timedelta(days=4)
    if len == 2:
        endDT += datetime.timedelta(days=7)

    return "%s ~ %s" % (start_date.strftime("%Y.%m.%d."), endDT.strftime("%Y.%m.%d."))


def today_date() -> str:
    """
    주간업무회의 시트 이름을 생성한다.
        Returns:
            str: ex)"[2022-09-30]"
    """
    return "[%s]" % datetime.date.today().strftime("%Y-%m-%d")


if __name__ == "__main__":
    print("# 새로운 구글 시트를 생성중입니다...")
    create_weekly_sheet(1)
    # print(check_work())
