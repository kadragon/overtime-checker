"""
Manages the creation and updating of weekly work plan Google Sheets.
"""
import gspread
from contextlib import contextmanager
import datetime
from config import SPREADSHEET_ID, GOOGLE_CREDENTIALS_FILE


@contextmanager
def google_sheets():
    """
    A context manager for authenticating and providing a gspread client instance.

    Yields:
        gspread.client.Client: An authenticated gspread client.
    """
    gc = gspread.service_account(GOOGLE_CREDENTIALS_FILE)
    yield gc


def create_weekly_sheet(length: int):
    """
    Creates a new weekly sheet in Google Sheets by copying a template.

    It updates the new sheet's title with the current date, sets the date range
    for the week, clears previous content from designated areas, and sets the
    sheet as the first one in the workbook.

    Args:
        length (int): Specifies the week to create the sheet for.
                      1 for the current week, 2 for the next week.
    """
    with google_sheets() as gc:
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        copied_sheet = spreadsheet.get_worksheet(0).copy_to(SPREADSHEET_ID)
        copied_sheet = spreadsheet.worksheet(copied_sheet['title'])

        print("## 복사 완료...")

        # Update copied sheet title
        title = today_date()
        copied_sheet.update_title(title)

        # Set date
        copied_sheet.update_acell('A2', make_date(length))

        # index update
        copied_sheet.update_index(0)

        print("## 내용 정리 시작...")

        # Clear contents
        copied_sheet.batch_clear(['C4:C35'])
        copied_sheet.batch_clear(['E4:I35'])
        copied_sheet.batch_clear(['A37:A39'])

        print("## 작업 완료...")


def make_date(length: int) -> str:
    """
    Generates the date string for the weekly work meeting title.

    Args:
        length (int): 1 for the current week, 2 for the next week.
    Returns:
        str: Date range string, e.g., "2022.09.30. ~ 2022.10.05."
    """
    today_date_info = datetime.date.today()
    # Calculate the start of the current week (Monday)
    start_date = today_date_info - datetime.timedelta(days=today_date_info.weekday())
    # Calculate the end of the work week (Friday)
    end_date = start_date + datetime.timedelta(days=4)

    if length == 2: # If next week is requested
        start_date += datetime.timedelta(days=7)
        end_date += datetime.timedelta(days=7)

    return f"{start_date.strftime('%Y.%m.%d.')} ~ {end_date.strftime('%Y.%m.%d.')}"


def today_date() -> str:
    """
    Generates the sheet name for the weekly work meeting based on the current date.

    Returns:
        str: Sheet name string, e.g., "[2022-09-30]"
    """
    return f"[{datetime.date.today().strftime('%Y-%m-%d')}]"


if __name__ == "__main__":
    print("# 새로운 구글 시트를 생성중입니다...")
    create_weekly_sheet(1)
    # print(check_work())
