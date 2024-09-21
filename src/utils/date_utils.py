import datetime


def get_current_and_previous_month():
    """
    오늘 기준 이번 달과 지난 달의 YYYYMM 형식 문자열을 반환합니다.

    Returns:
        tuple: (이번 달 YYYYMM, 지난 달 YYYYMM)
    """
    today = datetime.date.today()
    this_month = today.strftime("%Y%m")

    last_month = (today.replace(day=1) -
                  datetime.timedelta(days=1)).strftime("%Y%m")

    return (this_month, last_month)
