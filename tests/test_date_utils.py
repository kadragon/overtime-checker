from pytest import fail

from utils.date_utils import get_current_and_previous_month


from unittest.mock import patch
from datetime import date


def test_january_case():
    with patch('utils.date_utils.datetime.date') as mock_date:
        mock_date.today.return_value = date(2023, 1, 15)
        mock_date.side_effect = lambda *args, **kwargs: date(
            *args, **kwargs)
        current, previous = get_current_and_previous_month()
        if current != "202301":
            fail("Expected current to be '202301'")
        if previous != "202212":
            fail("Expected previous to be '202212'")


def test_february_case():
    with patch('utils.date_utils.datetime.date') as mock_date:
        mock_date.today.return_value = date(2023, 2, 10)
        mock_date.side_effect = lambda *args, **kwargs: date(
            *args, **kwargs)
        current, previous = get_current_and_previous_month()
        if current != "202302":
            fail("Expected current to be '202302'")
        if previous != "202301":
            fail("Expected previous to be '202301'")


def test_june_case():
    with patch('utils.date_utils.datetime.date') as mock_date:
        mock_date.today.return_value = date(2023, 6, 5)
        mock_date.side_effect = lambda *args, **kwargs: date(
            *args, **kwargs)
        current, previous = get_current_and_previous_month()
        if current != "202306":
            fail("Expected current to be '202306'")
        if previous != "202305":
            fail("Expected previous to be '202305'")
