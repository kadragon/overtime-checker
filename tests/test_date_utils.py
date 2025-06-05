import os
import sys
from datetime import date

# Ensure src directory is on the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from utils.date_utils import get_current_and_previous_month


def test_get_current_and_previous_month():
    current, previous = get_current_and_previous_month()

    assert isinstance(current, str) and len(current) == 6 and current.isdigit()
    assert isinstance(previous, str) and len(previous) == 6 and previous.isdigit()

    year = int(current[:4])
    month = int(current[4:])
    if month == 1:
        prev_year = year - 1
        prev_month = 12
    else:
        prev_year = year
        prev_month = month - 1
    expected_previous = f"{prev_year:04d}{prev_month:02d}"

    assert previous == expected_previous

