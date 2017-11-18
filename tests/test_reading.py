import os
from datetime import datetime

from nose.tools import eq_
from pyexcel_xlsxr import get_data


def test_reading():
    data = get_data(os.path.join("tests", "fixtures", "date_field.xlsx"),
                    library='pyexcel-xlsxr')
    expected = {
        "sheet1": [
            ['Date', 'Time'],
            [datetime(year=2014, month=12, day=25), '11:11'],
            [datetime(2014, 12, 26, 0, 0), '12:12'],
            [datetime(2015, 1, 1, 0, 0), '13:13'],
            [datetime(year=1899, month=12, day=30), '00:00']
        ],
        "sheet2": [],
        "sheet3": []
    }
    eq_(data, expected)
