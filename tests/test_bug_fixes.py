import os
#from nose import eq_
from pyexcel_xlsxr import get_data


def test_issue_1():
    test_file = get_fixture('issue_1.xlsx')
    data = get_data(test_file)


def get_fixture(file_name):
    return os.path.join("tests", "fixtures", file_name)
