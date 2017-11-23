import os
from textwrap import dedent
from nose.tools import eq_

import pyexcel as pe


class TestDateFormat:
    def test_reading_date_format(self):
        """
        date     time
        25/12/14 11:11:11
        25/12/14 12:12:12
        01/01/15 13:13:13
        0.0      0.0
        """
        import datetime
        r = pe.get_sheet(file_name=os.path.join("tests", "fixtures",
                                                "date_field.xlsx"),
                         library='pyexcel-xlsxr')
        assert isinstance(r[1, 0], datetime.date)
        eq_(r[1, 0].strftime("%d/%m/%y"), "25/12/14")
        assert isinstance(r[1, 1], datetime.time) is True
        assert r[1, 1].strftime("%H:%M:%S") == "11:11:11"
        value = r[4, 0].isoformat()
        eq_(value, "1899-12-30T00:00:00")
        eq_(r[4, 1].isoformat(), "00:00:00")


class TestAutoDetectInt:
    def setUp(self):
        self.content = [[1, 2, 3.1]]
        self.test_file = "test_auto_detect_init.xlsx"
        pe.save_as(
            array=self.content, dest_file_name=self.test_file
         )

    def test_auto_detect_int(self):
        sheet = pe.get_sheet(file_name=self.test_file, library="pyexcel-xlsxr")
        expected = dedent("""
        pyexcel_sheet1:
        +---+---+-----+
        | 1 | 2 | 3.1 |
        +---+---+-----+""").strip()
        eq_(str(sheet), expected)

    def test_get_book_auto_detect_int(self):
        book = pe.get_book(file_name=self.test_file, library="pyexcel-xlsxr")
        expected = dedent("""
        pyexcel_sheet1:
        +---+---+-----+
        | 1 | 2 | 3.1 |
        +---+---+-----+""").strip()
        eq_(str(book), expected)

    def test_auto_detect_int_false(self):
        sheet = pe.get_sheet(file_name=self.test_file, auto_detect_int=False,
                             library="pyexcel-xlsxr")
        expected = dedent("""
        pyexcel_sheet1:
        +-----+-----+-----+
        | 1.0 | 2.0 | 3.1 |
        +-----+-----+-----+""").strip()
        eq_(str(sheet), expected)

    def test_get_book_auto_detect_int_false(self):
        book = pe.get_book(file_name=self.test_file, auto_detect_int=False,
                           library="pyexcel-xlsxr")
        expected = dedent("""
        pyexcel_sheet1:
        +-----+-----+-----+
        | 1.0 | 2.0 | 3.1 |
        +-----+-----+-----+""").strip()
        eq_(str(book), expected)

    def tearDown(self):
        os.unlink(self.test_file)
