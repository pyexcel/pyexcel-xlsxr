from io import BytesIO
from datetime import date, time, datetime

import pyexcel_io.service as service
from pyexcel_io.plugin_api import ISheet, IReader, NamedContent
from pyexcel_xlsxr.messy_xlsx import XLSXBookSet


class XLSXSheet(ISheet):
    def __init__(
        self,
        sheet,
        auto_detect_int=True,
        auto_detect_float=True,
        auto_detect_datetime=True,
    ):
        self.xlsx_sheet = sheet
        self.__auto_detect_int = auto_detect_int
        self.__auto_detect_float = auto_detect_float
        self.__auto_detect_datetime = auto_detect_datetime

    def row_iterator(self):
        return self.xlsx_sheet.raw()

    def column_iterator(self, row):
        for cell in row:
            yield self.__convert_cell(cell)

    def __convert_cell(self, cell):
        if cell is None:
            return None
        if isinstance(cell, (datetime, date, time)):
            return cell
        ret = None
        if isinstance(cell, str):
            if self.__auto_detect_int:
                ret = service.detect_int_value(cell)
            if ret is None and self.__auto_detect_float:
                ret = service.detect_float_value(cell)
                shall_we_ignore_the_conversion = (
                    ret in [float("inf"), float("-inf")]
                ) and self.__ignore_infinity
                if shall_we_ignore_the_conversion:
                    ret = None
        if ret is None:
            ret = cell
        return ret


class XLSXBook(IReader):
    def __init__(self, file_alike_object, _, **keywords):
        self.xlsx_book = XLSXBookSet(file_alike_object)
        self._keywords = keywords
        tables = self.xlsx_book.make_tables()
        self.content_array = [
            NamedContent(table.name, table) for table in tables
        ]

    def read_sheet(self, sheet_index):
        """read a sheet at a specified index"""
        table = self.content_array[sheet_index].payload
        sheet = XLSXSheet(table, **self._keywords)
        return sheet

    def close(self):
        self.xlsx_book.close()


class XLSXBookInContent(XLSXBook):
    def __init__(self, file_content, file_type, **keywords):
        file_stream = BytesIO(file_content)
        super().__init__(file_stream, file_type, **keywords)
