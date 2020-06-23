from datetime import date, datetime, time
from io import UnsupportedOperation

import pyexcel_io.service as service
from pyexcel_io._compact import BytesIO, OrderedDict
from pyexcel_io.book import BookReader
from pyexcel_io.sheet import SheetReader

from pyexcel_xlsxr.messy_xlsx import XLSXBookSet


class XLSXSheet(SheetReader):
    def __init__(
        self,
        sheet,
        auto_detect_int=True,
        auto_detect_float=True,
        auto_detect_datetime=True,
        **keywords
    ):
        SheetReader.__init__(self, sheet, **keywords)
        self.__auto_detect_int = auto_detect_int
        self.__auto_detect_float = auto_detect_float
        self.__auto_detect_datetime = auto_detect_datetime

    @property
    def name(self):
        return self._native_sheet.name

    def row_iterator(self):
        return self._native_sheet.raw()

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


class XLSXBook(BookReader):
    def open(self, file_name, **keywords):
        BookReader.open(self, file_name, **keywords)
        self._load_from_file()

    def open_stream(self, file_stream, **keywords):
        if not hasattr(file_stream, "seek"):
            # python 2
            # Hei zipfile in odfpy would do a seek
            # but stream from urlib cannot do seek
            file_stream = BytesIO(file_stream.read())
        try:
            file_stream.seek(0)
        except UnsupportedOperation:
            # python 3
            file_stream = BytesIO(file_stream.read())
        BookReader.open_stream(self, file_stream, **keywords)
        self._load_from_memory()

    def read_sheet_by_name(self, sheet_name):
        tables = self._native_book.make_tables()
        rets = [table for table in tables if table.name == sheet_name]
        if len(rets) == 0:
            raise ValueError("%s cannot be found" % sheet_name)
        else:
            return self.read_sheet(rets[0])

    def read_sheet_by_index(self, sheet_index):
        """read a sheet at a specified index"""
        tables = self._native_book.make_tables()
        length = len(tables)
        if sheet_index < length:
            return self.read_sheet(tables[sheet_index])
        else:
            raise IndexError(
                "Index %d of out bound %d" % (sheet_index, length)
            )

    def read_all(self):
        """read all sheets"""
        result = OrderedDict()
        for sheet in self._native_book.make_tables():
            ods_sheet = XLSXSheet(sheet, **self._keywords)
            result[ods_sheet.name] = ods_sheet.to_array()

        return result

    def read_sheet(self, native_sheet):
        """read one native sheet"""
        sheet = XLSXSheet(native_sheet, **self._keywords)
        return {sheet.name: sheet.to_array()}

    def _load_from_memory(self):
        self._native_book = XLSXBookSet(self._file_stream)

    def _load_from_file(self):
        self._native_book = XLSXBookSet(self._file_name)
