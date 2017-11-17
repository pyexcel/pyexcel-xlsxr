from nose.tools import eq_
from pyexcel_xlsxr.messy_xlsx import find_sheets
from pyexcel_xlsxr.messy_xlsx import parse_row
from pyexcel_xlsxr.messy_xlsx import parse_styles
from pyexcel_xlsxr.messy_xlsx import parse_book_properties
from pyexcel_xlsxr.messy_xlsx import parse_xfs_styles
from datetime import datetime, time

def test_list_one():
    test_sample = [
        '_rels/.rels',
        'xl/_rels/workbook.xml.rels',
        'xl/worksheets/sheet2.xml',
        'xl/worksheets/sheet3.xml',
        'xl/worksheets/sheet1.xml',
        'xl/sharedStrings.xml',
        'xl/workbook.xml',
        'xl/styles.xml',
        'docProps/app.xml',
        'docProps/core.xml',
        '[Content_Types].xml'
    ]

    expected = [
        'xl/worksheets/sheet2.xml',
        'xl/worksheets/sheet3.xml',
        'xl/worksheets/sheet1.xml',
    ]

    sheet_files = find_sheets(test_sample)
    eq_(sheet_files, expected)


def test_alternative_file_list():
    test_sample = [
        '_rels/.rels',
        'xl/_rels/workbook.xml.rels',
        'xl/worksheets/worksheet2.xml',
        'xl/worksheets/worksheet3.xml',
        'xl/worksheets/worksheet1.xml',
        'xl/sharedStrings.xml',
        'xl/workbook.xml',
        'xl/styles.xml',
        'docProps/app.xml',
        'docProps/core.xml',
        '[Content_Types].xml'
    ]

    expected = [
        'xl/worksheets/worksheet2.xml',
        'xl/worksheets/worksheet3.xml',
        'xl/worksheets/worksheet1.xml',
    ]

    sheet_files = find_sheets(test_sample)
    eq_(sheet_files, expected)


def test_single_sheet():
    test_sample = [
        '_rels/.rels',
        'xl/_rels/workbook.xml.rels',
        'xl/worksheets/sheet.xml',
        'xl/sharedStrings.xml',
        'xl/workbook.xml',
        'xl/styles.xml',
        'docProps/app.xml',
        'docProps/core.xml',
        '[Content_Types].xml'
    ]

    expected = [
        'xl/worksheets/sheet.xml'
    ]

    sheet_files = find_sheets(test_sample)
    eq_(sheet_files, expected)

    
def test_alternative_single_sheet():
    test_sample = [
        '_rels/.rels',
        'xl/_rels/workbook.xml.rels',
        'xl/worksheets/worksheet.xml',
        'xl/sharedStrings.xml',
        'xl/workbook.xml',
        'xl/styles.xml',
        'docProps/app.xml',
        'docProps/core.xml',
        '[Content_Types].xml'
    ]

    expected = [
        'xl/worksheets/worksheet.xml'
    ]

    sheet_files = find_sheets(test_sample)
    eq_(sheet_files, expected)


def test_parse_row():
    xml_string = '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.75" outlineLevel="0" r="4"><c r="A4" s="1" t="n"><v>42005</v></c><c r="B4" s="2" t="n"><v>0.550844907407407</v></c></row>'  # flake8: noqa
    class Book:
        def __init__(self):
            self.xfs_styles = [1, 1, 2]
            self.styles = {'1': 'dd/mm/yy', '2': 'h:mm:ss;@'}
            self.properties = {'date1904': False}
    data = parse_row(xml_string, Book())
    eq_([cell.value for cell in data],
        [datetime(year=2015, month=1, day=1), '13:13'])


def test_parse_styles():
    sample = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt formatCode="GENERAL" numFmtId="164"/><numFmt formatCode="DD/MM/YY" numFmtId="165"/><numFmt formatCode="H:MM:SS;@" numFmtId="166"/></numFmts><fonts count="4"><font><name val="Arial"/>'
    styles = parse_styles(sample)
    eq_(list(styles.values()),
        ['general', 'dd/mm/yy', 'h:mm:ss;@'])


def test_parse_properties():
    sample = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>'
    properties = parse_book_properties(sample)
    eq_(properties, {'date1904': False})


def test_parse_xfs_styles():
    sample = '<cellXfs count="3"><xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"></xf><xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"></xf><xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="166" xfId="0"></xf></cellXfs><cellStyles count="6">'
    xfs_styles = parse_xfs_styles(sample)
    eq_(xfs_styles, [164, 165, 166])
