import io
import re
from lxml import etree
import zipfile
from datetime import datetime, timedelta


STYLE_FILENAME = "xl/styles.xml"
SHEET_MATCHER = 'xl/worksheets/(work)?sheet([0-9]+)?.xml'
XLSX_ROW_MATCH = re.compile(
    b".*?(<row.*?<\/.*?row>).*?",
    re.MULTILINE)
NUMBER_FMT_MATCHER = re.compile(
    b".*?(<numFmts.*?<\/.*?numFmts>).*?",
    re.MULTILINE)
XFS_FMT_MATCHER = re.compile(
    b".*?(<cellXfs.*?<\/.*?cellXfs>).*?",
    re.MULTILINE)
DATE_1904_MATCHER = re.compile(
    b".*?(<workbookPr.*?\/>).*?",
    re.MULTILINE)

# see also ruby-roo lib at: http://github.com/hmcgowan/roo
FORMATS = {
  'general': 'float',
  '0': 'float',
  '0.00': 'float',
  '#,##0': 'float',
  '#,##0.00': 'float',
  '0%': 'percentage',
  '0.00%': 'percentage',
  '0.00e+00': 'float',
  'mm-dd-yy': 'date',
  'd-mmm-yy': 'date',
  'd-mmm': 'date',
  'mmm-yy': 'date',
  'h:mm am/pm': 'date',
  'h:mm:ss am/pm': 'date',
  'h:mm': 'time',
  'h:mm:ss': 'time',
  'm/d/yy h:mm': 'date',
  '#,##0 ;(#,##0)': 'float',
  '#,##0 ;[red](#,##0)': 'float',
  '#,##0.00;(#,##0.00)': 'float',
  '#,##0.00;[red](#,##0.00)': 'float',
  'mm:ss': 'time',
  '[h]:mm:ss': 'time',
  'mmss.0': 'time',
  '##0.0e+0': 'float',
  '@': 'float',
  'yyyy\\-mm\\-dd': 'date',
  'dd/mm/yy': 'date',
  'hh:mm:ss': 'time',
  "dd/mm/yy\\ hh:mm": 'date',
  'dd/mm/yyyy hh:mm:ss': 'date',
  'yy-mm-dd': 'date',
  'd-mmm-yyyy': 'date',
  'm/d/yy': 'date',
  'm/d/yyyy': 'date',
  'dd-mmm-yyyy': 'date',
  'dd/mm/yyyy': 'date',
  'mm/dd/yy h:mm am/pm': 'date',
  'mm/dd/yy hh:mm': 'date',
  'mm/dd/yyyy h:mm am/pm': 'date',
  'mm/dd/yyyy hh:mm:ss': 'date',
  'yyyy-mm-dd hh:mm:ss': 'date',
  '#,##0;(#,##0)': 'float',
  '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)': 'float',
  '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)': 'float'
}
STANDARD_FORMATS = {
  0: 'general',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00e+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm am/pm',
  19: 'h:mm:ss am/pm',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0e+0',
  49: '@',
}


class XLSXTable(object):
    def __init__(self, name, file_content, book):
        self.name = name
        self.content = file_content
        self.book = book

    def raw(self):
        rows = XLSX_ROW_MATCH.findall(self.content)
        for row in rows:
            yield parse_row(row, self.book)


class XLSXBookSet(object):
    def __init__(self, file_alike, **keywords):
        if hasattr(file_alike, 'read'):
            file_alike = io.BytesIO(file_alike.read())
        self.zip_file = zipfile.ZipFile(file_alike)
        self.styles, self.xfs_styles = self.__extract_styles()
        self.properties = self.__extract_book_properties()

    def __extract_styles(self):
        style_content = self.zip_file.open(STYLE_FILENAME).read()
        return parse_styles(style_content), parse_xfs_styles(style_content)

    def __extract_book_properties(self):
        book_content = self.zip_file.open("xl/workbook.xml").read()
        return parse_book_properties(book_content)

    def __del__(self):
        self.zip_file.close()

    def make_tables(self):
        sheet_files = find_sheets(self.zip_file.namelist())
        for sheet_file in sheet_files:
            content = self.zip_file.open(sheet_file).read()
            yield XLSXTable(sheet_file, content, self)


def find_sheets(file_list):
    return [sheet_file for sheet_file in file_list
            if re.match(SHEET_MATCHER, sheet_file)]


class Cell(object):
    def __init__(self):
        self.type_string = ''
        self.style_string = ''
        self.value = ''
        self.type = ''

    def __repr__(self):
        return str(self.value)


def parse_row(row_xml_string, book):
    partial = io.BytesIO(row_xml_string)
    cells = []
    cell = Cell()
    for action, element in etree.iterparse(partial, ('end',)):
        print(element.tag)
        if element.tag == 'v':
            cell.value = element.text
        elif element.tag == 'c':
            local_type = element.attrib.get('t')
            style_int = element.attrib.get('s')
            xfs_style_int = book.xfs_styles[int(style_int)]
            cell.type_string = local_type
            cell.style_string = book.styles.get(str(xfs_style_int))
            parse_cell(cell, book)
            cells.append(cell)
            cell = Cell()
    return cells


def parse_cell(cell, book):
    cell.type = parse_cell_type(cell)
    parse_cell_value(cell, book)


def parse_cell_type(cell):
    cell_type = None
    if cell.style_string:
        print(cell.style_string)
        date_time_flag = (
            re.match("^\d+(\.\d+)?$", cell.value) and
            re.match(".*[hsmdyY]", cell.style_string) and
            not re.match('.*\[.*[dmhys].*\]', cell.style_string))
        if cell.style_string in FORMATS:
            cell_type = FORMATS[cell.style_string]
        elif date_time_flag:
            if float(cell.value) < 1:
                cell_type = "time"
            else:
                cell_type = "date"
        elif re.match("^-?\d+(.\d+)?$", cell.value):
            cell_type = "float"
    return cell_type


def parse_cell_value(cell, book):
    print(cell.value, cell.type)
    try:
        if cell.type == 'date':  # date/time
            if book.properties['date1904']:
                start = datetime(1904, 1, 1)
            else:
                start = datetime(1899, 12, 30)
            print(cell.value)
            cell.value = start + timedelta(float(cell.value))
        elif cell.type == 'time':  # time
            # round to microseconds
            t = int(round((float(cell.value) % 1) * 24 * 60 * 60, 6)) / 60
            # str(t / 60) + ":" + ('0' + str(t % 60))[-2:]
            cell.value = "%.2i:%.2i" % (t / 60, t % 60)
        elif cell.type == 'float' and ('E' in cell.value or 'e' in cell.value):
            cell.value = ("%f" % (float(cell.value))).rstrip('0').rstrip('.')
    except (ValueError, OverflowError):
        # invalid date format
        pass


def parse_styles(style_content):
    styles = {}
    formats = NUMBER_FMT_MATCHER.findall(style_content)
    for aformat in formats:
        partial = io.BytesIO(aformat)
        for action, element in etree.iterparse(partial, ('end',)):
            if element.tag != 'numFmt':
                continue
            numFmtId = element.attrib.get('numFmtId')
            formatCode = element.attrib.get('formatCode')
            formatCode = formatCode.lower().replace('\\', '')
            styles.update({
                numFmtId: formatCode
            })
    return styles


def parse_xfs_styles(style_content):
    styles = []
    formats = XFS_FMT_MATCHER.findall(style_content)
    for aformat in formats:
        partial = io.BytesIO(aformat)
        for action, element in etree.iterparse(partial, ('end',)):
            if element.tag != 'xf':
                continue
            styles.append(int(element.attrib.get('numFmtId')))
    return styles


def parse_book_properties(book_content):
    properties = {}
    date1904 = DATE_1904_MATCHER.findall(book_content)
    for apr in date1904:
        partial = io.BytesIO(apr)
        for action, element in etree.iterparse(partial, ('end', )):
            if element.tag == 'workbookPr':
                value = element.attrib.get('date1904')
                properties['date1904'] = value.lower().strip() == 'true'
    return properties