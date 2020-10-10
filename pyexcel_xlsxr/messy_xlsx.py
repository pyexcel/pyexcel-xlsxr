import io
import re
import zipfile
from datetime import time, datetime, timedelta

from lxml import etree
from pyexcel_io._compact import OrderedDict

STYLE_FILENAME = "xl/styles.xml"
SHARED_STRING = "xl/sharedStrings.xml"
WORK_BOOK = "xl/workbook.xml"
SHEET_MATCHER = "xl/worksheets/(work)?sheet([0-9]+)?.xml"
SHEET_INDEX_MATCHER = "xl/worksheets/(work)?sheet(([0-9]+)?).xml"
XLSX_ROW_MATCH = re.compile(rb".*?(<row.*?<\/.*?row>).*?", re.MULTILINE)
NUMBER_FMT_MATCHER = re.compile(
    rb".*?(<numFmts.*?<\/.*?numFmts>).*?", re.MULTILINE
)
XFS_FMT_MATCHER = re.compile(
    rb".*?(<cellXfs.*?<\/.*?cellXfs>).*?", re.MULTILINE
)
SHEET_FMT_MATCHER = re.compile(rb".*?(<sheet .*?\/>).*?", re.MULTILINE)
DATE_1904_MATCHER = re.compile(rb".*?(<workbookPr.*?\/>).*?", re.MULTILINE)
# "xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
# But it not used for now
X14AC_NAMESPACE = b'xmlns:x14ac="http://not.used.com/"'

# see also ruby-roo lib at: http://github.com/hmcgowan/roo
FORMATS = {
    "general": "float",
    "0": "float",
    "0.00": "float",
    "#,##0": "float",
    "#,##0.00": "float",
    "0%": "percentage",
    "0.00%": "percentage",
    "0.00e+00": "float",
    "mm-dd-yy": "date",
    "d-mmm-yy": "date",
    "d-mmm": "date",
    "mmm-yy": "date",
    "h:mm am/pm": "date",
    "h:mm:ss am/pm": "date",
    "h:mm": "time",
    "h:mm:ss": "time",
    "m/d/yy h:mm": "date",
    "#,##0 ;(#,##0)": "float",
    "#,##0 ;[red](#,##0)": "float",
    "#,##0.00;(#,##0.00)": "float",
    "#,##0.00;[red](#,##0.00)": "float",
    "mm:ss": "time",
    "[h]:mm:ss": "time",
    "mmss.0": "time",
    "##0.0e+0": "float",
    "@": "float",
    "yyyy\\-mm\\-dd": "date",
    "dd/mm/yy": "date",
    "hh:mm:ss": "time",
    "dd/mm/yy\\ hh:mm": "date",
    "dd/mm/yyyy hh:mm:ss": "date",
    "yy-mm-dd": "date",
    "d-mmm-yyyy": "date",
    "m/d/yy": "date",
    "m/d/yyyy": "date",
    "dd-mmm-yyyy": "date",
    "dd/mm/yyyy": "date",
    "mm/dd/yy h:mm am/pm": "date",
    "mm/dd/yy hh:mm": "date",
    "mm/dd/yyyy h:mm am/pm": "date",
    "mm/dd/yyyy hh:mm:ss": "date",
    "yyyy-mm-dd hh:mm:ss": "date",
    "#,##0;(#,##0)": "float",
    '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)': "float",
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)': "float",
}
STANDARD_FORMATS = {
    0: "general",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00e+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm am/pm",
    19: "h:mm:ss am/pm",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0e+0",
    49: "@",
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
        if hasattr(file_alike, "read"):
            file_alike = io.BytesIO(file_alike.read())
        self.zip_file = zipfile.ZipFile(file_alike)
        self.styles, self.xfs_styles = self.__extract_styles()
        self.properties = self.__extract_book_properties()
        self.shared_strings = list(self.__extract_shared_strings())

    def __extract_shared_strings(self):
        try:
            shared_string_content = self.zip_file.open(SHARED_STRING).read()
            return parse_shared_strings(shared_string_content)
        except KeyError:
            return []

    def __extract_styles(self):
        style_content = self.zip_file.open(STYLE_FILENAME).read()
        return parse_styles(style_content), parse_xfs_styles(style_content)

    def __extract_book_properties(self):
        book_content = self.zip_file.open(WORK_BOOK).read()
        return parse_book_properties(book_content)

    def close(self):
        if self.zip_file:
            self.zip_file.close()

    def make_tables(self):
        sheet_files = find_sheets(self.zip_file.namelist())
        for sheet_file in sorted(sheet_files):
            content = self.zip_file.open(sheet_file).read()
            sheet_index = get_sheet_index(sheet_file)
            sheet_name = self.properties["sheets"][sheet_index]
            yield XLSXTable(sheet_name, content, self)


def find_sheets(file_list):

    return [
        sheet_file
        for sheet_file in file_list
        if re.match(SHEET_MATCHER, sheet_file)
    ]


def get_sheet_index(file_name):
    if re.match(SHEET_MATCHER, file_name):
        result = re.search(SHEET_INDEX_MATCHER, file_name)
        index = int(result.group(3)) if result.group(3) else 1
        return index - 1
    else:
        raise Exception("Invalid sheet file name")


class Cell(object):
    def __init__(self):
        self.column_type = ""
        self.style_string = ""
        self.value = ""
        self.type = ""

    def __repr__(self):
        return str(self.value)


def parse_row(row_xml_string, book):
    if b"x14ac" in row_xml_string:
        row_xml_string = row_xml_string.replace(
            b"<row", (b"<row " + X14AC_NAMESPACE)
        )
    partial = io.BytesIO(row_xml_string)
    cells = []
    cell = Cell()

    for action, element in etree.iterparse(partial):

        if element.tag in ["v", "t"]:
            cell.value = element.text
        elif element.tag in ["c"]:
            local_type = element.attrib.get("t")
            cell.column_type = local_type
            style_int = element.attrib.get("s")
            if style_int:
                xfs_style_int = book.xfs_styles[int(style_int)]
                cell.style_string = book.styles.get(str(xfs_style_int))
            parse_cell(cell, book)
            cells.append(cell)
            cell = Cell()
    return [c.value for c in cells]


def parse_cell(cell, book):
    cell.type = parse_cell_type(cell)
    parse_cell_value(cell, book)


def parse_cell_type(cell):
    cell_type = None
    if cell.style_string:
        date_time_flag = (
            re.match(r"^\d+(\.\d+)?$", cell.value)
            and re.match(".*[hsmdyY]", cell.style_string)
            and not re.match(r".*\[.*[dmhys].*\]", cell.style_string)
        )
        if cell.style_string in FORMATS:
            cell_type = FORMATS[cell.style_string]
        elif date_time_flag:
            if float(cell.value) < 1:
                cell_type = "time"
            else:
                cell_type = "date"
        elif re.match(r"^-?\d+(.\d+)?$", cell.value):
            cell_type = "float"
    return cell_type


def parse_cell_value(cell, book):
    if cell.column_type == "s":
        cell.value = book.shared_strings[int(cell.value)]
    elif cell.column_type == "b":
        cell.value = (
            (int(cell.value) == 1 and "TRUE")
            or (int(cell.value) == 0 and "FALSE")
            or cell.value
        )
    elif cell.column_type == "n":
        parse_numeric_cell_value(cell, book)
    # else
    #   no action


def parse_numeric_cell_value(cell, book):
    try:
        if cell.type == "date":  # date/time
            if book.properties["date1904"]:
                start = datetime(1904, 1, 1)
            else:
                start = datetime(1899, 12, 30)
            cell.value = start + timedelta(float(cell.value))
        elif cell.type == "time":  # time
            # round to microseconds
            seconds_in_total = int(
                round((float(cell.value) % 1) * 24 * 60 * 60, 6)
            )
            minutes_in_total = int(seconds_in_total / 60)
            second = int(minutes_in_total % 60)
            hour = int(minutes_in_total / 60)
            # str(t / 60) + ":" + ('0' + str(t % 60))[-2:]
            cell.value = time(
                hour=hour, minute=minutes_in_total % 60, second=second
            )
        elif cell.type == "float" and ("E" in cell.value or "e" in cell.value):
            cell.value = ("%f" % (float(cell.value))).rstrip("0").rstrip(".")
    except (ValueError, OverflowError):
        # invalid date format
        pass


def parse_styles(style_content):
    styles = OrderedDict()
    formats = NUMBER_FMT_MATCHER.findall(style_content)
    for aformat in formats:
        partial = io.BytesIO(aformat)
        for action, element in etree.iterparse(partial, ("end",)):
            if element.tag != "numFmt":
                continue
            numFmtId = element.attrib.get("numFmtId")
            formatCode = element.attrib.get("formatCode")
            formatCode = formatCode.lower().replace("\\", "")
            styles.update({numFmtId: formatCode})
    return styles


def parse_xfs_styles(style_content):
    styles = []
    formats = XFS_FMT_MATCHER.findall(style_content)
    for aformat in formats:
        partial = io.BytesIO(aformat)
        for action, element in etree.iterparse(partial, ("end",)):
            if element.tag != "xf":
                continue
            styles.append(int(element.attrib.get("numFmtId")))
    return styles


def parse_book_properties(book_content):
    properties = {"sheets": []}
    date1904 = DATE_1904_MATCHER.findall(book_content)
    for apr in date1904:
        partial = io.BytesIO(apr)
        for action, element in etree.iterparse(partial, ("end",)):
            if element.tag == "workbookPr":
                value = element.attrib.get("date1904")
                if value:
                    properties["date1904"] = value.lower().strip() == "true"
                else:
                    properties["date1904"] = False

    ns = (
        "http://schemas.openxmlformats.org/"
        + "officeDocument/2006/relationships"
    )
    namespaces = {"r": ns}

    xlsx_header = u"<wrapper {0}>".format(
        " ".join('xmlns:{0}="{1}"'.format(k, v) for k, v in namespaces.items())
    ).encode("utf-8")
    xlsx_footer = u"</wrapper>".encode("utf-8")
    sheets = SHEET_FMT_MATCHER.findall(book_content)
    for sheet in sheets:
        block = xlsx_header + sheet + xlsx_footer
        partial = io.BytesIO(block)
        for action, element in etree.iterparse(partial, ("end",)):
            if element.tag == "sheet":
                value = element.attrib.get("name")
                properties["sheets"].append(value)
    return properties


def parse_shared_strings(content):
    root = etree.fromstring(content)
    for si in root.iterchildren():
        content = ""
        for text in si.iterchildren():
            if text.text:
                content += text.text
        yield (content)
