"""
    pyexcel_xlsxr
    ~~~~~~~~~~~~~~~~~~~
    The lower level xlsx file format handler using lxml
    :copyright: (c) 2015-2020 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
from pyexcel_io.io import get_data as read_data
from pyexcel_io.io import isstream
from pyexcel_io.plugins import IOPluginInfoChainV2
from pyexcel_xlsxr._version import __author__, __version__  # noqa

__FILE_TYPE__ = "xlsx"

IOPluginInfoChainV2(__name__).add_a_reader(
    relative_plugin_class_path="xlsxr.XLSXBook",
    locations=["file", "memory"],
    file_types=[__FILE_TYPE__],
    stream_type="binary",
).add_a_reader(
    relative_plugin_class_path="xlsxr.XLSXBookInContent",
    locations=["content"],
    file_types=[__FILE_TYPE__],
    stream_type="binary",
)


def get_data(afile, file_type=None, **keywords):
    """standalone module function for reading module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = __FILE_TYPE__
    return read_data(afile, file_type=file_type, **keywords)
