import datetime  # noqa
import os  # noqa

import pyexcel
from nose.tools import eq_, raises  # noqa


def create_sample_file1(file):
    data = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", 1.1, 1]
    table = []
    table.append(data[:4])
    table.append(data[4:8])
    table.append(data[8:12])
    pyexcel.save_as(array=table, dest_file_name=file)
