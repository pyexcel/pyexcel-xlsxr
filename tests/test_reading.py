import os
from pyexcel_xlsxr import get_data


def test_reading():
    data = get_data(os.path.join("tests", "fixtures", "date_field.xlsx"),
                    library='pyexcel-xlsxr')
    print(data)

