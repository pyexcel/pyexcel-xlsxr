{% extends "README.rst.jj2" %}

{%block documentation_link%}
{%endblock%}

{%block description %}
**{{name}}** is a specialized xlsx reader using lxml. It does partial reading, meaning
it wont load all content into memory.


lxml installation
=================

This library depends on lxml. Because its availablity, the use of this library is restricted.

for PyPy, lxml == 3.4.4 are tested to work well. But lxml above 3.4.4 is difficult to get installed.

for Python 3.7, please use lxml==4.1.1.

Otherwise, this library works OK with lxml 3.4.4 or above.


{%endblock%}

{% block write_to_file %}

.. testcode::
   :hide:

    >>> from pyexcel_xlsxw import save_data
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> save_data("your_file.{{file_type}}", data)

{% endblock %}


{% block write_to_memory %}

.. testcode::
   :hide:

    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = StringIO()
    >>> save_data(io, data)
    >>> unused = io.seek(0)
    >>> # do something with the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading


{%endblock%}

{% block pyexcel_write_to_file%}

.. testcode::
   :hide:

    >>> sheet.save_as("another_file.{{file_type}}")

{% endblock %}

{% block pyexcel_write_to_memory%}
{% endblock %}
