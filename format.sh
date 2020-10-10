isort $(find pyexcel_xlsxr -name "*.py"|xargs echo) $(find tests -name "*.py"|xargs echo)
black -l 79 pyexcel_xlsxr
black -l 79 tests
