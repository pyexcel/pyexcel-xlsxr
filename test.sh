pip freeze
nosetests --with-coverage --cover-package pyexcel_xlsxr --cover-package tests tests --with-doctest --doctest-extension=.rst README.rst  pyexcel_xlsxr && flake8 . --exclude=.moban.d,docs --builtins=unicode,xrange,long
