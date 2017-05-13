xls_summary
===========

.. image:: https://img.shields.io/pypi/v/xls_summary.svg
    :target: https://pypi.python.org/pypi/xls_summary
    :alt: Latest PyPI version

.. image:: https://travis-ci.org/ngoonee/xls_summary.png
   :target: https://travis-ci.org/ngoonee/xls_summary
   :alt: Latest Travis CI build status

Summarizes multiple copies of the same excel spreadsheet

Usage
-----

For windows, either use pip as usual or use pyinstaller to generate an exe.

For using pyinstaller, first install the dependencies (check setup.py). The
openpyxl dependency requires a hook file to be properly detected by
pyinstaller, and xlrd requires the --hidden-import option. The command is:-

pyinstaller.exe --additional-hooks-dir=. --hidden-import=xlrd --onefile app.py

Can be done with wine as well, of on a Linux only system trying to generate
the exe. As of today (14 May 2017) you'll need wine-staging and python 3.5.x
(the 3.6 series does not work with wine). Install pyinstaller in wine, then
generate the exe using the above command.

Installation
------------

Requirements
^^^^^^^^^^^^

Compatibility
-------------

Licence
-------

Authors
-------

`xls_summary` was written by `Ng Oon-Ee <ngoe@utar.edu.my>`_.
