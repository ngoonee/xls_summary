xls_summary
===========

Summarizes multiple copies of the same excel spreadsheet.

Usage
-----

1. Place all the Excel files you're summarizing in a single folder

2. Copy one of them and rename it to "summary_template.xlsx"

3. Edit "summary_template.xlsx" and place ::A:: in the cell which should be
summarized to the A column, ::B:: in the cell which should be summarized to the
B column and so on.

4. Run your binary/executable and point it to the folder with all your Excel
spreadsheets. The summary is generated in "summary.xlsx"

Installation
------------

For Windows, either use pip as usual or use pyinstaller to generate an exe.
Most Windows users would prefer to use the precompiled binaries available at
the [releases page](https://github.com/ngoonee/xls_summary/releases).

For using pyinstaller, first install the dependencies (check setup.py). The
openpyxl dependency requires a hook file to be properly detected by
pyinstaller, and xlrd requires the --hidden-import option. The command is:-

pyinstaller.exe --additional-hooks-dir=. --hidden-import=xlrd --onefile app.py

Can be done with wine as well, if on a Linux only system trying to generate
the exe. As of today (19 Oct 2017) you'll need to use python 3.4.x
(due to a bug in pyinstaller for 3.5 and up). Install pyinstaller in wine, then
generate the exe using the above command.

Requirements
^^^^^^^^^^^^

Compatibility
-------------

Licence
-------

Authors
-------

`xls_summary` was written by `Ng Oon-Ee <ngoe@utar.edu.my>`_.
