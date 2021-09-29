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
the [releases page](https://github.com/ngoonee/xls_summary/releases). You may
need to install the [Visual C++ Redistributable for Visual Studio 2015](https://www.microsoft.com/en-us/download/details.aspx?id=48145).

For using pyinstaller, first install the dependencies (check setup.py, mainly
just openpyxl, and of course pyinstaller). The command is:-

pyinstaller.exe --onefile app.py

Can be done with wine as well, if on a Linux only system trying to generate
the exe. Tested (29 Sep 2021) with python 3.9.7. Install pyinstaller in wine,
then generate the exe using the above command.

Requirements
^^^^^^^^^^^^

Compatibility
-------------

Licence
-------

Authors
-------

`xls_summary` was written by `Ng Oon-Ee <ngoe@utar.edu.my>`_.
