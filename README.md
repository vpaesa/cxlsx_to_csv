#cxlsx_to_csv
**cxlsx_to_csv is a simple converter of Excel 2007 files (a.k.a. Open Office XML) to .CSV**

###FEATURES:
* Simple converter, pretty fast as done in C.
* Only depends on [miniz](https://code.google.com/p/miniz/) (included for convenience) and [expat](http://expat.sourceforge.net/).

The .XLSX format is just a glorified .ZIP (that I open thanks to miniz), containing a set of .XML files (that I parse thanks to Expat).
Dates are exported as the number of days that have elapsed since 1-January-1900 (the Excel Epoch).

###SYNOPSIS:
```
cxlsx_to_csv -if input.xlsx [-sh sheet_id] [-of output.csv]
    input.xlsx  input spreadsheet in Excel 2007 format (Office Open XML)
    sheet_id    number of the sheet within the workbook (default is first one)
    output.csv  output CSV file (default is STDOUT)
```
###SPEED COMPARATION:
* Tested under Ubuntu 15.04 on an Intel(R) Core(TM) i3-3217U CPU @ 1.80GHz, with SSD.
* Locale set to LC_ALL=C.UTF-8
* The input spreadsheet has 47544 rows and 46 columns, and weights 8962KB.

| Tool | Command | Time (real/user/sys)|
|:------------ |:------------|:--|
| [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv) | `cxlsx_to_csv -if a.xlsx -of a.csv` | `3.036s/2.836s/0.200s` |
| [Gnumeric](http://www.gnumeric.org/) | `ssconvert --export-type=Gnumeric_stf:stf_assistant -O 'eol=windows separator=, format=raw transliterate-mode=escape quoting-mode=auto' a.xlsx a.csv` | `11.735s/11.544s/0.196s` |
| [LibreOffice](https://www.libreoffice.org/) | `soffice --headless --convert-to csv a.xlsx` | `14.103s/17.388s/0.644s` |
