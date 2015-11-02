# cxlsx_to_csv
**cxlsx_to_csv is a simple converter of Excel 2007 files (a.k.a. Open Office XML) to .CSV**

FEATURES:
* Simple converter, pretty fast as done in C.
* Only depends on [miniz](https://code.google.com/p/miniz/) (included for convenience) and [expat](http://expat.sourceforge.net/).

The .XLSX format is just a glorified .ZIP (that I open thanks to miniz), containing a set of .XML files (that I parse thanks to Expat).
Dates are exported as the number of days that have elapsed since 1-January-1900 (the Excel Epoch).

SYNOPSIS:
```
cxlsx_to_csv -if input.xlsx [-sh sheet_id] [-of output.csv]
    input.xlsx  input spreadsheet in Excel 2007 format (Office Open XML)
    sheet_id    number of the sheet within the workbook (default is first one)
    output.csv  output CSV file (default is STDOUT)
```
