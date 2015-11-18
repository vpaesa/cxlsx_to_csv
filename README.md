#cxlsx_to_csv
**cxlsx_to_csv is a simple converter of Excel 2007 files (a.k.a. Open Office XML) to .CSV**

###FEATURES:
* Simple converter, pretty fast as done in C.
* Only depends on [miniz](https://code.google.com/p/miniz/) (included for convenience) and one XML library, that can be either [expat](http://expat.sourceforge.net/) or [Mini-XML](http://www.msweet.org/projects.php?Z3).

The .XLSX format is just a glorified .ZIP (that I open thanks to miniz), containing a set of .XML files (that I parse thanks to Expat or Mini-XML).
Notice that Excel stores dates as the number of days that have elapsed since 1-January-1900 (the Excel Epoch), and this program exports dates simply as the floating point value they are stored.

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
* The input spreadsheet `a.xlsx` has 47544 rows and 46 columns, and weights 8962KB.

| Tool | Command | Time (real/user/sys)|
|:------------ |:------------|:--|
| [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv) ([expat](http://expat.sourceforge.net/))| `cxlsx_to_csv -if a.xlsx -of a.csv` | `2.886s/2.784s/0.100s` |
| [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv) ([Mini-XML](http://www.msweet.org/projects.php?Z3))| `cxlsx_to_csv -if a.xlsx -of a.csv` | `4.802s/4.732s/0.072s` |
| [Gnumeric](http://www.gnumeric.org/) | `ssconvert --export-type=Gnumeric_stf:stf_assistant -O 'eol=windows separator=, format=raw transliterate-mode=escape quoting-mode=auto' a.xlsx a.csv` | `11.735s/11.544s/0.196s` |
| [LibreOffice](https://www.libreoffice.org/) | `soffice --headless --convert-to csv a.xlsx` | `14.103s/17.388s/0.644s` |
| [xlsx2csv](https://github.com/dilshod/xlsx2csv) | `xlsx2csv.py a.xlsx a.csv` | `38.944s/38.616s/0.220s` |


