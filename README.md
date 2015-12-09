#cxlsx_to_csv
**cxlsx_to_csv is a simple converter of Excel 2007 files (a.k.a. Open Office XML) to .CSV**

###FEATURES:
* Simple converter, pretty fast as done in C.
* Only depends on [miniz](https://code.google.com/p/miniz/) (included for convenience) and one XML library, that can be either [expat](http://expat.sourceforge.net/) or [Parsifal](http://www.saunalahti.fi/~samiuus/toni/xmlproc/) or [Mini-XML](http://www.msweet.org/projects.php?Z3).

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
* Tested under Ubuntu 15.10 on an Intel i3-3217U CPU @ 1.80GHz, with a Crucial CT120M500 SSD.
* Locale set to LC_ALL=C.UTF-8
* The input spreadsheet `a.xlsx` has 48665 rows and 46 columns, and weights 9165KB.

| Tool | Command | Time (real/user/sys)|
|:------------ |:------------|:--|
| [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv) ([expat](http://expat.sourceforge.net/))| `cxlsx_to_csv -if a.xlsx -of a.csv` | `3.020s/2.960s/0.064s` |
| [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv) ([Parsifal](http://www.saunalahti.fi/~samiuus/toni/xmlproc/))| `cxlsx_to_csv -if a.xlsx -of a.csv` | `3.957s/3.884s/0.072s` |
| [cxlsx_to_csv](https://github.com/vpaesa/cxlsx_to_csv) ([Mini-XML](http://www.msweet.org/projects.php?Z3))| `cxlsx_to_csv -if a.xlsx -of a.csv` | `4.967s/4.880s/0.088s` |
| [Gnumeric](http://www.gnumeric.org/) | `ssconvert --export-type=Gnumeric_stf:stf_assistant -O 'eol=windows separator=, format=raw transliterate-mode=escape quoting-mode=auto' a.xlsx a.csv` | `11.869s/11.692s/0.184s` |
| [LibreOffice 5](https://www.libreoffice.org/) | `soffice --headless --convert-to csv a.xlsx` | `12.779s/15.292s/0.536s` |
| [xlsx2csv](https://github.com/dilshod/xlsx2csv) | `xlsx2csv.py a.xlsx a.csv` | `39.925s/39.684s/0.148s` |

