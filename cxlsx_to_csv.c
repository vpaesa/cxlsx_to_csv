/*****************************************************************
 NAME:
   cxlsx_to_csv - convert Excel 2007 files to .CSV

 USAGE:
   cxlsx_to_csv -if input.xlsx [-sh sheet_id] [-of output.csv]
  
 COMPILATION:
   cc -o cxlsx_to_csv cxlsx_to_csv.c -l expat
   
 Must be used with Expat compiled for UTF-8 output.

** Copyright (C) 2015 Victor Paesa
**
** This program is free software; you can redistribute it and/or modify
** it under the terms of the GNU General Public License as published by
** the Free Software Foundation; either version 2 of the License, or
** (at your option) any later version.
**
** This program is distributed in the hope that it will be useful,
** but WITHOUT ANY WARRANTY; without even the implied warranty of
** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
** GNU General Public License for more details.
**
** You should have received a copy of the GNU General Public License
** along with this program; if not, write to the Free Software
** Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

*****************************************************************/

#include <ctype.h>
#include <stdlib.h>
#include "miniz.c"

typedef unsigned char uint8;
typedef unsigned short uint16;
typedef unsigned int uint;

#include <stdio.h>
#include <expat.h>
#include <string.h>

#if defined(__amigaos__) && defined(__USE_INLINE__)
#include <proto/expat.h>
#endif

#ifdef XML_LARGE_SIZE
#if defined(XML_USE_MSC_EXTENSIONS) && _MSC_VER < 1400
#define XML_FMT_INT_MOD "I64"
#else
#define XML_FMT_INT_MOD "ll"
#endif
#else
#define XML_FMT_INT_MOD "l"
#endif

static char *usage_str = "\n\
NAME:\n\
cxlsx_to_csv - convert Excel 2007 files to .CSV\n\
\n\
SYNOPSIS:\n\
cxlsx_to_csv -if input.xlsx [-sh sheet_id] [-of output.csv]\n\
    input.xlsx        input spreadsheet in Excel 2007 format (Office Open XML)\n\
    sheet_id        name of the sheet within the workbook (default is first one)\n\
    output.csv        output CSV file (default is STDOUT)\n\
\n\
CAVEATS:\n\
Separator in output CSV is comma.\n\
";

// https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3&usg=AFQjCNHniIQ4KTIFQZ6efVfpDtETwU9Cmw
// Total number of characters that an Excel cell can contain: 32,767
#define BUFFSIZE 40960

/*
** An object used to parse XML content of XLSX
*/
typedef struct XLSXCtx XLSXCtx;
struct XLSXCtx {
  FILE  *outf;
  int    xml_depth;      /* Current dept while parsing the XML tree */
  char **shr_str;
  int    shr_str_num, shr_str_cnt;
  char  *shr_tv_val;
  int    shr_si, shr_tv;
  int    sheet_num_rows, sheet_num_cols;
  int    current_row, current_col, expected_col;
  int    lookup_v;
  char   shr_buff[BUFFSIZE];
};

/*  
    XLSX files are zip files which contain several xml files with data:
 _rels/.rels
 docProps/app.xml
 docProps/core.xml
 xl/_rels/workbook.xml.rels
 xl/sharedStrings.xml
 xl/worksheets/sheet2.xml
 xl/worksheets/sheet1.xml
 xl/styles.xml
 xl/workbook.xml
 [Content_Types].xml

  The name of each sheet is in xl/workbook.xml
  The individual sheets are kept in xl/worksheets/sheet1.xml
  To save on space, Microsoft stores all the character literal values in one common xl/sharedStrings.xml dictionary file. The individual cell value found for this string in the actual sheet1.xml file is just an index into this dictionary.
  Dates are stored as day number since 1900/01/01 (at least they are supposed to. I discovered that one has to subtract 2 from this number of days to get the correct conversion).
  Time portion of the date is stored as a fraction of a day, so it has to be multiplied by 60*60*24 (86400) to get the actual number os seconds.
  Microsoft does not store empty cells or rows in xl/worksheets/sheet1.xml, so any gaps between values have to be taken care by the code.
  To figure out the number of skipped columns one need to be able to figure out the distance between, say, cell "AB67" and "C67". The way columns are named: A through Z, then AA through AZ, then AAA through AAZ, etc., suggests that we may assume they are using a base-26 system and therefore use a simple conversion method from a base-26 to the decimal system and then use subtraction to find out the number of commas between columns.
    
-------------------------------------------------------------------------
xl/sharedStrings.xml has in "sst:uniqueCount" a count of the number of unique strings
unzip -c 3x2.xlsx xl/sharedStrings.xml | tidy -xml
-------------------------------------------------------------------------
<sst count="9" uniqueCount="5">
  <si>
    <t>Col1</t>
  </si>
  <si>
    <t>Col2</t>
  </si>
  <si>
    <t>Col3</t>
  </si>
  <si>
    <t>a</t>
  </si>
  <si>
    <t>b</t>
  </si>
  <si>
    <t>c</t>
  </si>
</sst>
-------------------------------------------------------------------------

-------------------------------------------------------------------------
xl/worksheets/sheet1.xml has in "dimension:ref" the enclosing range of cells used
unzip -c 3x2.xlsx xl/worksheets/sheet1.xml | tidy -indent -xml
-------------------------------------------------------------------------
<worksheet>
  <dimension ref="A1:C3" />
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1" t="s">
        <v>1</v>
      </c>
      <c r="C1" t="s">
        <v>27.344</v>
      </c>
    </row>
    <row r="2" spans="1:3">
      <c r="A2" t="s">
        <v>3</v>
      </c>
      <c r="B2" s="1">
        <v>42283</v>
      </c>
      <c r="C2" t="s">
        <v>3</v>
      </c>
    </row>
    <row r="3" spans="1:3">
      <c r="A3" t="s">
        <v>4</v>
      </c>
      <c r="B3" t="s">
        <v>4</v>
      </c>
      <c r="C3" t="s">
        <v>4</v>
      </c>
    </row>
  </sheetData>
</worksheet>
-------------------------------------------------------------------------

-------------------------------------------------------------------------
xl/styles.xml
-------------------------------------------------------------------------
<styleSheet>
  <numFmts count="10">
    <numFmt numFmtId="164" formatCode="GENERAL" />
    <numFmt numFmtId="165" formatCode="DD/MM/YY" />
    <numFmt numFmtId="166" formatCode="DD/MM/YYYY" />
    <numFmt numFmtId="167" formatCode="D&quot; de &quot;MMM&quot; de &quot;YY" />
    <numFmt numFmtId="168" formatCode="D&quot; de &quot;MMM&quot; de &quot;YYYY" />
    <numFmt numFmtId="169" formatCode="DDDD, MMMM\ DD&quot;, &quot;YYYY" />
    <numFmt numFmtId="170" formatCode="YYYY\-MM\-DD" />
    <numFmt numFmtId="171" formatCode="YYYY\-MM\-DD" />
    <numFmt numFmtId="172" formatCode="DDDD, D\ MMMM\ YYYY" />
    <numFmt numFmtId="173" formatCode="YYYY\-MM\-DD\ HH:MM:SS.SSS" />
  </numFmts>
</styleSheet>
-------------------------------------------------------------------------



*/

/* CSV escaping:
  If a cell value contains a comma or a line feed, the entire value has to be enclosed in doublequotes.
  If a cell value contains a doublequote each of them has to be doubled and then the value should be enclosed in doublequotes. 
*/

/*
** CSV code from sqlite
*/

/*
** If a field contains any character identified by a 1 in the following
** array, then the string must be quoted for CSV.
*/
static const char needCsvQuote[] = {
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 0, 1, 0, 0, 0, 0, 1,   0, 0, 0, 0, 0, 0, 0, 0, 
  0, 0, 0, 0, 0, 0, 0, 0,   0, 0, 0, 0, 0, 0, 0, 0, 
  0, 0, 0, 0, 0, 0, 0, 0,   0, 0, 0, 0, 0, 0, 0, 0, 
  0, 0, 0, 0, 0, 0, 0, 0,   0, 0, 0, 0, 0, 0, 0, 0, 
  0, 0, 0, 0, 0, 0, 0, 0,   0, 0, 0, 0, 0, 0, 0, 0, 
  0, 0, 0, 0, 0, 0, 0, 0,   0, 0, 0, 0, 0, 0, 0, 1, 
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
  1, 1, 1, 1, 1, 1, 1, 1,   1, 1, 1, 1, 1, 1, 1, 1,   
};

/*
** Output a single term of CSV.  Actually, colSeparator is used for
** the separator, which may or may not be a comma.  "" is
** the null value.  Strings are quoted if necessary.  The separator
** is only issued if bSep is true.
*/
static void output_csv(FILE *out, const char colSeparator, const char *z, int bSep)
{
  if (z==0) {
    //fprintf(out,"%s","");
  } else{
    int i;
    for(i=0; z[i]; i++){
      if (needCsvQuote[((unsigned char*)z)[i]] || (z[i]==colSeparator)) {
        i = 0;
        break;
      }
    }
    if (i==0) {
      putc('"', out);
      for (i=0; z[i]; i++) {
        if (z[i]=='"')
          putc('"', out);
        putc(z[i], out);
      }
      putc('"', out);
    } else {
      fprintf(out, "%s", z);
    }
  }
  if (bSep) {
    putc(colSeparator, out);
  }
}

void excelcolrow(char *string, int *outcol, int *outrow)
{
  int i, col, base;

  col = 0;
  base = 1;
  for (i = 0; i < strlen(string); i++) {
    if (isalpha(string[i])) {
      col = col * 26 + ((toupper(string[i])) - 'A' + 1);
    }
    else
      break;
  }
  *outcol = col;
  *outrow = atoi(string + i);
  return;
}

void rangecolrow(char *string, int *outcol, int *outrow)
{
  int col, row;
  char *coloninstr;

  coloninstr = strchr(string, ':');
  if (coloninstr) {
    string = coloninstr + 1;
    //fprintf(stderr, "rangecolrow: %s \n", string);
  }
  excelcolrow(string, &col, &row);
  *outcol = col;
  *outrow = row;
}

static void XMLCALL StartSharedStrings(void *data, const char *el, const char **attr)
{
  int i;
  XLSXCtx *ctx = data;
  
  if ((ctx->xml_depth == 0) && (!strcmp(el, "sst"))) {
    for (i = 0; attr[i]; i += 2) {
      if (!strcmp(attr[i], "uniqueCount")) {
        //printf(" %s='%s'\n", attr[i], attr[i + 1]);
        ctx->shr_str_cnt = atoi(attr[i + 1]);
        ctx->shr_str = malloc(sizeof(char *) * ctx->shr_str_cnt);
      }
    }
  }
  if ((ctx->xml_depth == 2) && (!strcmp(el, "t"))) {
    ctx->shr_tv = 1;
    ctx->shr_tv_val = ctx->shr_buff;
    *(ctx->shr_tv_val) = 0;
  }
  ctx->xml_depth++;
}

static void XMLCALL EndSharedStrings(void *data, const char *el)
{
  XLSXCtx *ctx = data;

  ctx->xml_depth--;
  if ((ctx->xml_depth == 2) && (!strcmp(el, "t"))) {
    ctx->shr_tv = 0;
    ctx->shr_str[ctx->shr_str_num] = strdup(ctx->shr_buff);
    ctx->shr_str_num++;
  }
}

static void XMLCALL ChrHndlr(void *data, const char *s, int len)
{
  char *src;
  XLSXCtx *ctx = data;

  if (ctx->shr_tv) {
    src = (char *) s;
    while (len) {
      *(ctx->shr_tv_val)++ = *src++;
      len--;
    }
    *(ctx->shr_tv_val) = 0;
  }
}

static void XMLCALL StartSheet(void *data, const char *el, const char **attr)
{
  int i, j;
  XLSXCtx *ctx = data;

  if ((ctx->xml_depth == 1) && (!strcmp(el, "dimension"))) {
    for (i = 0; attr[i]; i += 2) {
      if (!strcmp(attr[i], "ref")) {
        //fprintf(stderr, "dimension %s='%s'\n", attr[i], attr[i + 1]);
        rangecolrow((char *) attr[i + 1], &(ctx->sheet_num_cols), &(ctx->sheet_num_rows));
        //fprintf(stderr, "cols: %d  rows: %d\n", ctx->sheet_num_cols, ctx->sheet_num_rows);
      }
    }
  }
  if ((ctx->xml_depth == 2) && (!strcmp(el, "row"))) {
    for (i = 0; attr[i]; i += 2) {
      if (!strcmp(attr[i], "r")) {
        //fprintf(stderr, "row %s='%s'\n", attr[i], attr[i + 1]);
        ctx->expected_col = 1;
      }
    }
  }
  if ((ctx->xml_depth == 3) && (!strcmp(el, "c"))) {
    ctx->lookup_v = 0;
    for (i = 0; attr[i]; i += 2) {
      if (!strcmp(attr[i], "r")) {
        //fprintf(stderr, "c %s='%s'\n", attr[i], attr[i + 1]);
        excelcolrow((char *) attr[i + 1], &(ctx->current_col), &(ctx->current_row));
        for (j = ctx->expected_col; (j<ctx->current_col)&&(j<ctx->sheet_num_cols); j++)
          putc(',', ctx->outf);
        ctx->expected_col = ctx->current_col+1;
      }
      else if (!strcmp(attr[i], "t")) {
        //printf("c %s='%s'\n", attr[i], attr[i + 1]);
        if (*attr[i + 1] == 's') {
          ctx->lookup_v = -1;
        }
        //fprintf(stderr, "cols: %d  rows: %d\n", ctx->num_col, ctx->num_row);
      }
    }
  }
  if ((ctx->xml_depth == 4) && (!strcmp(el, "v"))) {
    ctx->shr_tv = 1;
    ctx->shr_tv_val = ctx->shr_buff;
    *(ctx->shr_tv_val) = 0;
  }
  ctx->xml_depth++;
}

static void XMLCALL EndSheet(void *data, const char *el)
{
  int j;
  XLSXCtx *ctx = data;

  ctx->xml_depth--;
  if ((ctx->xml_depth == 4) && (!strcmp(el, "v"))) {
    ctx->shr_tv = 0;
    if (ctx->lookup_v) {
      //printf("v %s\n", ctx->shr_str[atoi(ctx->shr_buff)]);
      output_csv(ctx->outf, ',', ctx->shr_str[atoi(ctx->shr_buff)], (ctx->current_col < ctx->sheet_num_cols));
    }
    else {
      //printf("v %s\n", ctx->shr_buff);
      output_csv(ctx->outf, ',', ctx->shr_buff, (ctx->current_col < ctx->sheet_num_cols));
    }
  }
  if ((ctx->xml_depth == 2) && (!strcmp(el, "row"))) {
    for (j = ctx->expected_col; j<ctx->sheet_num_cols; j++)
      putc(',', ctx->outf);
    fprintf(ctx->outf, "\r\x0A");
    // TODO: Check if \r\x0A portable between Windows & UNIX
  }
}

int main(int argc, char *argv[])
{
  int i;
  size_t sheet_size;
  void *sheet_ptr;
  XML_Parser p;
  XLSXCtx parse_ctx;
  char sheetname[64];
  
  int opt_if = 0;
  int opt_sh = 0;
  int opt_of = 0;

  for (i=1; i<argc; i++) {
    if (i==opt_if)
      continue;
    if (!strcmp("-if", argv[i]))
      if ((i+1) < argc)
        opt_if = i+1;
      else {
        fputs("'-if' needs an Excel file name for input\n", stderr);
        fputs(usage_str, stderr);
        return 1;
      }
    if (i==opt_sh)
      continue;
    if (!strcmp("-sh", argv[i]))
      if ((i+1) < argc)
        opt_sh = i+1;
      else {
        fputs("'-sh' needs a sheet number\n", stderr);
        fputs(usage_str, stderr);
        return 1;
      }
    if (i==opt_of)
      continue;
    if (!strcmp("-of", argv[i]))
      if ((i+1) < argc)
        opt_of = i+1;
      else {
        fputs("'-of' needs an CSV file name for output\n", stderr);
        fputs(usage_str, stderr);
        return 1;
      }
  }

  if (!opt_if) {
    fputs("Missing '-if input.xlsx'\n", stderr);
    fputs(usage_str, stderr);
    return 1;
  }
  if (!opt_sh) {
    //fputs("Missing '-sh sheetnum', hence assuming first sheet.\n", stderr);
    opt_sh = 1;
  }
  else {
    opt_sh = atoi(argv[opt_sh]);
    // TODO: Check sheet number among existing sheets. Accept sheet names.
    if (!opt_sh)
      opt_sh = 1;
  }
  if (!opt_of) {
    //fputs("Missing '-of output.csv', hence assuming STDOUT.\n", stderr);
    parse_ctx.outf = stdout; 
  }
  else {
    parse_ctx.outf = fopen(argv[opt_of], "w");
    if (!parse_ctx.outf) {
      fprintf(stderr, "Couldn't open output file '%s' .\n", argv[opt_of]);
      exit(-1);
    }
  }

  // Process xl/sharedStrings.xml and load them into shr_str[]
  sheet_ptr = mz_zip_extract_archive_file_to_heap(argv[opt_if], "xl/sharedStrings.xml", &sheet_size, MZ_ZIP_FLAG_CASE_SENSITIVE);
  if (sheet_ptr) {
    parse_ctx.xml_depth = 0;
    p = XML_ParserCreate(NULL);
    if (!p) {
      fprintf(stderr, "Couldn't allocate memory for parser\n");
      exit(-1);
    }
    XML_SetUserData(p, &parse_ctx);
    XML_SetElementHandler(p, StartSharedStrings, EndSharedStrings);
    XML_SetCharacterDataHandler(p, ChrHndlr);
    if (XML_Parse(p, sheet_ptr, sheet_size, -1) == XML_STATUS_ERROR) {
      fprintf(stderr, "Parse error at line %" XML_FMT_INT_MOD "u:\n%s\n",
               XML_GetCurrentLineNumber(p),
               XML_ErrorString(XML_GetErrorCode(p)));
      exit(-1);
    }
    XML_ParserFree(p);
    //for (i = 0; i < ctx->shr_str_cnt; i++)
    //  printf("%s\n", ctx->shr_str[i]);
  }
  else {
    //fprintf(stderr, "Warning: could not read xl/sharedStrings.xml\n");
    // TODO: Only warn about missing xl/sharedStrings.xml is it referenced by some t="s"
  }

  // Process xl/worksheets/sheet1.xml and load them into sheet_tbl[,]
  sprintf(sheetname, "xl/worksheets/sheet%d.xml", opt_sh);
  sheet_ptr = mz_zip_extract_archive_file_to_heap(argv[opt_if], sheetname, &sheet_size, MZ_ZIP_FLAG_CASE_SENSITIVE);
  if (sheet_ptr) {
    parse_ctx.xml_depth = 0;
    parse_ctx.shr_tv = 0;
    p = XML_ParserCreate(NULL);
    if (!p) {
      fprintf(stderr, "Couldn't allocate memory for parser\n");
      exit(-1);
    }
    XML_SetUserData(p, &parse_ctx);
    XML_SetElementHandler(p, StartSheet, EndSheet);
    XML_SetCharacterDataHandler(p, ChrHndlr);
    if (XML_Parse(p, sheet_ptr, sheet_size, -1) == XML_STATUS_ERROR) {
      fprintf(stderr, "Parse error at line %" XML_FMT_INT_MOD "u:\n%s\n",
               XML_GetCurrentLineNumber(p),
               XML_ErrorString(XML_GetErrorCode(p)));
      exit(-1);
    }
    XML_ParserFree(p);
  }
  else {
    fprintf(stderr, "Error: could not read sheet number %d.\n", opt_sh);
    exit(-1);
  }

  return 0;
}
