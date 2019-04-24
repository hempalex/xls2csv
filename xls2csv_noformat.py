#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import csv
import xlrd
from optparse import OptionParser

def xls2csv(infilepath, ofile, sheetid=1, delim=",", sheetdelimiter="--------", encoding="utf-8"):

    writer = csv.writer(ofile, dialect='excel', quoting=csv.QUOTE_ALL, delimiter=delim)

    book = xlrd.open_workbook(infilepath, encoding_override=encoding)

    if sheetid > 0:
        # xlrd has zero-based sheet enumeration, but 0 means "convert all"
        sheet_to_csv(book, sheetid - 1, writer)
    else:
        for sheet_n in range(book.nsheets):
            sheet_to_csv(book, sheet_n, writer)
            if sheetdelimiter and sheet_n < book.nsheets - 1:
                ofile.write(sheetdelimiter + "\r\n")


def sheet_to_csv(book, sheetid, writer):

    sheet = book.sheet_by_index(sheetid)

    if not sheet:
        raise Exception("Sheet %i Not Found" % sheetid)

    for i in range(sheet.nrows):

        row = []

        ctys = sheet.row_types(i)
        cvals = sheet.row_values(i)

        for j in range(sheet.row_len(i)):

            cty = ctys[j]
            cval = cvals[j]

            if cty == xlrd.XL_CELL_NUMBER:

                if cval == int(cval):
                    cval = int(cval)
                else:
                    cval = str(cval)

            elif cty == xlrd.XL_CELL_DATE:
                try:
                    cval = xlrd.xldate_as_tuple(cval, book.datemode)
                except xlrd.XLDateError:
                    e1, e2 = sys.exc_info()[:2]
                    cval = "%s:%s" % (e1.__name__, e2)

            elif cty != xlrd.XL_CELL_TEXT:  # XL_CELL_EMPTY, XL_CELL_ERROR, XL_CELL_BLANK
                cval = ""

            row.append(cval)

        writer.writerow(row)


if __name__ == "__main__":
    parser = OptionParser(usage="%prog [options] infile [outfile]", version="0.1")
    parser.add_option("-s", "--sheet", dest="sheetid", default=1, type="int", help="sheet no to convert (0 for all sheets)")
    parser.add_option("-d", "--delimiter", dest="delimiter", default=",", help="delimiter - csv columns delimiter, 'tab' or 'x09' for tab (comma is default)")
    parser.add_option("-p", "--sheetdelimiter", dest="sheetdelimiter", default="--------", help="sheets delimiter used to separate sheets, pass '' if you don't want delimiters (default '--------')")
    parser.add_option("-e", "--encoding", dest="encoding", default="utf-8", help="xls file encoding if the CODEPAGE record is missing")

    (options, args) = parser.parse_args()

    if len(options.delimiter) == 1:
        delimiter = options.delimiter
    elif options.delimiter == 'tab':
        delimiter = '\t'
    elif options.delimiter == 'comma':
        delimiter = ','
    elif options.delimiter[0] == 'x':
        delimiter = chr(int(options.delimiter[1:]))
    else:
        raise Exception("Invalid delimiter")

    kwargs = {
        'sheetid': options.sheetid,
        'delim': delimiter,
        'sheetdelimiter': options.sheetdelimiter,
        'encoding': options.encoding,
    }

    if len(args) < 1:
        parser.print_help()
    else:
        if len(args) > 1:
            if sys.version_info[0] == 2:
                outfile = open(args[1], 'wb+')
            elif sys.version_info[0] == 3:
                outfile = open(args[1], 'w+', encoding="utf-8", newline="")
            else:
                sys.stderr.write("error: version of your python is not supported: " + str(sys.version_info) + "\n")
                sys.exit(1)


            xls2csv(args[0], outfile, **kwargs)
            outfile.close()
        else:
            xls2csv(args[0], sys.stdout, **kwargs)
