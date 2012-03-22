#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import csv
import xlrd
from optparse import OptionParser


def xls2csv(infilepath, outfile, sheetid=1, delimiter=",", sheetdelimiter="--------"):
    writer = csv.writer(outfile, quoting=csv.QUOTE_MINIMAL, delimiter=delimiter)

    book = xlrd.open_workbook(infilepath)

    if sheetid > 0:
        # xlrd has zero-based sheet enumeration, but 0 means "convert all"
        sheet_to_csv(book, sheetid - 1, writer)
    else:
        for sheetid in xrange(book.nsheets):
            sheet_to_csv(book, sheetid, writer)
            if sheetdelimiter and sheetid < book.nsheets - 1:
                outfile.write(sheetdelimiter + "\r\n")


def sheet_to_csv(book, sheetid, writer):

    sheet = book.sheet_by_index(sheetid)

    if not sheet:
        raise Exception("Sheet %i Not Found" % sheetid)

    for i in xrange(sheet.nrows):
        row = [""] * sheet.ncols
        ctys = sheet.row_types(i)
        cvals = sheet.row_values(i)
        for j in xrange(sheet.ncols):
            cty = ctys[j]
            cval = cvals[j]
            if cty == xlrd.XL_CELL_DATE:
                try:
                    row[j] = xlrd.xldate_as_tuple(cval, book.datemode)
                except xlrd.XLDateError:
                    e1, e2 = sys.exc_info()[:2]
                    row[j] = "%s:%s" % (e1.__name__, e2)
                    cty = xlrd.XL_CELL_ERROR
            elif cty == xlrd.XL_CELL_ERROR:
                row[j] = xlrd.error_text_from_code.get(cval, '<unknown error="" code="" 0x%02x="">' % cval)
            else:
                row[j] = ('%s' % cval).encode('utf-8')
        writer.writerow(row)


if __name__ == "__main__":
    parser = OptionParser(usage="%prog [options] infile [outfile]", version="0.1")
    parser.add_option("-s", "--sheet", dest="sheetid", default=1, type="int",
      help="sheet no to convert (0 for all sheets)")
    parser.add_option("-d", "--delimiter", dest="delimiter", default=",",
      help="delimiter - csv columns delimiter, 'tab' or 'x09' for tab (comma is default)")
    parser.add_option("-p", "--sheetdelimiter", dest="sheetdelimiter", default="--------",
      help="sheets delimiter used to separate sheets, pass '' if you don't want delimiters (default '--------')")

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
      'delimiter': delimiter,
      'sheetdelimiter': options.sheetdelimiter,
    }

    if len(args) < 1:
        parser.print_help()
    else:
        if len(args) > 1:
            outfile = open(args[1], 'w+')
            xls2csv(args[0], outfile, **kwargs)
            outfile.close()
        else:
            xls2csv(args[0], sys.stdout, **kwargs)
