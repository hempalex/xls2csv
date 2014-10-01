#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import csv
import xlrd
from optparse import OptionParser

def xls2csv(infilepath, outfile, sheetid=1, delimiter=",", sheetdelimiter="--------", encoding="cp1251"):
    writer = csv.writer(outfile, dialect='excel', quoting=csv.QUOTE_ALL, delimiter=delimiter)

    book = xlrd.open_workbook(infilepath, encoding_override=encoding, formatting_info=True)

    formats = {}
    for i, f in book.format_map.iteritems():
        if f.format_str != None:
            formats[i] = extract_number_format(f.format_str)

    if sheetid > 0:
        # xlrd has zero-based sheet enumeration, but 0 means "convert all"
        sheet_to_csv(book, sheetid - 1, writer, formats)
    else:
        for sheetid in xrange(book.nsheets):
            sheet_to_csv(book, sheetid, writer, formats)
            if sheetdelimiter and sheetid < book.nsheets - 1:
                outfile.write(sheetdelimiter + "\r\n")


def sheet_to_csv(book, sheetid, writer, formats):

    sheet = book.sheet_by_index(sheetid)

    if not sheet:
        raise Exception("Sheet %i Not Found" % sheetid)

    for i in xrange(sheet.nrows):
        row = [""] * sheet.ncols
        cells = sheet.row(i)
        for j in xrange(sheet.ncols):

            cell = cells[j]

            if cell.ctype == xlrd.XL_CELL_TEXT:
                cval = cell.value.encode('utf-8')

            elif cell.ctype == xlrd.XL_CELL_NUMBER:

                if cell.xf_index != None:

                    a_fmt = formats[book.xf_list[cell.xf_index].format_key]
                    if a_fmt:
                        cval = format_number(cell.value, a_fmt, "", ".")
                    elif cell.value == int(cell.value):
                        cval = int(cell.value)
                    else:
                        cval = "%s" % cell.value

            elif cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    cval = xlrd.xldate_as_tuple(cell.value, book.datemode)
                except xlrd.XLDateError:
                    e1, e2 = sys.exc_info()[:2]
                    cval = "%s:%s" % (e1.__name__, e2)

            else:  # XL_CELL_EMPTY, XL_CELL_ERROR, XL_CELL_BLANK
                cval = ""

            row[j] = cval

        writer.writerow(row)


#xlrd format excel number
#http://uucode.com/blog/2013/10/22/using-xlrd-and-formatting-excel-numbers/

import re
re_maybe_numfmt = re.compile('[0#.,]*[0#][0#.,]*')


def extract_number_format(s_fmt):

    # If don't know what does the format "Standard/GENERAL" mean.
    # As far as I understand, the presentation can differ depending
    # on the locale and user settings. Here is a my proposal.

    if None == s_fmt or 'GENERAL' == s_fmt:
        return (None, '#', '#')

    # Find the number-part
    m = re_maybe_numfmt.search(s_fmt)
    if m is None:
        return None  # return
    s_numfmt = str(m.group(0))

    # Only one comma
    pos_comma = s_numfmt.find(',')
    if pos_comma > -1:
        pos2 = s_numfmt.find(',', pos_comma + 1)
        if pos2 > -1:
            return None  # return

  # Only one dot

    pos_dot = s_numfmt.find('.')
    if pos_dot > -1:
        pos2 = s_numfmt.find('.', pos_dot + 1)
        if pos2 > -1:
            return None  # return

  # Exactly three positions between comma and dot

    if pos_comma > -1:
        pos2 = (pos_dot if pos_dot > -1 else len(s_numfmt))
        if pos2 - pos_comma != 4:
            return None  # return

  # Create parts

    (part_above1000, part_below1000, part_below1) = (None, None, None)
    if pos_dot > -1:
        part_below1 = s_numfmt[pos_dot + 1:]
        s_numfmt = s_numfmt[:pos_dot]
    if pos_comma > -1:
        part_above1000 = s_numfmt[:pos_comma]
        part_below1000 = s_numfmt[pos_comma + 1:]
    else:
        part_below1000 = s_numfmt
    return (part_above1000, part_below1000, part_below1)


def format_number(f, a_fmt, div1000, div1):
    (part_above1000, part_below1000, part_below1) = a_fmt
    s_fmt = '%'
    if f < 0:
        is_negative = 1
        f = abs(f)
    else:
        is_negative = 0

    # Float to string with a minimal precision after comma.
    # Filling the integer part with '0' at left doesn't work for %f.

    precision = (len(part_below1) if part_below1 else 0)
    s_fmt = '%.' + str(precision) + 'f'
    s_f = s_fmt % f

    # Postprocessing. Drop insignificant zeros.
    while precision:
        if '0' == part_below1[precision - 1]:
            break
        if '0' != s_f[-1]:
            break
        s_f = s_f[:-1]
        precision = precision - 1

    if '.' == s_f[-1]:
        s_f = s_f[:-1]
        precision = 0

    # Add significant zeros
    part_above1 = (part_above1000 + part_below1000 if part_above1000 else part_below1000)
    i = part_above1.find('0')
    if i > -1:
        if precision:
            need_len = len(part_above1) - i + 1 + precision
        else:
            need_len = len(part_above1) - i
        if need_len > len(s_f):
            s_f = s_f.rjust(need_len, '0')

    # Put dots and commas
    if '.' != div1:
        s_f = s_f.replace('.', div1)
    if part_above1000:
        if precision:
            div_pos = len(s_f) - precision - 4
        else:
            div_pos = len(s_f) - 3
        if div_pos > 0:
            s_f = s_f[:div_pos] + div1000 + s_f[div_pos:]

    # Add negative sign
    if is_negative:
        if '0' == s_f[0]:
            s_f = '-' + s_f[1:]
        else:
            s_f = '-' + s_f

    return s_f






if __name__ == "__main__":
    parser = OptionParser(usage="%prog [options] infile [outfile]", version="0.1")
    parser.add_option("-s", "--sheet", dest="sheetid", default=1, type="int", help="sheet no to convert (0 for all sheets)")
    parser.add_option("-d", "--delimiter", dest="delimiter", default=",", help="delimiter - csv columns delimiter, 'tab' or 'x09' for tab (comma is default)")
    parser.add_option("-p", "--sheetdelimiter", dest="sheetdelimiter", default="--------", help="sheets delimiter used to separate sheets, pass '' if you don't want delimiters (default '--------')")
    parser.add_option("-e", "--encoding", dest="encoding", default="cp1251", help="xls file encoding if the CODEPAGE record is missing")

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
        'encoding': options.encoding,
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
