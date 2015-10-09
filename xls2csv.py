#!/usr/bin/python
# coding: utf-8

from xlrd import open_workbook, cellname

import collections
import csv
import os.path
import re
import sys


def get_obcine_column(sheet):
    obcine_counter = collections.Counter()
    min_row = 1000
    max_row = 0
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            val = sheet.cell(row_index, col_index).value
            if isinstance(val, unicode):
                if val.strip().startswith(u'OBČINA ') or val.strip().startswith(u'AJDOVŠČINA') or val.strip().startswith(u'ŽUŽEMBERK'):
                    obcine_counter[col_index] += 1
                    min_row = min(min_row, row_index)
                    max_row = max(max_row, row_index)

    obcine_column = obcine_counter.most_common(1)[0][0]
    return obcine_column, min_row, max_row


def get_legenda_row(sheet, data_start):
    legenda_ctr = collections.Counter()
    min_col = 1000
    max_col = 0
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            val = sheet.cell(row_index, col_index).value

            if row_index < data_start:
                if isinstance(val, float) or (isinstance(val, unicode) and re.match('^[IXV\.]+$', val.strip())):
                    legenda_ctr[row_index] += 1
                    min_col = min(min_col, col_index)
                    max_col = max(max_col, col_index)

    print legenda_ctr.most_common()

    obcine_column = legenda_ctr.most_common(1)[0][0]
    return obcine_column, min_col, max_col


def get_obcine(sheet, col, ystart, yend):
    obcine = {}
    for row in xrange(ystart, yend + 1):
        cell = sheet.cell(row, col)
        obcine[row] = cell.value
    return obcine


def get_postavke(sheet, row, xstart, xend):
    postavke = {}
    for col in xrange(xstart, xend + 1):
        cell = sheet.cell(row, col)
        postavke[col] = cell.value
    return postavke


def get_obcine_lookup():
    f = open('obcine_lookup.csv')
    rdr = csv.reader(f)
    obcine_lookup = dict([(r[0].decode('utf-8'), r[1].decode('utf-8')) for r in rdr])
    return obcine_lookup


def do_sheets(fn):
    obcine_lookup = get_obcine_lookup()
    book = open_workbook(fn)
    for n in xrange(book.nsheets - 1):
        sheet = book.sheet_by_index(n)

        print 'SHEET NAME', sheet.name
        print 'SHEET ROWS', sheet.nrows
        print 'SHEET COLS', sheet.ncols

        obcine_col = get_obcine_column(sheet)
        print 'COLUMN', obcine_col

        legenda_row = get_legenda_row(sheet, data_start=obcine_col[1])
        print 'LEGENDA', legenda_row

        obcine = get_obcine(sheet, obcine_col[0], obcine_col[1], obcine_col[2])
        postavke = get_postavke(sheet, legenda_row[0], legenda_row[1], legenda_row[2])

        for x in xrange(legenda_row[1], legenda_row[2] + 1):
            for y in xrange(obcine_col[1], obcine_col[2] + 1):
                fn1 = os.path.basename(fn)
                m = re.match('^(\w\w)-(\d{4})-', fn1)
                if not m:
                    raise ValueError(fn1)
                klasifikacija, leto = m.groups()
                d = collections.OrderedDict()

                d['postavka'] = postavke[x].strip()
                d['obcina'] = obcine_lookup[obcine[y].strip()]
                d['vrednost'] = sheet.cell(y, x).value
                d['sheet'] = sheet.name
                d['cell'] = cellname(y, x)
                d['x'] = x
                d['y'] = y
                d['fn'] = fn1
                d['klasifikacija'] = klasifikacija
                d['leto'] = leto

                yield d


class Cell(object):
    def __init__(self, val):
        self.value = val


class CSVSheet(object):
    def __init__(self, data, name):
        self.data = data
        self.nrows = len(data)
        self.ncols = max([len(i) for i in data])
        self.name = name

    def cell(self, y, x):
        return Cell(self.data[y][x])


def do_csv(fn, startid=1):
    fd = open(fn)
    rdr = csv.reader(fd)
    data = []
    for i in rdr:
        data.append([isinstance(fld, basestring) and fld.decode('utf-8') or fld for fld in i])

    sheetname = fn.rsplit('.', 1)[0].rsplit('-', 1)[1]
    sheet = CSVSheet(data, name=sheetname)

    obcine_lookup = get_obcine_lookup()

    obcine_col = get_obcine_column(sheet)
    print 'COLUMN', obcine_col

    legenda_row = get_legenda_row(sheet, data_start=obcine_col[1])
    print 'LEGENDA', legenda_row

    obcine = get_obcine(sheet, obcine_col[0], obcine_col[1], obcine_col[2])
    postavke = get_postavke(sheet, legenda_row[0], legenda_row[1], legenda_row[2])

    for x in xrange(legenda_row[1], legenda_row[2] + 1):
        for y in xrange(obcine_col[1], obcine_col[2] + 1):
            fn1 = os.path.basename(fn)
            m = re.match('^(\w\w)-(\d{4})-', fn1)
            if not m:
                raise ValueError(fn1)
            klasifikacija, leto = m.groups()
            d = collections.OrderedDict()

            norm_obcina = obcine[y].strip()
            if norm_obcina.endswith(' (M)'):
                norm_obcina = norm_obcina[:-4]
            if norm_obcina.startswith(u'MESTNA OBČINA '):
                norm_obcina = norm_obcina[14:]
            if norm_obcina.startswith(u'OBČINA '):
                norm_obcina = norm_obcina[7:]
            norm_obcina = obcine_lookup[norm_obcina]

            d['postavka'] = postavke[x].strip()
            d['obcina'] = norm_obcina
            d['vrednost'] = sheet.cell(y, x).value.strip().replace('.', '').replace(',', '.')
            d['sheet'] = sheet.name
            d['cell'] = cellname(y, x)
            d['x'] = x
            d['y'] = y
            d['fn'] = fn1
            d['klasifikacija'] = klasifikacija
            d['leto'] = leto

            yield d


def do_all(dirname):
    fd = open('bilanca.csv', 'wb')
    w = csv.writer(fd)

    if dirname.lower().endswith('.csv'):
        files = [dirname]
    else:
        files = []
        for f in sorted(os.listdir(dirname)):
            fullf = os.path.join(dirname, f)
            if f.lower().endswith('.csv') and os.path.isfile(fullf):
                files.append(fullf)

    files.sort()
    pending = set(files)
    header = None
    startid = 1
    for f in files:
        print 'Reading', f
        try:
            for item in do_csv(f, startid=startid):
                if header is None:
                    header = item.keys()
                    print header
                    w.writerow(['id'] + header)
                startid += 1
                w.writerow([startid] + [isinstance(i, basestring) and unicode(i).encode('utf-8') or i for i in item.values()])
        except Exception, e:
            print 'Failed', f, e
            raise
        else:
            pending.remove(f)

if __name__ == "__main__":
    do_all(sys.argv[1])
