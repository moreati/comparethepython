#!/usr/bin/env python
# encoding: utf-8

import codecs
import datetime 
import itertools
import json
import re
import string
import sys

import genshi
import genshi.template.loader
import odf.opendocument
import odf.table
import odf.text
import xlrd

COLUMNS = [
    ('d', 'CPython', '1.5', '1997-12-31', False),
    ('e', 'CPython', '1.6', '2000-09-05', False),
    ('f', 'CPython', '2.0', '2000-10-16', False),
    ('g', 'CPython', '2.1', '2001-04-15', False),
    ('h', 'CPython', '2.2', '2001-12-21', True),
    ('i', 'CPython', '2.3', '2003-07-29', True),
    ('j', 'CPython', '2.4', '2004-11-30', True),
    ('k', 'CPython', '2.5', '2006-09-19', True),
    ('l', 'CPython', '2.6', '2008-10-02', True),
    ('m', 'CPython', '2.7', '2010-07-04', True),
    ('n', 'CPython', '3.0', '2008-12-03', True),
    ('o', 'CPython', '3.1', '2009-06-27', True),
    ('p', 'CPython', '3.2', '2011-02-05', True),
    ('q', 'Jython', '2.0', '2001-01-16', False),
    ('r', 'Jython', '2.1', '2001-12-30', False),
    ('s', 'Jython', '2.2', '2007-08-22', False),
    # TODO Insert CPython 3.3 here
    ('t', 'IronPython', '1.0', '2006-09-05', False),
    ('u', 'IronPython', '1.1', '2007-04-17', False),
    ('v', 'IronPython', '2.0', '2008-12-10', False),
    ('w', 'PyPy', '1.6', '2011-08-23', False),
    ]

URLS = {
    ('CPython', '2.0', 'whatsnew'): u'http://docs.python.org/whatsnew/2.0.html',
    ('CPython', '2.1', 'whatsnew'): u'http://docs.python.org/whatsnew/2.1.html',
    ('CPython', '2.2', 'whatsnew'): u'http://docs.python.org/whatsnew/2.2.html',
    ('CPython', '2.3', 'whatsnew'): u'http://docs.python.org/whatsnew/2.3.html',
    ('CPython', '2.4', 'whatsnew'): u'http://docs.python.org/whatsnew/2.4.html',
    ('CPython', '2.5', 'whatsnew'): u'http://docs.python.org/whatsnew/2.5.html',
    ('CPython', '2.6', 'whatsnew'): u'http://docs.python.org/whatsnew/2.6.html',
    ('CPython', '2.7', 'whatsnew'): u'http://docs.python.org/whatsnew/2.7.html',
    ('CPython', '3.0', 'whatsnew'): u'http://docs.python.org/py3k/whatsnew/3.0.html',
    ('CPython', '3.1', 'whatsnew'): u'http://docs.python.org/py3k/whatsnew/3.1.html',
    ('CPython', '3.2', 'whatsnew'): u'http://docs.python.org/py3k/whatsnew/3.2.html',
    ('CPython', '2.0', 'whatsnew'): u'http://docs.python.org/whatsnew/2.0.html',
    ('CPython', '1.5', ''): u'',
    ('CPython', '1.6', ''): u'',
    ('CPython', '2.0', ''): u'',
    ('CPython', '2.1', ''): u'',
    ('CPython', '2.2', ''): u'',
    ('CPython', '2.3', ''): u'',
    ('CPython', '2.4', ''): u'',
    ('CPython', '2.5', ''): u'',
    ('CPython', '2.6', ''): u'',
    ('CPython', '2.7', ''): u'',
    ('CPython', '3.0', ''): u'',
    ('CPython', '3.1', ''): u'',
    ('CPython', '3.2', ''): u'',
    }

SHEET_NAMES = [
    'Builtins', 
    'Keywords',
    'Modules',
    'Command line',
    'Platforms',
    'Features',
    ]

# ■ 25A0 BLACK SQUARE                    □ 25A1 WHITE SQUARE                   
# ◧ 25E7 SQUARE WITH LEFT HALF BLACK     ◨ 25E8 SQUARE WITH RIGHT HALF BLACK
# ▣ 25A3 WHITE SQUARE CONTAINING BLACK SMALL SQUARE
# ⬚ 2B1A DOTTED SQUARE                   ◌ 25CC DOTTED CIRCLE
# ● 25CF BLACK CIRCLE                    ○ 25CB WHITE CIRCLE
# ◐ 25D0 CIRCLE WITH LEFT HALF BLACK     ◑ 25D1 CIRCLE WITH RIGHT HALF BLACK
# ◖ 25D6 LEFT HALF BLACK CIRCLE          ◗ 25D7 RIGHT HALF BLACK CIRCLE
# ⬛ 2B1B BLACK LARGE SQUARE              ⬜ 2B1C WHITE LARGE SQUARE
# ⬤ 2B24 BLACK LARGE CIRCLE              ◯ 25EF LARGE CIRCLE
# ◼ 25FC BLACK MEDIUM SQUARE             ◻ 25FB WHITE MEDIUM SQUARE
# ◾ 25FE BLACK MEDIUM SMALL SQUARE       ◽ 25FD WHITE MEDIUM SMALL SQUARE
# ☐ 2610 BALLOT BOX                      ☑ 2611 BALLOT BOX WITH CHECK
# ☒ 2612 BALLOT BOX WITH X

_CODES = [
    ('' , u'',  "Not supported"),
    ('/', u"⬛", "Supported"),
    ('f', u'▣', "Supported, with __future__ import"),
    ('e', u"⬛", "Supported, enhanced"),
    ('*', u'⬤', "Supported, changed semantics"),
    ('d', u'◧', "Deprecated"),
    ('u', u'',  "Unsupported in this version"),
    ('?', u'�', "Uknown support"),
     
    ('D', u'D', "Default packaged version"),
    ('O', u'O', "Optional packaged version"),
    ]

MAPPING = dict((char, output) for (char, output, description) in _CODES)
KEY = [(output, description) for (char, output, description) in _CODES]

def transform(s, mapping):
    """Strip leading/trailing whitespace from s, return mapped value or s.
    """
    s = s.strip()
    return mapping.get(s, s)

class ODS(object):
    def __init__(self, fname, start_col, end_col, start_row):
        self.doc = odf.opendocument.load(fname)
        self.start_col = start_col
        self.end_col = end_col
        self.start_row = start_row

    def sheets(self):
        # odf table <==> xls worksheet?
        return [s for s in self.doc.spreadsheet.getElementsByType(odf.table.Table)
                if s.getAttribute('name') in SHEET_NAMES]

    def rows(self, sheet, limit=100):
        rows = sheet.getElementsByType(odf.table.TableRow)
        for row in rows:
            try:
                repeats = row.getAttrNS(odf.opendocument.TABLENS,
                                        'number-rows-repeated') or '1'
            except ValueError:
                repeats = '1'
            for i in xrange(min(int(repeats), limit)):
                yield row

    def cells(self, row, limit=100):
        cells = row.getElementsByType(odf.table.TableCell)
        for cell in cells:
            try:
                repeats = cell.getAttrNS(odf.opendocument.TABLENS,
                                         'number-columns-repeated') or '1'
            except ValueError:
                repeats = '1'
            for i in xrange(min(int(repeats), limit)):
                yield cell

    def do_table(self, sheet):
        subsection = []
        start_subsection = True
        subsection_name = ''
        rows = list(self.rows(sheet))
        for r, row_elem in enumerate(rows[self.start_row:], self.start_row+1):
            row = list(self.cells(row_elem))
            if start_subsection:
                start_subsection = False
                subsection_name = self.fmt_cell(row[0])['text']
            elif not any(cell.childNodes for cell in row[:self.start_col]):
                start_subsection = True
                continue
            labels = tuple(self.fmt_cell(c) for c in row[:self.start_col])
            values = tuple(self.fmt_cell(c)
                           for c, x in itertools.izip_longest(row[self.start_col:self.end_col], COLUMNS)
                           if x[-1])
            if any(c['text'] for c in labels) or any(c['text'] for c in values):
                subsection.append((subsection_name, labels, values))
        return [(x, list(y)) for x, y in 
                itertools.groupby(subsection, lambda x:x[0])]

    def do_tables(self):
        sheets = self.sheets()
        tables = {}
        for sheet in self.sheets():
            table = self.do_table(sheet)
            table_name = sheet.getAttribute('name')
            tables[table_name] = table
        return tables

    def fmt_cell(self, cell):
        try:
            m = re.match(r'of:=hyperlink\("([^"]+)"; *"([^"]+)"\)',
                         cell.getAttribute('formula') or '', re.I | re.U)
        except (AttributeError, ValueError):
            m = None
        try:
            text = '\n'.join(unicode(p) 
                             for p in cell.getElementsByType(odf.text.P))
        except AttributeError:
            text = ''
        if m:
            return dict(href=m.group(1), text=text)
        else:
            return dict(text=text)

def read_ods(fname, start_col, end_col, start_row):
    ods = ODS(fname, start_col, end_col, start_row)
    return ods.do_tables()
    
def read_xls(fname, start_col, end_col, start_row):
    def fmt_cell(cell):
        return dict(text=cell.value if cell else '')

    tables = {}
    workbook = xlrd.open_workbook(fname)
    sheets = [workbook.sheet_by_name(name) for name in SHEET_NAMES]
    for sheet in sheets:
        subsection = []
        start_subsection = True
        subsection_name = ''
        for i in xrange(start_row, sheet.nrows):
            row = sheet.row(i)
            if start_subsection:
                start_subsection = False
                subsection_name = row[0].value
            elif not any(c.value for c in row[:start_col]):
                start_subsection = True
                continue
            labels = tuple(fmt_cell(c) for c in row[:start_col])
            values = tuple(fmt_cell(c)
                           for c, x in itertools.izip_longest(row[start_col:end_col], COLUMNS)
                           if x[-1])
            subsection.append((subsection_name, labels, values))
        tables[sheet.name] = [(x, list(y)) for x, y in 
                              itertools.groupby(subsection, lambda x:x[0])]
    return tables

def gdoc_cell(cell):
    x = str(cell)
    #if 'python.org' in x:
    print cell
    print repr(cell)
    print x
    try:
        m = re.match(r'=hyperlink\("([^"]+)", *"([^"]+)"\)',
                     str(cell.cell.inputValue), re.I)
        if m:
            return dict(href=m.group(1), text=m.group(2).strip())
    except AttributeError:
        print cell
        pass
    except:
        print cell
        raise
    try:
        return dict(text=cell.content.text.strip())
    except AttributeError:
        return dict(text='')
       
def read_gdocs(spreadsheet_id, start_col, end_col, start_row, auth=None):
    try: 
        from xml.etree import ElementTree
    except ImportError:  
        from elementtree import ElementTree
    import gdata.spreadsheet.service
    import gdata.service
    import atom.service
    import gdata.spreadsheet
    import atom

    tables = {}

    # As of Sept 2011 only basic and values projection is implemented by the API
    # for public spreadsheets. 'full' projection (and hence private visibility)
    # is necessary to retrieve formulas - which contain hyperlink URLs
    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    if auth:
        gd_client.email = auth['email']
        gd_client.password = auth['password']
        gd_client.source = auth['source']
        gd_client.ProgrammaticLogin()
        visibility = 'private'
        projection = 'full'
    else:
        visibility = 'public',
        projection = 'values'

    feed = gd_client.GetWorksheetsFeed(spreadsheet_id, visibility=visibility,
                                       projection=projection)
    sheet_ids = dict((entry.title.text, entry.id.text.split('/')[-1])
                     for entry in feed.entry)
    for sheet_name in SHEET_NAMES:
        query = gdata.spreadsheet.service.CellQuery()
        # CellQuery row/col parameters are 1-based
        query['min-row'] = str(start_row + 1)
        query['min-col'] = str(1)
        query['max-col'] = str(end_col + 1)
        query['return-empty'] = str(True)
        feed = gd_client.GetCellsFeed(spreadsheet_id, sheet_ids[sheet_name],
                                      query=query, visibility=visibility,
                                      projection=projection)
        subsection = []
        start_subsection = True
        subsection_name = ''
        for row_num, cells in itertools.groupby(feed.entry, lambda e: e.cell.row):
            cells = list(cells)
            row = [gdoc_cell(c) for c in cells]
            if start_subsection:
                start_subsection = False
                subsection_name = row[0]['text']
            elif not any(c['text'] for c in row[:start_col]):
                start_subsection = True
                continue
            labels = tuple(row[:start_col])
            values = tuple(c for c, x in zip(row[start_col:end_col], COLUMNS)
                           if x[-1])
            subsection.append((subsection_name, labels, values))
        tables[sheet_name] = [(x, list(y)) for x, y in
                              itertools.groupby(subsection, lambda x:x[0])]
    print sheet_name, row_num, subsection_name
    return tables

def write_json(x, filename):
    f = open(filename, 'w')
    json.dump(x, f, indent=4)
    f.close()

def main():
    import argparse
    
    parser = argparse.ArgumentParser()
    parser.add_argument('--refresh', )
    args = parser.parse_args()

    pythons = [(name, ver, datetime.datetime.strptime(rel_date, '%Y-%m-%d'))
               for col, name, ver, rel_date, show in COLUMNS
               if show]
    start_col = ord(COLUMNS[0][0]) - ord('a')
    end_col = ord(COLUMNS[-1][0]) - ord('a')
    start_row = 3

    if args.refresh == 'xls':
        tables = read_xls('/home/alex/Downloads/Python comparison matrix.xls',
                          start_col, end_col, start_row)
        write_json(tables, 'index.json')
    elif args.refresh == 'ods':
        tables = read_ods('/home/alex/Downloads/Python comparison matrix.ods',
                          start_col, end_col, start_row)
        write_json(tables, 'index.json')
    elif args.refresh == 'gdocs':
        tables = read_gdocs('0At5kubLl6ri7dHU2OEJFWkJ1SE16NUNvaGg2UFBxMUE',
                            start_col, end_col, start_row,
                            )
        write_json(tables, 'index.json')
    else:
        tables = json.load(open('index.json', 'r'))

    loader = genshi.template.loader.TemplateLoader(['./templates'])
    template = loader.load('index.html')
    stream = template.generate(
                sections=SHEET_NAMES,
                colgroups=[(python, len(list(releases))) 
                           for python, releases in itertools.groupby(pythons, lambda x:x[0])],
                pythons=pythons,
                nlabelcols=start_col,
                tables=tables,
                transform=transform,
                mapping=MAPPING,
                key=KEY,
                )
    #outf = codecs.open('index.html', 'w', encoding='utf-8')
    outf = open('index.html', 'w')
    outf.write(stream.render('xhtml'))
    outf.close()
    
    #for name in SHEET_NAMES:
    #    print name
    #    for section_name, labels, values in tables[name]:
    #        print '  %s %s' % (section_name, labels)

if __name__ == '__main__':
    main()

