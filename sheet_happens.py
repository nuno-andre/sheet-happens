#!/usr/bin/env python3
"""
Sheet Happens
https://github.com/nuno-andre/sheet-happens

Copyright (C) 2017-2021 Nuno André <mail@nunoand.re>
SPDX-License-Identifier: MIT
"""
__version__ = 0, 0, 5
__description__ = ('Simple Excel 2007+ to CSV and JSON converter without '
                   'dependencies')


from xml.etree.ElementTree import fromstring
from string import digits, ascii_uppercase
from zipfile import ZipFile, BadZipFile
from functools import wraps
from pathlib import Path
import json
import csv
try:
    import yaml
    __description__ = __description__.replace(
        ' and JSON', ', JSON, and YAML')
except ImportError:
    yaml = None


MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'namespaces': {'main': MAIN}}


def lazyproperty(method):
    '''Decorator for lazy evaluated properties.
    '''
    _name = "_" + method.__name__

    @wraps(method)
    def wrapper(self, *args, **kwargs):
        if not hasattr(self, _name):
            setattr(self, _name, method(self, *args, **kwargs))
        return getattr(self, _name)

    return property(wrapper, doc=method.__doc__)


class Book:
    '''Excel file (workbook).

    Args:
        path: Archive's path.
        sanitize: Trims strings and replaces line feeds with whitespaces.
    '''
    def __init__(self, path, sanitize=True):
        self.path     = Path(path).resolve()
        self.sanitize = sanitize
        self._names   = dict()

    def __iter__(self):
        return self.sheets.__iter__()

    def __next__(self):
        return next(self.__iter__())

    @lazyproperty
    def shared(self):
        '''Shared strings
        '''
        with self.zipfile.open('xl/sharedStrings.xml') as f:
            tree = fromstring(f.read())
            return [x.text for x in tree.iterfind('.//main:t', **NS)]

    @lazyproperty
    def sheets(self):
        sheets = list()
        tag_id = '{{{}}}id'.format(REL)

        with ZipFile(str(self.path), 'r') as z:
            with z.open('xl/workbook.xml') as f:
                tree = fromstring(f.read())
                for s in tree.iterfind('.//main:sheet', **NS):
                    id = s.get(tag_id).replace('rId', '')
                    self._names[int(id)] = s.get('name')
            self.zipfile = z
            for path in z.namelist():
                if path.startswith('xl/worksheets/sheet'):
                    with z.open(path) as f:
                        sheets.append(Sheet(Path(path), f.read(), self))
            self.shared
        return sheets


class Sheet:
    '''Worksheet
    '''
    def __init__(self, path, text, book):
        self.book = book
        self.path = path
        self.tree = fromstring(text)
        self.cols = dict()
        self.width, self.height = self.shape()

    @lazyproperty
    def name(self):
        pad = len(str(list(self.book._names.keys())[-1]))
        ix  = int(self.path.stem.replace('sheet', ''))
        return '{i:0{p}}_{n}'.format(i=ix, p=pad, n=self.book._names[ix])

    def col(self, col):
        '''Converts and caches a letter-based col to 0-based coord.
        '''
        x = reversed([ascii_uppercase.find(l) for l in col])
        x = sum(26 * i + n for i, n in enumerate(x))
        self.cols[col] = x
        return x

    def coords(self, cell):
        '''Converts Excel coords to 0-based.
        '''
        col = ''.join(x for x in cell if x not in digits)
        row = int(''.join(x for x in cell if x not in col)) - 1
        col = self.cols.get(col, self.col(col))
        return col, row

    def cell(self, node):
        '''Extracts cell's coords.
        '''
        return self.coords(node.attrib['r'])

    def shape(self):
        '''Returns table dimensions.
        '''
        dim    = self.tree.find('.//main:dimension', **NS)
        nw, se = dim.attrib['ref'].split(':')
        width, height = (n + 1 for n in self.coords(se))
        return width, height

    def value(self, node):
        '''Returns cell's value.
        '''
        try:
            v = node.find('.//main:v', **NS).text
            if node.attrib.get('t') == 's':
                value = self.book.shared[int(v)]
            else:
                value = v
        except AttributeError:
            value = None

        if self.book.sanitize and value:
            return ' '.join(filter(None, value.strip().splitlines()))
        else:
            return value

    @lazyproperty
    def parsed(self):
        '''Returns parsed sheet preallocating rows.
        '''
        parsed = [[None for _ in range(self.width)]
                  for _ in range(self.height)]

        for node in self.tree.findall('.//main:c', **NS):
            col, row = self.cell(node)
            value = self.value(node)
            parsed[row][col] = value

        return parsed

    def parse(self):
        '''Returns a parsed rows generator.
        '''
        row     = [None for _ in range(self.width)]
        lastcol = self.width - 1

        for node in self.tree.findall('.//main:c', **NS):
            col, _   = self.cell(node)
            row[col] = self.value(node)
            if col == lastcol:
                yield row
                row = [None for _ in range(self.width)]

    def to_dict(self):
        '''Returns a list of dicts generator.
        '''
        output = self.parse()
        header = next(output)
        return (dict(zip(header, row)) for row in output)

    @lazyproperty
    def dict(self):
        if hasattr(self, '_parsed'):
            header, *rows = self.parsed
            return [dict(zip(header, row)) for row in rows]
        else:
            return list(self.to_dict())

    def filedes(self, ext, dirpath=None, mode='w', newline='\n'):
        '''Returns a file descriptor.

        Path defaults to:
          ``<file_path>/<file_stem>/<sheet_no>_<sheet_name>.<ext>``
        '''
        if dirpath:
            dirpath = Path(dirpath)
        else:
            # set excel archive's filename as directory
            dirpath = self.book.path.parent / self.book.path.stem

        try:
            dirpath.mkdir(exist_ok=True, parents=True)
        except FileExistsError:
            err = f"cannot creake directory '{dirpath}'"
            raise ValueError(err) from None

        path = dirpath / f'{self.name}.{ext}'
        return open(str(path), mode, newline=newline)

    def to_csv(self, path=None):
        with self.filedes('csv', path, 'w', newline='') as f:
            writer = csv.writer(f)
            for row in self.parse():
                row = [c.encode('utf8').decode('utf-8') for c in row]
                writer.writerow(row)
        return True

    def to_json(self, path=None):
        with self.filedes('json', path, 'w', newline='') as f:
            json.dump(self.dict, f, indent=4, ensure_ascii=False)
            return True

    def to_yaml(self, path=None):
        with self.filedes('yaml', path, 'w', newline='') as f:
            f.write(yaml.dump(self.dict, default_flow_style=False,
                              allow_unicode=True, sort_keys=False))
            return True


def main():
    import argparse

    p = argparse.ArgumentParser(
        prog='sheet-happens',
        description=__description__,
    )
    p.add_argument('filepath')
    p.add_argument('--csv', action='store_const', const=True)
    p.add_argument('--json', action='store_const', const=True)
    if yaml:
        p.add_argument('--yaml', action='store_const', const=True)

    args = vars(p.parse_args())
    path = args.pop('filepath')
    fmts = [k for k, v in args.items() if v]

    if not fmts:
        print('\nERROR. Choose at least one output format.\n')
        p.print_help()
        return 1

    try:
        for sheet in Book(path):
            for fmt in fmts:
                print("Saving '{}' as '{}'".format(sheet.name, fmt))
                method = 'to_{}'.format(fmt)
                getattr(sheet, method)()
        return 0
    except BadZipFile:
        msg = "ERROR. '{}' is not an Excel 2007+ file"
        print(msg.format(path))
    except Exception as e:
        print('ERROR. {!r}'.format(e))
        return 1


__all__ = ['Book', 'Sheet']

if __name__ == '__main__':
    import sys
    sys.exit(main())
