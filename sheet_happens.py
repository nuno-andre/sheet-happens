#!/usr/bin/env python3
"""
Sheet Happens
https://github.com/nuno-andre/sheet-happens

Copyright (C) 2017 Nuno André <mail@nunoand.re>
SPDX-License-Identifier: MIT
"""
__version__ = '0.0.2'
__description__ = 'Simple Excel 2007+ to CSV and JSON converter without dependencies'


from string import digits, ascii_uppercase
from xml.etree import ElementTree
from zipfile import ZipFile, BadZipFile
from pathlib import Path
import json
import csv
try:
    import yaml
    __description__ = __description__.replace(
        ' and JSON', ', JSON, and YAML')
except ImportError:
    yaml = None


ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
NS = {'namespaces': ns}


class Book:
    '''Excel file  

    :param path: Archive's path  
    :param sanitize: Trims strings and replaces line feeds with whitespaces
    '''
    def __init__(self, path, sanitize=True):
        self.path     = Path(path).resolve()
        self.sanitize = sanitize
        self._shared  = None

    @property
    def shared(self):
        '''Shared strings
        '''
        if self._shared is None:
            with self.zipfile.open('xl/sharedStrings.xml') as f:
                tree = ElementTree.fromstring(f.read())
            nodes = tree.findall('.//main:t', **NS)
            self._shared = [x.text for x in nodes]
        return self._shared

    @property
    def sheets(self):
        with ZipFile(str(self.path), 'r') as z:
            self.zipfile = z
            for path in z.namelist():
                if path.startswith('xl/worksheets/sheet'):
                    with z.open(path) as f:
                        yield Sheet(Path(path), f.read(), self)

class Sheet:
    '''Worksheet
    '''
    def __init__(self, path, text, book):
        self.book    = book
        self.path    = path
        self.tree    = ElementTree.fromstring(text)
        # TODO: current name in "sheetx" form
        #     search for and sanitize sheet's name
        self.name    = path.stem
        self.cols    = dict()
        self._dict   = None
        self._parsed = None
        self.width, self.height = self.shape()


    def col(self, col):
        '''Converts and caches a letter-based col to 0-based coord
        '''
        x = reversed([ascii_uppercase.find(l) for l in col])
        x = sum(26 * i + n for i, n in enumerate(x))
        self.cols[col] = x
        return x


    def coords(self, cell):
        '''Converts Excel coords to 0-based
        '''
        col = ''.join(x for x in cell if x not in digits)
        row = int(''.join(x for x in cell if x not in col)) - 1
        col = self.cols.get(col, self.col(col))
        return col, row


    def cell(self, node):
        '''Extracts cell's coords
        '''
        return self.coords(node.attrib['r'])


    def shape(self):
        '''Returns table dimensions
        '''
        dim    = self.tree.find('.//main:dimension', **NS)
        nw, se = dim.attrib['ref'].split(':')
        width, height = [n + 1 for n in self.coords(se)]
        return width, height


    def value(self, node):
        '''Returns cell's value
        '''
        v = node.find('.//main:v', **NS).text
        if node.attrib.get('t') == 's':
            value = self.book.shared[int(v)]
        else:
            value = v

        if self.book.sanitize:
            return ' '.join(filter(None, value.strip().splitlines()))
        else:
            return value

    @property
    def parsed(self):
        '''Returns parsed sheet preallocating rows
        '''
        if self._parsed is None:
            self._parsed = [[None for _ in range(self.width)]
                for _ in range(self.height)]

            for node in self.tree.findall('.//main:c', **NS):
                col, row = self.cell(node)
                value    = self.value(node)
                self._parsed[row][col] = value
        return self._parsed


    def parse(self):
        '''Returns a parsed rows generator
        '''
        row     = [None for _ in range(self.width)]
        lastcol = self.width - 1

        for node in self.tree.findall('.//main:c', **NS):
            col, _   = self.cell(node)
            value    = self.value(node)
            row[col] = value
            if col == lastcol:
                yield row
                row = [None for _ in range(self.width)]


    def to_dict(self):
        '''Returns a list of dicts generator
        '''
        output = self.parse()
        header = next(output)
        return (dict(zip(header, row)) for row in output)


    @property
    def dict(self):
        if self._dict is None:
            if self._parsed:
                header, *rows = self.parsed
                self._dict = [dict(zip(header, row)) for row in rows]
            else:
                self._dict = list(self.to_dict())
        return self._dict


    def filedes(self, ext, path, mode='w', newline='\n'):
        '''Returns a file descriptor
        '''
        path = Path(path or self.book.path)
        if path.is_dir():
            path = path / self.name
        else:
            name = self.name.replace('sheet', path.stem + '.')
            path = path.with_name(name)
        path = path.with_name('{}.{}'.format(path.name, ext))
        return open(str(path), mode, newline=newline)

             
    def to_csv(self, path=None):
        '''Path defaults to <file_path>/<file_stem>.<sheet_no>.csv
        '''
        with self.filedes('csv', path, 'w', newline='') as f:
            writer = csv.writer(f)
            for row in self.parse():
                row = [c.encode('utf8').decode('utf-8') for c in row]
                writer.writerow(row)
        return True


    def to_json(self, path=None):
        '''Path defaults to <file_path>/<file_stem>.<sheet_no>.json
        '''
        with self.filedes('json', path, 'w', newline='') as f:
            json.dump(self.dict, f, indent=4, ensure_ascii=False)
            return True


    def to_yaml(self, path=None):
        '''Path defaults to <file_path>/<file_stem>.<sheet_no>.yaml
        '''
        with self.filedes('json', path, 'w', newline='') as f:
            yaml.dump(list(self.to_dict()), default_flow_style=False)
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
        book = Book(path)
        for sheet in book.sheets:
            for fmt in fmts:
                print('Saving {} as {}'.format(sheet.name, fmt))
                method = 'to_{}'.format(fmt)
                getattr(sheet, method)()
        return 0
    except BadZipFile:
        msg = 'ERROR. "{}" is not an Excel 2007+ file'
        print(msg.format(path))
    except Exception as e:
        print('ERROR. {} {}'.format(e, type(e)))
        return 1

if __name__ == '__main__':
    import sys
    sys.exit(main())
