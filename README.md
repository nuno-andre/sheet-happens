# Sheet Happens

Simple `.xlsx` (Excel 2007+) to `.csv`, `.json`, and `.yaml` converter without dependencies.

## Installation
```
git clone https://github.com/nuno-andre/sheet-happens.git
cd sheet-happens
python setup.py install
```

## Usage
### As an application

```
$ sheet-happens <path-to-file> --csv --json
```

Output path is `<file_path>/<file_stem>.<sheet_no>.<format>`

### As a library
```
from sheet_happens import Book

for sheet in Book(path):
    print('Sheet:', sheet.name)
    for row in sheet.to_dict():
        print(row)
```

## YAML support

If [`PyYAML`](https://github.com/yaml/pyyaml) is available, _Sheet Happens_ will add a `--yaml` option.

---
Copyright &copy; 2017-2021 Nuno André <<mail@nunoand.re>>
