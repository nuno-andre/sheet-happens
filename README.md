# Sheet Happens

Simple `.xlsx` (Excel 2007+) to `.csv`, `.json`, and `.yaml` converter without
dependencies.

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

Output path is `<file_path>/<file_stem>/<sheet_no>_<sheet_name>.<format>`

### As a library

```python
from sheet_happens import Book

for sheet in Book(path):
    print('Sheet:', sheet.name)
    # print a dict {field:value} for each row
    for row in sheet:
        print(row)
```

## YAML support

If [`PyYAML`][0] is available, _Sheet Happens_ will add a `--yaml` option.

---
Copyright &copy; 2017-2021 Nuno André <<mail@nunoand.re>>

[0]: https://github.com/yaml/pyyaml
