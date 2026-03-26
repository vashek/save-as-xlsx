# save_as_xlsx

[![PyPI - Version](https://img.shields.io/pypi/v/save-as-xlsx.svg)](https://pypi.org/project/save-as-xlsx)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/save-as-xlsx.svg)](https://pypi.org/project/save-as-xlsx)

-----

## Table of Contents

- [About](#about)
- [Installation](#installation)
- [Usage](#usage)
- [License](#license)

## About

If you have some tabular data, this package gives you a trivial one-line way of saving it as an Excel (.xlsx)
file. The data will be saved formatted as a Table (with a header row, filtering, sorting and nice formatting).

Just pass any iterable (list, tuple, generator...) of:
* dictionaries
* dataclasses
* Pydantic model instances

Nesting of complex data types is handled.

Enums are saved as the enum member name.

Uses the xlsxwrite package to do the actual writing.

## Installation

```console
pip install save-as-xlsx
```

## Usage

```python3
from save_as_xlsx import SaveAsXlsx

DATA = [
    {"a": 1, "b": "qwe"},
    {"b": "asd", "c": True},
]

# simplest case
SaveAsXlsx(DATA, "file.xlsx", auto_save=True)

# or if you want to customize the XLSX file before saving:
with SaveAsXlsx(DATA, "file.xlsx") as saver:
    # do something with saver.workbook or saver.worksheet (see xlsxwriter)
    pass

# the data can be any iterable - tuple, generator...
SaveAsXlsx(({"num": i} for i in range(5)), "file.xlsx", auto_save=True)

# file name can be a Path
from pathlib import Path
SaveAsXlsx(DATA, Path("file.xlsx"), auto_save=True)
# saved columns: a, b, c

# you can specify the order of columns - these will be first, remaining ones after them
SaveAsXlsx(DATA, "file.xlsx", column_order=("b", "c"))
# saved columns: b, c, a

# or maybe you just want some of the columns, and an empty one
SaveAsXlsx(DATA, "file.xlsx", column_order=("b", "empty"), extra_columns=False)
# saved columns: b, empty
```

## License

`save-as-xlsx` is distributed under the terms of the [MIT](https://spdx.org/licenses/MIT.html) license.
