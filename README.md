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

You can pass data in many different formats and SaveAsXlsx tries to automagically do the right thing.
Just pass any:
* iterable (list, tuple, generator...) of:
  * dictionaries / mappings
  * dataclasses
  * Pydantic model instances
* or even just a mapping (dictionary) of:
  * simple values (this will produce columns "key" and "value")
  * dicts / mappings or dataclasses or Pydantic model instances
    (this will produce columns "key" and then columns based on the keys of the mappings/classes)
  * iterables (this will produce columns "key" and then "col1", "col2" etc.)

Nesting of complex data types is handled.

Enums are saved as the enum member name.
UUIDs as their hex representation (e.g. "5f456a18-29f0-11f1-a203-e41fd5b9abcb").
Decimal and Fraction as their float representation.

Uses the xlsxwrite package to do the actual writing.

## Installation

```console
pip install save-as-xlsx
```

## Usage

```python3
from save_as_xlsx import SaveAsXlsx, save_as_xlsx

DATA = [
    {"a": 1, "b": "qwe"},
    {"b": "asd", "c": True},
]
OTHER_DATA = [
    {"Name": "John", "Age": 46},
    {"Name": "Jane", "Age": 42},
]

# simplest case
save_as_xlsx("file.xlsx", DATA)

# or if you want to customize the XLSX file before saving, e.g. add another sheet:
with SaveAsXlsx("file.xlsx", DATA) as saver:
    # do something with saver.workbook or saver.worksheet (see xlsxwriter)
    saver.add_sheet(OTHER_DATA)
    # and maybe you want to protect the sheet except one column:
    saver.worksheet.protect()
    saver.worksheet.unprotect_range(saver.column_ref("Age"))

# the data can be any iterable - tuple, generator...
SaveAsXlsx("file.xlsx", ({"num": i} for i in range(5)), auto_save=True)

# file name can be a Path
from pathlib import Path
save_as_xlsx(Path("file.xlsx"), DATA)
# saved columns: a, b, c

# you can specify the order of columns - these will be first, remaining ones after them
save_as_xlsx("file.xlsx", DATA, column_order=("b", "c"))
# saved columns: b, c, a

# or maybe you just want some of the columns, and an empty one
save_as_xlsx("file.xlsx", DATA, column_order=("b", "empty"), extra_columns=False)
# saved columns: b, empty

# you can also specify the sheet and/or table name
with SaveAsXlsx("file.xlsx", DATA, sheet_name="FirstSheet", table_name="FirstTable") as saver:
    saver.add_sheet(OTHER_DATA, sheet_name="AnotherSheet", table_name="AnotherTable")

# or you can do the same like this
with SaveAsXlsx("file.xlsx") as saver:
    saver.add_sheet(DATA, sheet_name="FirstSheet", table_name="FirstTable")
    saver.add_sheet(OTHER_DATA, sheet_name="AnotherSheet", table_name="AnotherTable")

# you can also specify custom column headings
save_as_xlsx(Path("file.xlsx"), DATA, column_headings={"a": "First Column"})

# to retry saving, perhaps with a different name:
from xlsxwriter.exceptions import FileCreateError
saver = SaveAsXlsx("file.xlsx", DATA)
try:
    saver.close()
except FileCreateError:
    saver.close("file-new.xlsx")
```

## License

`save-as-xlsx` is distributed under the terms of the [MIT](https://spdx.org/licenses/MIT.html) license.
