# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
from __future__ import annotations
import json
from dataclasses import is_dataclass, fields, asdict
from datetime import datetime, date, time, timedelta
from decimal import Decimal
from enum import Enum
from fractions import Fraction
from os import PathLike, fspath
from typing import Any, ClassVar, Iterable, Protocol

import xlsxwriter
from xlsxwriter.exceptions import XlsxWriterException

try:
    from pydantic import BaseModel
    import pydantic
    PYDANTIC_VER = int(pydantic.__version__.split(".")[0])
except ImportError:
    class BaseModel: pass
    PYDANTIC_VER = -1


class TableAddError(XlsxWriterException):
    pass


class UnsupportedTypeError(XlsxWriterException, TypeError):
    pass


class DataclassInstance(Protocol):
    __dataclass_fields__: ClassVar[dict[str, Any]]


class SaveAsXlsx:
    def __init__(self,
                 data: Iterable[dict | DataclassInstance | BaseModel],
                 filename: str | PathLike,
                 sheet_name: str | None = None,
                 table_name: str | None = None,
                 column_order: Iterable[str] | None = None,
                 extra_columns: bool = True,
                 total_row: bool = False,
                 auto_save: bool = False,
                 ) -> None:
        self.closed = False
        self.filename = filename
        self.workbook = workbook = xlsxwriter.Workbook(fspath(filename))
        self.worksheet = worksheet = workbook.add_worksheet(sheet_name)
        self.columns = columns = {}
        for column in column_order or ():
            columns[column] = {"header": column}
        if not isinstance(data, (list, set, tuple)):
            data = list(data)
        if extra_columns:
            for row in data:
                # missing_cols = row.keys() - columns.keys()  # not order-preserving, so instead:
                missing_cols = [col_name for col_name in (
                    (f.name for f in fields(row)) if is_dataclass(row) else
                    type(row).model_fields.keys() if PYDANTIC_VER >= 2 and isinstance(row, BaseModel) else
                    row.__fields__.keys() if isinstance(row, BaseModel) else
                    row.keys()
                ) if col_name not in columns.keys()]
                for column in missing_cols:
                    columns[column] = {"header": column}
        col_names = columns.keys()
        self.columns_values = list(columns.values())
        self.number_of_value_rows = len(data)
        result = worksheet.add_table(0, 0, len(data), len(columns) - 1, {
            "header_row": True,
            "columns": self.columns_values,
            "total_row": total_row,
            **({"name": table_name} if table_name else {}),
            "data": [
                [self.convert_value(row_dict.get(col_name)) for col_name in col_names]
                for row_union in data
                if (row_dict := (asdict(row_union) if is_dataclass(row_union) else
                                 row_union.model_dump() if PYDANTIC_VER >= 2 and isinstance(row_union, BaseModel) else
                                 row_union.dict() if isinstance(row_union, BaseModel) else
                                 row_union))
            ],
        })
        if result != 0:
            raise TableAddError(f"Table add error: {result}")
        if auto_save:
            self.close()

    def close(self) -> None:
        self.workbook.close()
        self.closed = True

    def __enter__(self) -> SaveAsXlsx:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if exc_type is None and exc_val is None:
            self.close()

    @classmethod
    def convert_value(cls, input_value):
        if isinstance(input_value, (str, int, float, bool, Decimal, Fraction, datetime, date, time, timedelta)):
            return input_value
        elif input_value is None:
            return None
        elif isinstance(input_value, list):
            return "[" + ", ".join(str(cls.convert_value(value)) for value in input_value) + "]"
        elif isinstance(input_value, set):
            return "{" + ", ".join(str(cls.convert_value(value)) for value in input_value) + "}"
        elif isinstance(input_value, dict):
            return json.dumps(input_value, default=cls.convert_value)
        elif isinstance(input_value, Enum):
            return input_value.name
        else:
            raise UnsupportedTypeError(input_value)
