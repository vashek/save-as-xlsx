# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
from __future__ import annotations

import json
from collections.abc import Iterable
from dataclasses import asdict, fields, is_dataclass
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from enum import Enum
from fractions import Fraction
from os import PathLike, fspath
from typing import Any, ClassVar, Protocol, TypeVar

import xlsxwriter
import xlsxwriter.worksheet
from xlsxwriter.exceptions import XlsxWriterException

try:
    import pydantic
    from pydantic import BaseModel
    PYDANTIC_VER = int(pydantic.__version__.split(".")[0])
except ImportError:
    class BaseModel: pass
    PYDANTIC_VER = -1

from .__about__ import __version__ as __version__


class TableAddError(XlsxWriterException):
    pass


class WorkbookClosedError(XlsxWriterException):
    pass


class UnsupportedTypeError(XlsxWriterException, TypeError):
    pass


class DataclassInstance(Protocol):
    __dataclass_fields__: ClassVar[dict[str, Any]]


class SaveAsXlsx:
    PydanticModel = TypeVar("PydanticModel", bound=BaseModel)

    def __init__(self,
                 filename: str | PathLike,
                 data: Iterable[dict | DataclassInstance | PydanticModel] | None = None,
                 sheet_name: str | None = None,
                 table_name: str | None = None,
                 column_order: Iterable[str] | None = None,
                 *,
                 extra_columns: bool = True,
                 total_row: bool = False,
                 auto_save: bool = False,
                 ) -> None:
        self.closed = False
        self.filename = filename
        self.workbook = xlsxwriter.Workbook(fspath(filename))
        self.worksheet: xlsxwriter.worksheet.Worksheet | None = None
        self.columns: dict[str, dict[str, str | int | float]] = {}
        self.columns_values: list[dict[str, str | int | float]] = []
        self.number_of_value_rows = 0
        if data is not None:
            self.add_sheet(data, sheet_name=sheet_name, table_name=table_name, column_order=column_order, extra_columns=extra_columns, total_row=total_row)
        if auto_save:
            self.close()

    def add_sheet(self,
                  data: Iterable[dict | DataclassInstance | BaseModel],
                  sheet_name: str | None = None,
                  table_name: str | None = None,
                  column_order: Iterable[str] | None = None,
                  *,
                  extra_columns: bool = True,
                  total_row: bool = False,
                  ) -> xlsxwriter.worksheet.Worksheet:
        if self.closed:
            raise WorkbookClosedError()
        self.worksheet = worksheet = self.workbook.add_worksheet(sheet_name)
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
                ) if col_name not in columns]
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
        return worksheet

    def close(self) -> None:
        if not self.closed:
            self.workbook.close()
        self.closed = True

    def __enter__(self) -> SaveAsXlsx:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if exc_type is None and exc_val is None:
            self.close()

    @classmethod
    def convert_value(cls, input_value, *, for_json: bool = False):
        if isinstance(input_value, Enum):  # must be first, because enum may match str or int
            return input_value.name
        if isinstance(input_value, (str, int, float, bool, Decimal, Fraction, datetime, date, time, timedelta)):
            if for_json and not isinstance(input_value, (str, int, float, bool)):
                return float(input_value)
            return input_value
        if input_value is None:
            return None
        if isinstance(input_value, list):
            return "[" + ", ".join(str(cls.convert_value(value)) for value in input_value) + "]"
        if isinstance(input_value, set):
            return "{" + ", ".join(str(cls.convert_value(value)) for value in input_value) + "}"
        if isinstance(input_value, dict):
            return json.dumps(input_value, default=lambda value: cls.convert_value(value, for_json=True))
        raise UnsupportedTypeError(input_value)
