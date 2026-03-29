# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
from __future__ import annotations

import json
from collections.abc import Iterable, Mapping, Set, Sized
from dataclasses import asdict, fields, is_dataclass
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from enum import Enum, IntEnum
from fractions import Fraction
from os import PathLike, fspath
from typing import Annotated, Any, ClassVar, Protocol, TypeAlias
from uuid import UUID

import xlsxwriter  # type: ignore
import xlsxwriter.worksheet  # type: ignore
from annotated_types import Gt, Lt, Unit
from xlsxwriter.exceptions import XlsxWriterException  # type: ignore
from xlsxwriter.utility import xl_col_to_name  # type: ignore

try:
    import pydantic  # type: ignore
    from pydantic import BaseModel
    PYDANTIC_VER = int(pydantic.__version__.split(".")[0])
except ImportError:
    class BaseModel:  # type: ignore
        pass
    PYDANTIC_VER = -1

from .__about__ import __version__

__all__ = [
    "ColumnWidth",
    "SaveAsXlsx",
    "TableAddError",
    "UnsupportedTypeError",
    "WorkbookClosedError",
    "__version__",
    "save_as_xlsx",
]


class TableAddError(XlsxWriterException):
    pass


class WorkbookClosedError(XlsxWriterException):
    pass


class UnsupportedTypeError(XlsxWriterException, TypeError):
    pass


class DataclassInstance(Protocol):
    __dataclass_fields__: ClassVar[dict[str, Any]]


class ColumnWidth(IntEnum):
    AUTOFIT = -1
    HIDE = 0


ColumnWidthType: TypeAlias = None | ColumnWidth | Annotated[int|float, Gt(0), Unit("px")] | Annotated[int|float, Lt(0), Unit("characterUnits")]


class SaveAsXlsx:
    def __init__(self,
                 filename: str | PathLike,
                 data: Iterable[Mapping[str, Any] | DataclassInstance | BaseModel] | Mapping[str, Any] | None = None,
                 sheet_name: str | None = None,
                 table_name: str | None = None,
                 column_order: Iterable[str] | None = None,
                 column_width: ColumnWidthType | Mapping[str, ColumnWidthType] | Iterable[ColumnWidthType] = None,
                 *,
                 extra_columns: bool = True,
                 total_row: bool = False,
                 strings_to_numbers: bool = False,
                 strings_to_formulas: bool = False,
                 strings_to_urls: bool = True,
                 nan_inf_to_errors: bool = True,
                 remove_timezone: bool = False,
                 default_date_format: str | None = None,
                 auto_save: bool = False,
                 ) -> None:
        self.closed = False
        self.filename = filename
        workbook_options: dict[str, bool | str] = {  # TODO: document these
            "strings_to_numbers": strings_to_numbers,
            "strings_to_formulas": strings_to_formulas,
            "strings_to_urls": strings_to_urls,
            "nan_inf_to_errors": nan_inf_to_errors,
            "remove_timezone": remove_timezone,
        }
        if default_date_format:
            workbook_options["default_date_format"] = default_date_format
        self.workbook = xlsxwriter.Workbook(fspath(filename), workbook_options)
        self.worksheet: xlsxwriter.worksheet.Worksheet | None = None
        self.columns: dict[str, dict[str, str | int | float]] = {}
        self.columns_values: tuple[dict[str, str | int | float], ...] = ()
        self.number_of_value_rows = 0
        if data is not None:
            self.add_sheet(data, sheet_name=sheet_name, table_name=table_name, column_order=column_order, column_width=column_width, extra_columns=extra_columns, total_row=total_row)
        if auto_save:
            self.close()

    def add_sheet(self,
                  data: Iterable[Mapping[str, Any] | DataclassInstance | BaseModel] | Mapping[str, Any],
                  sheet_name: str | None = None,
                  table_name: str | None = None,
                  column_order: Iterable[str] | None = None,
                  column_width: ColumnWidthType | Mapping[str, ColumnWidthType] | Iterable[ColumnWidthType] = None,
                  *,
                  extra_columns: bool = True,
                  total_row: bool = False,
                  ) -> xlsxwriter.worksheet.Worksheet:
        if self.closed:
            raise WorkbookClosedError
        self.worksheet = worksheet = self.workbook.add_worksheet(sheet_name)
        columns: dict[str, dict[str, str | int | float]] = {}
        self.columns = columns
        for column in column_order or ():
            columns[column] = {"header": column}
        if isinstance(data, Mapping):
            any_value = next(iter(data.values()))
            if is_dataclass(any_value):
                data = tuple(dict(key=key, **asdict(value)) for key, value in data.items())
            elif isinstance(any_value, Mapping):
                data = tuple(dict(key=key, **value) for key, value in data.items())
            elif isinstance(any_value, Iterable) and not isinstance(any_value, (str, bytes, bytearray)):
                data = tuple(dict(key=key, **{f"col{i}": item for i, item in enumerate(value, 1)}) for key, value in data.items())
            else:
                data = tuple({"key": key, "value": value} for key, value in data.items())
        else:
            if not isinstance(data, Iterable):
                raise TypeError("data must be an iterable")
            if not isinstance(data, Sized):
                data = tuple(data)
        if extra_columns:
            for row in data:
                # missing_cols = row.keys() - columns.keys()  # not order-preserving, so instead:
                missing_cols = [col_name for col_name in (
                    (f.name for f in fields(row)) if is_dataclass(row) else
                    type(row).model_fields.keys() if PYDANTIC_VER >= 2 and isinstance(row, BaseModel) else
                    row.__fields__.keys() if isinstance(row, BaseModel) else
                    row.keys()  # type: ignore
                ) if col_name not in columns]
                for column in missing_cols:
                    columns[column] = {"header": column}
        col_names = columns.keys()
        self.columns_values = tuple(columns.values())
        self.number_of_value_rows = len(data)
        result = worksheet.add_table(0, 0, len(data), len(columns) - 1, {
            "header_row": True,
            "columns": self.columns_values,
            "total_row": total_row,
            **({"name": table_name} if table_name else {}),
            "data": [
                [self.convert_value(row_dict.get(col_name)) for col_name in col_names]  # type: ignore
                for row_union in data
                if (row_dict := (asdict(row_union) if is_dataclass(row_union) else  # type: ignore
                                 row_union.model_dump() if PYDANTIC_VER >= 2 and isinstance(row_union, BaseModel) else
                                 row_union.dict() if isinstance(row_union, BaseModel) else
                                 row_union))
            ],
        })
        if result != 0:
            raise TableAddError(f"Table add error: {result}")
        self.set_column_widths(worksheet, columns, column_width)
        return worksheet

    @classmethod
    def set_column_widths(cls,
                          worksheet: xlsxwriter.worksheet.Worksheet,
                          columns: dict[str, dict[str, str | int | float]],
                          column_width: ColumnWidthType | Mapping[str, ColumnWidthType] | Iterable[ColumnWidthType],
                          ) -> None:
        if column_width is None:
            return
        if column_width == ColumnWidth.AUTOFIT:
            worksheet.autofit()
        elif column_width == ColumnWidth.HIDE:
            worksheet.hide()
        elif isinstance(column_width, (int, float)):
            if column_width == 0:
                raise ValueError("column_width cannot be 0, use None to keep the default width or ColumnWidth.AUTOFIT or ColumnWidth.HIDE")
            if column_width < 0:
                worksheet.set_column(0, len(columns) - 1, -column_width)
            else:
                worksheet.set_column_pixels(0, len(columns) - 1, column_width)
        elif isinstance(column_width, Mapping):
            column_keys = tuple(columns.keys())
            for column_name_or_number, desired_width in column_width.items():
                try:
                    column_number = column_keys.index(column_name_or_number)
                except ValueError:
                    if isinstance(column_name_or_number, int) and column_name_or_number >= 0:
                        column_number = column_name_or_number
                    else:
                        raise
                cls.set_column_width(worksheet, column_number, desired_width)
        elif isinstance(column_width, Iterable):
            for column_number, desired_width in enumerate(column_width):
                cls.set_column_width(worksheet, column_number, desired_width)
        else:
            raise TypeError(f"unsupported column_width type: {column_width!r}")

    @staticmethod
    def set_column_width(worksheet: xlsxwriter.worksheet.Worksheet, column_number: int, column_width: ColumnWidthType,
                         ) -> None:
        if column_width is None:
            return
        if column_width == ColumnWidth.AUTOFIT:
            # dangerous, fragile hack...
            orig_dim_colmin, orig_dim_colmax = worksheet.dim_colmin, worksheet.dim_colmax
            worksheet.dim_colmin, worksheet.dim_colmax = column_number, column_number
            worksheet.autofit()
            worksheet.dim_colmin, worksheet.dim_colmax = orig_dim_colmin, orig_dim_colmax
        elif column_width == ColumnWidth.HIDE:
            worksheet.set_column(column_number, column_number, 8.43, None, {'hidden': 1})
        elif isinstance(column_width, (int, float)):
            if column_width == 0:
                raise ValueError("column_width cannot be 0, use None to keep the default width or ColumnWidth.AUTOFIT or ColumnWidth.HIDE")
            if column_width < 0:
                worksheet.set_column(column_number, column_number, -column_width)
            else:
                worksheet.set_column_pixels(column_number, column_number, column_width)
        else:
            raise TypeError(f"unsupported column_width type: {column_width!r}")

    def column_ref(self, column_name: str, *, absolute: bool = False) -> str:
        column_number = tuple(self.columns.keys()).index(column_name)
        column_letter = xl_col_to_name(column_number, col_abs=absolute)
        return f"{column_letter}:{column_letter}"

    def close(self, filename: str | PathLike | None = None) -> None:
        if filename is not None and self.closed:
            raise WorkbookClosedError
        if filename is not None:
            self.filename = filename
            self.workbook.filename = fspath(filename)
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
            if for_json and isinstance(input_value, (Decimal, Fraction)):
                return float(input_value)
            return input_value
        if input_value is None:
            return None
        if isinstance(input_value, UUID):
            return str(input_value)
        if isinstance(input_value, Mapping):
            return json.dumps(input_value, default=lambda value: cls.convert_value(value, for_json=True))
        if isinstance(input_value, Set):
            return "{" + ", ".join(str(cls.convert_value(value)) for value in input_value) + "}"
        if isinstance(input_value, Iterable):
            return "[" + ", ".join(str(cls.convert_value(value)) for value in input_value) + "]"
        raise UnsupportedTypeError(input_value)


def save_as_xlsx(filename: str | PathLike,
                 data: Iterable[Mapping[str, Any] | DataclassInstance | BaseModel] | None = None,
                 **kwargs) -> None:
    if "auto_save" in kwargs and kwargs["auto_save"] is not None and not kwargs["auto_save"]:
        raise ValueError("calling save_as_xlsx(auto_save=False) makes no sense")
    kwargs["auto_save"] = True
    SaveAsXlsx(filename, data, **kwargs)
