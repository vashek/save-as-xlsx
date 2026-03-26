# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
import os.path
import tempfile
from dataclasses import dataclass
from decimal import Decimal
from fractions import Fraction
from enum import IntEnum
from pathlib import Path

import save_as_xlsx


class EnumForTest(IntEnum):
    ONE = 1
    TWO = 2


@dataclass
class DataclassForTest:
    a: int
    b: str = None


TEST_DATA = [
    {"a": 1, "b": "2"},
    {"b": "B", "c": None},
]

TEST_DATA_COMPLEX = [
    {"dict": {"a": 1, "b": "2"}},
    {"list": [1, 2, 3]},
    {"set": {1, 2, 3}},
    {"float": -1.5, "dec": Decimal("2.99"), "frac": Fraction(3, 5), "bool": True},
]

TEST_DATA_WITH_ENUM = [
    {"a": 1, "enum": EnumForTest.ONE},
]

TEST_DATA_WITH_DATACLASS = [
    DataclassForTest(a=1),
    DataclassForTest(a=2, b="B"),
]


def test_save_on_explicit_close():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        saver = save_as_xlsx.SaveAsXlsx(TEST_DATA, fn)
        assert not fn.exists()
        saver.close()
        assert fn.exists()

def test_save_with_str_filename():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = os.path.join(tmpdir, "test.xlsx")
        assert isinstance(fn, str)
        saver = save_as_xlsx.SaveAsXlsx(TEST_DATA, fn)
        assert not Path(fn).exists()
        saver.close()
        assert Path(fn).exists()

def test_save_on_auto_close():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, auto_save=True)
        assert fn.exists()

def test_save_on_with():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn):
            assert not fn.exists()
        assert fn.exists()

def test_num_rows():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn) as saver:
            assert saver.number_of_value_rows == 2

def test_default_column_order():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn) as saver:
            assert len(saver.columns_values) == 3
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "b"
            assert saver.columns_values[2]["header"] == "c"

def test_column_order():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b",)) as saver:
            assert len(saver.columns_values) == 3
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "a"
            assert saver.columns_values[2]["header"] == "c"

def test_column_order_extra_nonexistent_column():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b", "nonexistent")) as saver:
            assert len(saver.columns_values) == 4
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "nonexistent"
            assert saver.columns_values[2]["header"] == "a"
            assert saver.columns_values[3]["header"] == "c"

def test_column_order_no_extras():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b", "c"), extra_columns=False) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "c"

def test_column_order_extra_nonexistent_column_no_extras():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b", "nonexistent"), extra_columns=False) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "nonexistent"

def test_generator():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(({"num": i} for i in range(5)), fn) as saver:
            assert len(saver.columns_values) == 1
            assert saver.columns_values[0]["header"] == "num"
            assert saver.number_of_value_rows == 5

def test_enum():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA_WITH_ENUM, fn) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "enum"

def test_complex():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA_COMPLEX, fn) as saver:
            assert len(saver.columns_values) == 7
            assert saver.number_of_value_rows == len(TEST_DATA_COMPLEX)

def test_dataclasses():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA_WITH_DATACLASS, fn) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "b"
            assert saver.number_of_value_rows == 2
