# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
import os.path
import tempfile
from dataclasses import dataclass
from decimal import Decimal
from enum import IntEnum
from fractions import Fraction
from pathlib import Path

import save_as_xlsx

from .test_pyopenxl_verifier import verify_using_pyopenxl


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
    {"dict": {"nested": {"enum": EnumForTest.ONE, "dec": Decimal("2.99")}}},
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
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_with_str_filename():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = os.path.join(tmpdir, "test.xlsx")
        assert isinstance(fn, str)
        saver = save_as_xlsx.SaveAsXlsx(TEST_DATA, fn)
        assert not Path(fn).exists()
        saver.close()
        assert Path(fn).exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_on_auto_close():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, auto_save=True)
        assert fn.exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_on_with():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn):
            assert not fn.exists()
        assert fn.exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_num_rows():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn) as saver:
            assert saver.number_of_value_rows == 2
        verify_using_pyopenxl(fn, "A1:C3")

def test_default_column_order():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn) as saver:
            assert len(saver.columns_values) == 3
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "b"
            assert saver.columns_values[2]["header"] == "c"
        verify_using_pyopenxl(fn, "A1:C3", data=[
            ("a", "b", "c"),
            (1, "2", None),
            (None, "B", None),
        ])

def test_column_order():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b",)) as saver:
            assert len(saver.columns_values) == 3
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "a"
            assert saver.columns_values[2]["header"] == "c"
        verify_using_pyopenxl(fn, "A1:C3", data=[
            ("b", "a", "c"),
            ("2", 1, None),
            ("B", None, None),
        ])

def test_column_order_extra_nonexistent_column():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b", "nonexistent")) as saver:
            assert len(saver.columns_values) == 4
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "nonexistent"
            assert saver.columns_values[2]["header"] == "a"
            assert saver.columns_values[3]["header"] == "c"
        verify_using_pyopenxl(fn, "A1:D3", data=[
            ("b", "nonexistent", "a", "c"),
            ("2", None, 1, None),
            ("B", None, None, None),
        ])

def test_column_order_no_extras():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b", "c"), extra_columns=False) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "c"
        verify_using_pyopenxl(fn, "A1:B3", data=[
            ("b", "c"),
            ("2", None),
            ("B", None),
        ])

def test_column_order_extra_nonexistent_column_no_extras():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA, fn, column_order=("b", "nonexistent"), extra_columns=False) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "nonexistent"
        verify_using_pyopenxl(fn, "A1:B3", data=[
            ("b", "nonexistent"),
            ("2", None),
            ("B", None),
        ])

def test_generator():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(({"num": i} for i in range(5)), fn) as saver:
            assert len(saver.columns_values) == 1
            assert saver.columns_values[0]["header"] == "num"
            assert saver.number_of_value_rows == 5
        verify_using_pyopenxl(fn, "A1:A6", data=[
            ("num",),
            (0,),
            (1,),
            (2,),
            (3,),
            (4,),
        ])

def test_enum():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA_WITH_ENUM, fn) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "enum"
        verify_using_pyopenxl(fn, "A1:B2", data=[
            ("a", "enum"),
            (1, "ONE"),
        ])

def test_complex():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA_COMPLEX, fn) as saver:
            assert len(saver.columns_values) == 7
            assert saver.number_of_value_rows == len(TEST_DATA_COMPLEX)
        verify_using_pyopenxl(fn, "A1:G6", data=[
            ("dict", "list", "set", "float", "dec", "frac", "bool"),
            ('{"a": 1, "b": "2"}', None, None, None, None, None, None),
            (None, '[1, 2, 3]', None, None, None, None, None),
            (None, None, '{1, 2, 3}', None, None, None, None),
            (None, None, None, -1.5, 2.99, 0.6, True),
            # TODO: XXX BUG enum should be "ONE"
            ('{"nested": {"enum": 1, "dec": 2.99}}', None, None, None, None, None, None),
        ])

def test_dataclasses():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(TEST_DATA_WITH_DATACLASS, fn) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "b"
            assert saver.number_of_value_rows == 2
        verify_using_pyopenxl(fn, "A1:B3", data=[
            ("a", "b"),
            (1, None),
            (2, "B"),
        ])
