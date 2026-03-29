# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
import os.path
import tempfile
from dataclasses import dataclass
from decimal import Decimal
try:
    from enum import IntEnum, StrEnum
except ImportError:
    from enum import IntEnum as IntEnum, Enum as StrEnum  # type: ignore
from fractions import Fraction
from pathlib import Path
from uuid import UUID

import save_as_xlsx

from .test_pyopenxl_verifier import verify_using_pyopenxl


class IntEnumForTest(IntEnum):
    ONE = 1
    TWO = 2


class StrEnumForTest(StrEnum):
    MALE = "M"
    FEMALE = "F"


@dataclass
class PersonDataclassForTest:
    age: int
    name: str | None = None
    sex: StrEnumForTest | str | None = None


TEST_DATA = [
    {"a": 1, "b": "2"},
    {"b": "B", "c": None},
]

TEST_DATA_COMPLEX = [
    {"dict": {"a": 1, "b": "2"}},
    {"list": [1, 2, 3]},
    {"set": {1, 2, 3}},
    {"float": -1.5, "dec": Decimal("2.99"), "frac": Fraction(3, 5), "bool": True},
    {"dict": {"nested": {"enum": IntEnumForTest.ONE, "dec": Decimal("2.99")}}},
]

TEST_DATA_WITH_ENUM = [
    {"a": 1, "enum": IntEnumForTest.ONE},
]

TEST_DATA_WITH_UUID = [
    {"a": 1, "uuid": UUID("5f456a18-29f0-11f1-a203-e41fd5b9abcb")},
]

TEST_DATA_WITH_DATACLASS = [
    PersonDataclassForTest(age=1),
    PersonDataclassForTest(age=2, name="B"),
]

TEST_DICT_SIMPLE = {
    "John": "male",
    "Jane": StrEnumForTest.FEMALE,
}

TEST_DICT_WITH_DICTS = {
    "John": {"Age": 69, "Sex": "male"},
    "Jane": {"Age": 42, "Sex": StrEnumForTest.FEMALE},
}

TEST_DICT_WITH_LISTS = {
    "John": [69, "male"],
    "Jane": (42, StrEnumForTest.FEMALE),
}

TEST_DICT_WITH_DATACLASSES = {
    "John": PersonDataclassForTest(age=69, sex="male"),
    "Jane": PersonDataclassForTest(age=42, sex=StrEnumForTest.FEMALE),
}


def test_save_with_function():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        assert not fn.exists()
        save_as_xlsx.save_as_xlsx(fn, TEST_DATA)
        assert fn.exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_on_explicit_close():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        saver = save_as_xlsx.SaveAsXlsx(fn, TEST_DATA)
        assert not fn.exists()
        saver.close()
        assert fn.exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_with_str_filename():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = os.path.join(tmpdir, "test.xlsx")
        assert isinstance(fn, str)
        saver = save_as_xlsx.SaveAsXlsx(fn, TEST_DATA)
        assert not Path(fn).exists()
        saver.close()
        assert Path(fn).exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_with_new_filename():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        fn2 = Path(tmpdir) / "test2.xlsx"
        saver = save_as_xlsx.SaveAsXlsx(fn, TEST_DATA)
        assert not fn.exists()
        assert not fn2.exists()
        saver.close(fn2)
        assert not fn.exists()
        assert fn2.exists()
        verify_using_pyopenxl(fn2, "A1:C3")

def test_save_on_auto_close():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, auto_save=True)
        assert fn.exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_save_on_with():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA):
            assert not fn.exists()
        assert fn.exists()
        verify_using_pyopenxl(fn, "A1:C3")

def test_num_rows_and_sheet_name():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA) as saver:
            assert saver.number_of_value_rows == 2
            assert saver.worksheet.name == "Sheet1"
        verify_using_pyopenxl(fn, "A1:C3")

def test_custom_sheet_name():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, sheet_name="My Sheet") as saver:
            assert saver.worksheet.name == "My Sheet"
        verify_using_pyopenxl(fn, "A1:C3", sheet_name="My Sheet")

def test_default_column_order():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA) as saver:
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
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, column_order=("b",)) as saver:
            assert len(saver.columns_values) == 3
            assert saver.columns_values[0]["header"] == "b"
            assert saver.columns_values[1]["header"] == "a"
            assert saver.columns_values[2]["header"] == "c"
            assert saver.column_ref("a") == "B:B"
            assert saver.column_ref("b", absolute=True) == "$A:$A"
        verify_using_pyopenxl(fn, "A1:C3", data=[
            ("b", "a", "c"),
            ("2", 1, None),
            ("B", None, None),
        ])

def test_column_order_extra_nonexistent_column():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, column_order=("b", "nonexistent")) as saver:
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
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, column_order=("b", "c"), extra_columns=False) as saver:
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
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, column_order=("b", "nonexistent"), extra_columns=False) as saver:
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
        with save_as_xlsx.SaveAsXlsx(fn, ({"num": i} for i in range(5))) as saver:
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
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA_WITH_ENUM) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "enum"
        verify_using_pyopenxl(fn, "A1:B2", data=[
            ("a", "enum"),
            (1, "ONE"),
        ])

def test_uuid():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA_WITH_UUID) as saver:
            assert len(saver.columns_values) == 2
            assert saver.columns_values[0]["header"] == "a"
            assert saver.columns_values[1]["header"] == "uuid"
        verify_using_pyopenxl(fn, "A1:B2", data=[
            ("a", "uuid"),
            (1, "5f456a18-29f0-11f1-a203-e41fd5b9abcb"),
        ])

def test_complex():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA_COMPLEX) as saver:
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
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA_WITH_DATACLASS) as saver:
            assert len(saver.columns_values) == 3
            assert saver.columns_values[0]["header"] == "age"
            assert saver.columns_values[1]["header"] == "name"
            assert saver.columns_values[2]["header"] == "sex"
            assert saver.number_of_value_rows == 2
        verify_using_pyopenxl(fn, "A1:C3", data=[
            ("age", "name", "sex"),
            (1, None, None),
            (2, "B", None),
        ])

def test_another_sheet():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DATA, sheet_name="FirstSheet", table_name="FirstTable") as saver:
            saver.add_sheet(TEST_DATA_WITH_ENUM, sheet_name="AnotherSheet", table_name="AnotherTable")
        verify_using_pyopenxl(fn, sheet_name="FirstSheet", table_name="FirstTable", dimensions="A1:C3", data=[
            ("a", "b", "c"),
            (1, "2", None),
            (None, "B", None),
        ])
        verify_using_pyopenxl(fn, sheet_name="AnotherSheet", table_name="AnotherTable", dimensions="A1:B2", data=[
            ("a", "enum"),
            (1, "ONE"),
        ])

def test_empty_then_two_sheets():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn) as saver:
            saver.add_sheet(TEST_DATA, sheet_name="FirstSheet", table_name="FirstTable")
            saver.add_sheet(TEST_DATA_WITH_ENUM, sheet_name="AnotherSheet", table_name="AnotherTable")
        verify_using_pyopenxl(fn, sheet_name="FirstSheet", table_name="FirstTable", dimensions="A1:C3", data=[
            ("a", "b", "c"),
            (1, "2", None),
            (None, "B", None),
        ])
        verify_using_pyopenxl(fn, sheet_name="AnotherSheet", table_name="AnotherTable", dimensions="A1:B2", data=[
            ("a", "enum"),
            (1, "ONE"),
        ])

def test_dict_simple():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DICT_SIMPLE) as saver:
            assert saver.number_of_value_rows == len(TEST_DICT_SIMPLE)
            assert len(saver.columns) == 2
            col_keys = tuple(saver.columns.keys())
            assert col_keys == ("key", "value")
        verify_using_pyopenxl(fn, dimensions="A1:B3", data=[
            ("key", "value"),
            ("John", "male"),
            ("Jane", "FEMALE"),
        ])

def test_dict_with_dicts():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DICT_WITH_DICTS) as saver:
            assert saver.number_of_value_rows == len(TEST_DICT_WITH_DICTS)
            assert len(saver.columns) == 3
            col_keys = tuple(saver.columns.keys())
            assert col_keys == ("key", "Age", "Sex")
        verify_using_pyopenxl(fn, dimensions="A1:C3", data=[
            ("key", "Age", "Sex"),
            ("John", 69, "male"),
            ("Jane", 42, "FEMALE"),
        ])

def test_dict_with_lists():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DICT_WITH_LISTS) as saver:
            assert saver.number_of_value_rows == len(TEST_DICT_WITH_LISTS)
            assert len(saver.columns) == 3
            col_keys = tuple(saver.columns.keys())
            assert col_keys == ("key", "col1", "col2")
        verify_using_pyopenxl(fn, dimensions="A1:C3", data=[
            ("key", "col1", "col2"),
            ("John", 69, "male"),
            ("Jane", 42, "FEMALE"),
        ])

def test_dict_with_dataclasses():
    with tempfile.TemporaryDirectory() as tmpdir:
        fn = Path(tmpdir) / "test.xlsx"
        with save_as_xlsx.SaveAsXlsx(fn, TEST_DICT_WITH_DATACLASSES) as saver:
            assert saver.number_of_value_rows == len(TEST_DICT_WITH_DATACLASSES)
            assert len(saver.columns) == 4
            col_keys = tuple(saver.columns.keys())
            assert col_keys == ("key", "age", "name", "sex")
        verify_using_pyopenxl(fn, dimensions="A1:D3", data=[
            ("key", "age", "name", "sex"),
            ("John", 69, None, "male"),
            ("Jane", 42, None, "FEMALE"),
        ])
