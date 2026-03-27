# SPDX-FileCopyrightText: 2026-present Vaclav Dvorak <vashek@gmail.com>
#
# SPDX-License-Identifier: MIT
import tempfile
from pathlib import Path
from typing import ClassVar

import pytest  # type: ignore

import save_as_xlsx

from .test_pyopenxl_verifier import verify_using_pyopenxl

try:
    import pydantic  # type: ignore
    BaseModel: type = pydantic.BaseModel
except ImportError:
    pydantic = None
    class BaseModel:  # type: ignore
        def __init__(self, *args, **kwargs):
            pass


def test_print_pydantic_version():
    print("pydantic version: " + getattr(pydantic, "__version__", "(none)") + "\n")


@pytest.mark.skipif(pydantic is None, reason="requires pydantic")
class TestPydantic:
    class ModelForTest(BaseModel):
        a: int
        b: str | None = None

    TEST_DATA_WITH_PYDANTIC: ClassVar[list[ModelForTest]] = [
        ModelForTest(a=1),
        ModelForTest(a=2, b="B"),
    ]

    def test_save_pydantic(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            fn = Path(tmpdir) / "test.xlsx"
            with save_as_xlsx.SaveAsXlsx(fn, self.TEST_DATA_WITH_PYDANTIC) as saver:
                assert len(saver.columns_values) == 2
                assert saver.columns_values[0]["header"] == "a"
                assert saver.columns_values[1]["header"] == "b"
                assert saver.number_of_value_rows == 2
            assert fn.exists()
            verify_using_pyopenxl(fn, "A1:B3", data=[
                ("a", "b"),
                (1, None),
                (2, "B"),
            ])
