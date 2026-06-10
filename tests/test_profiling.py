"""Tests for src.profiling.profiler — placeholder suite."""

import pytest
import pandas as pd

from src.profiling.profiler import profile_sheet


def test_profile_sheet_raises_not_implemented():
    df = pd.DataFrame({"a": [1, 2, None], "b": ["x", "y", "z"]})
    with pytest.raises(NotImplementedError):
        profile_sheet(df, sheet_name="TestSheet")
