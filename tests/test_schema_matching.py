"""Tests for src.schema_matching.matcher — placeholder suite."""

import pytest
import pandas as pd

from src.schema_matching.matcher import match_columns


def test_match_columns_raises_not_implemented():
    frames = {"Sheet1": pd.DataFrame({"revenue": [1], "cost": [2]})}
    with pytest.raises(NotImplementedError):
        match_columns(frames, frames)
