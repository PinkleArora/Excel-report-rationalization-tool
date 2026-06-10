"""Tests for src.reconciliation.reconciler — placeholder suite."""

import pytest
import pandas as pd

from src.reconciliation.reconciler import reconcile


def test_reconcile_raises_not_implemented():
    df = pd.DataFrame({"value": [10.0, 20.0]})
    with pytest.raises(NotImplementedError):
        reconcile([df], df, numeric_columns=["value"])
