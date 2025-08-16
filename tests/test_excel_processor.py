import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.append(str(ROOT))

from excel_processor import lookup_account_codes


def test_lookup_account_codes_success():
    bank, cari = lookup_account_codes("233442112")
    assert bank == "6293986"
    assert cari == "120.12.001"


def test_lookup_account_codes_missing():
    with pytest.raises(ValueError):
        lookup_account_codes("000000000")
