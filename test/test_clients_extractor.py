from pathlib import Path
import pandas as pd
import pytest
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.clients_extractor import ClientsExtractor
from src.data_classes import ValidationError


def _extractor() -> ClientsExtractor:
    return ClientsExtractor()


@pytest.mark.parametrize(
    "email",
    [
        "alice@example.com",
        "bob.smith@mail.co.uk",
        "user_name_123@domain.org",
    ],
)
def test_validate_email_valid(email):
    assert _extractor().validate_email(email) is True


@pytest.mark.parametrize(
    "email",
    [
        "aliceexample.com",
        "bob.smith@mail",
        "user@.com",
        "a@b.c",
        "",
    ],
)
def test_validate_email_invalid(email):
    with pytest.raises(ValueError):
        _extractor().validate_email(email)


def test_split_emails_keeps_unique_addresses_in_order():
    assert _extractor().split_emails("alice@example.com; bob@example.com") == [
        "alice@example.com",
        "bob@example.com",
    ]


def test_split_emails_keeps_repeated_address_once_on_the_same_row():
    extractor = _extractor()
    assert extractor.split_emails("korpheidi@gmail.com; korpheidi@gmail.com") == [
        "korpheidi@gmail.com"
    ]
    assert extractor.split_emails("A@example.com, a@example.com") == ["A@example.com"]


def test_excel_engine_for_xls_xlsx_xlsm():
    extractor = _extractor()
    assert extractor._excel_engine_for("clients.xls") == "xlrd"
    assert extractor._excel_engine_for("clients.xlsx") == "openpyxl"
    assert extractor._excel_engine_for("clients.xlsm") == "openpyxl"


def test_excel_engine_for_rejects_unknown_suffix():
    with pytest.raises(ValidationError):
        _extractor()._excel_engine_for("clients.csv")


def test_xlsx_read_uses_openpyxl(monkeypatch, tmp_path):
    path = tmp_path / "clients.xlsx"
    path.write_bytes(b"PK")
    seen = {}

    def fake_read_excel(p, engine=None, **kwargs):
        seen["engine"] = engine
        return pd.DataFrame(
            {
                "klient_mail": ["a@example.com"],
                "korter": ["1"],
                "yhistu": ["tee"],
                "maj_nr": ["1"],
            }
        )

    monkeypatch.setattr("src.clients_extractor.pd.read_excel", fake_read_excel)
    _extractor()._read_workbook(path)
    assert seen["engine"] == "openpyxl"
