import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.xls_extractor import split_emails, validate_email


@pytest.mark.parametrize("email", [
    "alice@example.com",
    "bob.smith@mail.co.uk",
    "user_name_123@domain.org"
])

def test_validate_email_valid(email):
    assert validate_email(email) is True


@pytest.mark.parametrize("email", [
    "aliceexample.com",  # Missing '@'
    "bob.smith@mail",    # Missing domain extension
    "user@.com",      # Missing domain name
    "a@b.c",         # Too short
    ""               # Empty string
])

def test_validate_email_invalid(email):
    with pytest.raises(ValueError):
        validate_email(email)


def test_split_emails_keeps_unique_addresses_in_order():
    assert split_emails("alice@example.com; bob@example.com") == [
        "alice@example.com",
        "bob@example.com",
    ]


def test_split_emails_keeps_repeated_address_once_on_the_same_row():
    assert split_emails("korpheidi@gmail.com; korpheidi@gmail.com") == [
        "korpheidi@gmail.com"
    ]
    assert split_emails("A@example.com, a@example.com") == ["A@example.com"]


    