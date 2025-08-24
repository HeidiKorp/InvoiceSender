import pytest
from xls_extractor import validate_email, split_emails


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


    