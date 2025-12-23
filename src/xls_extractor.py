from subprocess import call
import pandas as pd
import re, unicodedata
from email.utils import parseaddr

from utils.file_utils import get_field

LOCAL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+$")
DOMAIN_RE = re.compile(
    r"^(?=.{1,255}$)(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,}$"
)
RE_NUM = re.compile(r"^\d+$")


class ValidationError(ValueError):
    pass


class Person:
    def __init__(self, apartment, address, email=None):
        self.apartment = apartment
        self.address = address
        self.emails = split_emails(email)

    def __repr__(self):
        return f"Person(\nemails={self.emails}, \naddress={self.address}, \napartment={self.apartment}\n)"


def read_xls_with_fallback(path):
    for enc in ("cp1250", "cp1252", "latin1"):
        try:
            return pd.read_excel(
                path,
                engine="xlrd",
                engine_kwargs={"encoding_override": enc},
            )
        except UnicodeDecodeError:
            continue
    # If all failed, re-raise the last one by reading once without catching
    raise ValidationError(
        f"Ei saa faili {path!r} lugeda. Proovitud kodeeringud: cp1250, cp1252, latin. Palun salvesta fil Excelis ümber vormingusse .xlsx ja proovi uuesti."
    )


def split_emails(email_string: str) -> list[str]:
    """
    Split a string containing one or more emails separated by commas or semicolons.
    Validate each email and return a list of valid emails.
    """
    if not email_string or str(email_string).strip() == "":
        raise ValidationError("Meiliaadress on kohustuslik")
    parts = [part.strip() for part in re.split(r"[;,]", email_string) if part.strip()]
    valid_emails = []
    for part in parts:
        validate_email(part)
        valid_emails.append(part)
    if not valid_emails:
        raise ValidationError("Puuduvad kehtivad meiliaadressid")
    return valid_emails


def validate_email(email: str):
    if not email:
        raise ValidationError("Meil on puudu")

    # Normalize and strip
    norm_email = unicodedata.normalize("NFKC", email).strip()

    # Reject control chars
    if any(ord(c) < 32 for c in norm_email) or "\x7f" in norm_email:
        raise ValidationError(f"Juhtsümbolid pole lubatud! {email!r}")

    _, parsed_email = parseaddr(norm_email)
    if not parsed_email or " " in parsed_email or parsed_email.count("@") != 1:
        raise ValidationError(f"Vigane meiliaadress: {email!r}")

    local, domain = parsed_email.rsplit("@", 1)

    if not LOCAL_RE.match(local):
        raise ValidationError(f"Vigane kasutajanimi: {local!r}")
    if not DOMAIN_RE.match(domain):
        raise ValidationError(f"Vigane domeeninimi: {domain!r}")

    if not "@" in email and not "." in email and len(email) < 5:
        raise ValidationError(f"Vigane meiliaadress: {email!r}")
    return True





def _validate_person_row(row, row_num: int):
    """ Validate a single row of person data from the XLS file. """
    email = get_field(row, "klient_mail")
    apt = get_field(row, "korter")
    yhistu = get_field(row, "yhistu")
    maj_nr = get_field(row, "maj_nr")
    address = f"{yhistu.lower()}, {maj_nr}".strip()

    # --- Row-level checks
    if not RE_NUM.match(apt):
        raise ValidationError(f"Rida {row_num}: korter peab sisaldama ainult numbreid")
    if not email:
        raise ValidationError(f"Rida {row_num}: meiliaadress on kohustuslik")

    # split_emails does validation internally
    split_emails(email)
    return email, apt, address


def extract_person_data(input_file):
    # Required columns
    required = {"klient_mail", "korter", "yhistu", "maj_nr"}
    df = read_xls_with_fallback(input_file)

    # --- Header check
    missing = required - set(df.columns)
    if missing:
        raise ValidationError(
            f"Klientide failist on puudu tulp: {missing}. Palun kontrolli faili õigsust."
        )

    persons = []
    for row_num, row in enumerate(df.itertuples(index=False, name="Row"), start=2):
        email, apt, address = _validate_person_row(row, row_num)
        persons.append(Person(email=email, apartment=apt, address=address))
    if not persons:
        raise ValidationError("Klientide fail ei sisalda ühtegi kehtivat kirjet.")
    return persons
