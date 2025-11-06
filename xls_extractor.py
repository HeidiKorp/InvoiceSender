from subprocess import call
import pandas as pd
import re, unicodedata
from email.utils import parseaddr

LOCAL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+$")
DOMAIN_RE = re.compile(r"^(?=.{1,255}$)(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,}$")
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
    raise ValidationError(f"Ei saa faili {path!r} lugeda. Proovitud kodeeringud: cp1250, cp1252, latin. Palun salvesta fil Excelis ümber vormingusse .xlsx ja proovi uuesti.")


def split_emails(email_string):
    emails = []
    if email_string:
        emails = [email.strip() for email in re.split(r'[;,]', email_string) if email.strip()]
        emails = [email for email in emails if validate_email(email)]
    else:
        raise ValueError("Email is required")
    return emails


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


def extract_person_data(input_file):
    # Required columns
    required = {"klient_mail", "korter", "yhistu", "maj_nr"}

    df = read_xls_with_fallback(input_file)

    # --- Header check
    missing = required - set(df.columns)
    if missing:
        raise ValidationError("Klientide failist on puudu tulp: {missing}. Palun kontrolli faili õigsust.")

    print(f'Getting her~!')
    persons = []
    for _, row in df.iterrows():
        email = str(row['klient_mail']).strip()
        apt = str(row['korter']).strip()
        address = str(row['yhistu']).strip().lower() + " " + str(row['maj_nr']).strip()

        print(f'Processing row: email={email}, apt={apt}, address={address}')

        # --- Row-level checks
        if not RE_NUM.match(apt):
            raise ValidationError("Rida {i+2}: korter peab sisaldama ainult numbreid")

        person = Person(
            email = email,
            apartment = apt,
            address = address
        )
        persons.append(person)
    return persons