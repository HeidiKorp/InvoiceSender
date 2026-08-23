from pathlib import Path
import pandas as pd
import re, unicodedata
from email.utils import parseaddr

from utils.file_utils import get_field
from src.data_classes import Person, ValidationError

LOCAL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+$")
DOMAIN_RE = re.compile(
    r"^(?=.{1,255}$)(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,}$"
)
RE_NUM = re.compile(r"^\d+$")
SUPPORTED_CLIENT_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
REQUIRED_COLUMNS = {"klient_mail", "korter", "yhistu", "maj_nr"}


class ClientsExtractor:
    def extract(self, input_file) -> list[Person]:
        df = self._read_workbook(input_file)

        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            raise ValidationError(
                f"Klientide failist on puudu tulp: {missing}. Palun kontrolli faili õigsust."
            )

        persons = []
        for row_num, row in enumerate(df.itertuples(index=False, name="Row"), start=2):
            email, apt, address = self._validate_person_row(row, row_num)
            emails = self.split_emails(email)
            persons.append(Person(emails=emails, apartment=apt, address=address))
        if not persons:
            raise ValidationError("Klientide fail ei sisalda ühtegi kehtivat kirjet.")
        return persons

    def split_emails(self, email_string: str) -> list[str]:
        if not email_string or str(email_string).strip() == "":
            raise ValidationError("Meiliaadress on kohustuslik")
        parts = [part.strip() for part in re.split(r"[;,]", email_string) if part.strip()]
        valid_emails = []
        seen = set()
        for part in parts:
            self.validate_email(part)
            key = unicodedata.normalize("NFKC", part).strip().casefold()
            if key in seen:
                continue
            seen.add(key)
            valid_emails.append(part)
        if not valid_emails:
            raise ValidationError("Puuduvad kehtivad meiliaadressid")
        return valid_emails

    def validate_email(self, email: str):
        if not email:
            raise ValidationError("Meil on puudu")

        norm_email = unicodedata.normalize("NFKC", email).strip()

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

        return True

    def _validate_person_row(self, row, row_num: int):
        email = get_field(row, "klient_mail")
        apt = get_field(row, "korter")
        yhistu = get_field(row, "yhistu")
        maj_nr = get_field(row, "maj_nr")
        address = f"{yhistu.lower()}, {maj_nr}".strip()

        if not RE_NUM.match(apt):
            raise ValidationError(f"Rida {row_num}: korter peab sisaldama ainult numbreid")
        if not email:
            raise ValidationError(f"Rida {row_num}: meiliaadress on kohustuslik")

        self.split_emails(email)
        return email, apt, address

    def _read_workbook(self, path):
        engine = self._excel_engine_for(path)
        if engine == "xlrd":
            return self._read_xls_with_fallback(path)
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception as e:
            raise ValidationError(
                f"Ei saa faili {path!r} lugeda ({e}). Palun kontrolli, et fail oleks korrektne Exceli fail."
            ) from e

    def _excel_engine_for(self, path) -> str:
        suffix = Path(path).suffix.lower()
        if suffix == ".xls":
            return "xlrd"
        if suffix in {".xlsx", ".xlsm"}:
            return "openpyxl"
        raise ValidationError(
            f"Klientide faili vorming {suffix!r} ei ole toetatud. Kasuta .xls, .xlsx või .xlsm."
        )

    def _read_xls_with_fallback(self, path):
        for enc in ("cp1250", "cp1252", "latin1"):
            try:
                return pd.read_excel(
                    path,
                    engine="xlrd",
                    engine_kwargs={"encoding_override": enc},
                )
            except UnicodeDecodeError:
                continue
        raise ValidationError(
            f"Ei saa faili {path!r} lugeda. Proovitud kodeeringud: cp1250, cp1252, latin. "
            "Palun salvesta fail Excelis ümber vormingusse .xlsx ja proovi uuesti."
        )
