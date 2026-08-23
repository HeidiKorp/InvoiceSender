from __future__ import annotations

from configparser import ConfigParser
from pathlib import Path
import sys, os

from src.data_classes import InvoiceType

TEMPLATE_SECTION = "email_template"


def read_config(filename: str = "config.cfg"):
    config = ConfigParser()
    config_path = get_config_path()

    try:
        with config_path.open("r", encoding="utf-8") as f:
            config.read_file(f)
        return config
    except UnicodeDecodeError:
        with config_path.open("r", encoding="cp1252") as f:
            config.read_file(f)
        return config


def get_config_path() -> str:
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
        return Path(os.path.join(base_dir, "_internal", "config.cfg"))
    else:
        return Path(__file__).parent.parent / "config.cfg"


def load_app_name(config):
    return config.get("app", "NAME", fallback="Arvete Saatja")


def load_app_version(config):
    return config.get("app", "VERSION", fallback="1.0.0")


def load_invoice_types(config):
    hint = config.get("ui", "TYPE_HINT")

    def read_section(section: str) -> InvoiceType:
        return InvoiceType(
            key=config.get(section, "KEY"),
            label=config.get(section, "LABEL"),
            subject=config.get(section, "SUBJECT"),
            body=config.get(section, "BODY").replace("\\n", "\n")
        )
    t1 = read_section("invoice_type_kommunaal")
    t2 = read_section("invoice_type_kyte")

    types = {t1.key: t1, t2.key: t2}
    return types, hint


def load_template_config(path: str | Path) -> tuple[str, str]:
    path = Path(path)
    conf = ConfigParser(interpolation=None)
    conf.optionxform = str
    conf.read(path, encoding="utf-8")

    if not conf.has_section(TEMPLATE_SECTION):
        raise ValueError(f"Missing section [{TEMPLATE_SECTION}] in template config")

    subject = conf.get(TEMPLATE_SECTION, "SUBJECT", fallback="")
    body = _decode_body(conf.get(TEMPLATE_SECTION, "BODY", fallback=""))
    return subject, body


def save_template_config(path: str | Path, subject: str, body: str) -> None:
    path = Path(path)
    conf = ConfigParser(interpolation=None)
    conf.optionxform = str

    conf[TEMPLATE_SECTION] = {
        "SUBJECT": subject,
        "BODY": _encode_body(body)
    }

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        conf.write(f, space_around_delimiters=False)


def _decode_body(body: str) -> str:
    return (body or "").replace("\\n", "\n")


def _encode_body(body: str) -> str:
    body = (body or "").rstrip()
    return body.replace("\r\n", "\n").replace("\n", "\\n")
