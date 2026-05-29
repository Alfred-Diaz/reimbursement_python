import re
import unicodedata
from decimal import Decimal, InvalidOperation

from constants import COLUMNS, HEADER_ALIASES, BANK_ALIASES, IGNORED_COLUMNS


def fix_mojibake(value):
    if value is None:
        return ""
    text = str(value)
    if "Ã" in text or "Â" in text:
        try:
            text = text.encode("latin1").decode("utf-8")
        except Exception:
            pass
    return text


def remove_accents(value):
    text = fix_mojibake(value)
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )


def normalize_header(value):
    text = remove_accents(value).replace("\xa0", " ").replace("\n", " ")
    text = re.sub(r"\([^)]*\)", "", text)
    return re.sub(r"[^A-Za-z0-9]+", " ", text).strip().lower()


def is_blank(value):
    return value is None or str(value).strip() == ""


def clean_text(value, max_len=None):
    if is_blank(value):
        return ""
    text = remove_accents(value).upper().strip()
    text = re.sub(r"[^A-Z0-9 ]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len].rstrip() if max_len else text


def clean_account(value):
    if is_blank(value):
        return ""
    text = str(value).strip().replace(",", "")
    try:
        if re.fullmatch(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+", text):
            return str(Decimal(text).quantize(Decimal("1")))
        if re.fullmatch(r"\d+\.0+", text):
            return text.split(".")[0]
    except (InvalidOperation, ValueError):
        pass
    return re.sub(r"\D+", "", text)


def clean_amount(value):
    if is_blank(value):
        return ""
    try:
        return float(Decimal(str(value).replace(",", "").strip()))
    except Exception:
        return value


def canonical_lookup():
    lookup = {normalize_header(column): column for column in COLUMNS}
    lookup.update(HEADER_ALIASES)
    return lookup


def canonical_bank(value):
    name = clean_text(value)
    return clean_text(BANK_ALIASES.get(name, name))


def map_row_keys(row):
    lookup = canonical_lookup()
    mapped = {column: "" for column in COLUMNS}
    unmatched = []

    for key, value in row.items():
        normalized = normalize_header(key)

        if normalized in lookup:
            mapped[lookup[normalized]] = value
        elif str(key).strip() and normalized not in IGNORED_COLUMNS:
            unmatched.append(str(key))

    return mapped, unmatched
