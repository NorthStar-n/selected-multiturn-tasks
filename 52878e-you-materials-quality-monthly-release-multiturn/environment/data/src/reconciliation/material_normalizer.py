"""Normalize copper material labels used by the July qualification workpapers."""

from __future__ import annotations

ALIASES = {
    "C10100 OFHC": "MAT-CU-C101",
    "E-Cu58": "MAT-CU-WIRE",
    "HB electro-Cu": "MAT-CU-HBPLATE",
    "Cu^1O": "MAT-CU2O-CUPRITE",
    "Cu^2O": "MAT-CUO-TENORITE",
    "cupric sulfate": "MAT-CUSO4-CUPRIC",
}


def normalize_label(label: str) -> str:
    """Return the material identity used by the source register."""
    clean = " ".join(label.strip().split())
    if clean in ALIASES:
        return ALIASES[clean]
    return clean.upper().replace(" ", "-")


def is_july_scope(receipt_text: str, locale_hint: str) -> bool:
    """Recognize ISO dates and the Lisbon supplier export date convention."""
    if not receipt_text:
        return False
    if "Lisbon" in locale_hint and "/" in receipt_text:
        day, month, year = receipt_text.split("/")
        return year == "2026" and month == "07"
    return receipt_text.startswith("2026-07")
