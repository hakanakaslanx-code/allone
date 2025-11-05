"""Helper utilities for the Rinven Import Sheet Generator module."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

RINVEN_IMPORT_COLUMNS: List[str] = [
    "RugNo",
    "UPC",
    "RollNo",
    "VRugNo",
    "Vcollection",
    "Collection",
    "VDesign",
    "Design",
    "Brandname",
    "Ground",
    "Border",
    "ASize",
    "StSize",
    "Area",
    "type",
    "Rate",
    "Amount",
    "Shape",
    "Style",
    "ImageFileName",
    "Origin",
    "Retail",
    "SP",
    "MSRP",
    "Cost",
]

DEFAULT_PRICING: Dict[str, float] = {
    "sp_multiplier": 1.5,
    "retail_multiplier": 2.0,
    "msrp_multiplier": 2.4,
}

_DATA_FILE = Path(__file__).resolve().parent / "rinven_import_data.json"


def _round_nearest(value: float) -> int:
    """Round a floating point value to the nearest integer using half-up rules."""
    if value >= 0:
        return int(value + 0.5)
    return int(value - 0.5)


def make_empty_row(rugno: str = "") -> Dict[str, str]:
    """Return a fresh row dictionary with the expected column structure."""
    row = {column: "" for column in RINVEN_IMPORT_COLUMNS}
    if rugno:
        row["RugNo"] = rugno
    return row


def ensure_row_structure(row: Dict[str, object]) -> Dict[str, str]:
    """Ensure a row dictionary contains all required columns as strings."""
    sanitized: Dict[str, str] = {}
    for column in RINVEN_IMPORT_COLUMNS:
        value = row.get(column, "") if isinstance(row, dict) else ""
        if value is None:
            sanitized[column] = ""
        elif isinstance(value, str):
            sanitized[column] = value
        else:
            sanitized[column] = str(value)
    return sanitized


def ensure_rows(rows: Iterable[Dict[str, object]]) -> List[Dict[str, str]]:
    """Normalise a sequence of rows to the expected structure."""
    return [ensure_row_structure(row) for row in rows]


class RinvenImportStorage:
    """Simple JSON based persistence layer for the Rinven import grid."""

    def __init__(self, path: Optional[Path] = None):
        self.path = Path(path) if path else _DATA_FILE

    def load_rows(self) -> List[Dict[str, str]]:
        if not self.path.exists():
            return []
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return []
        if isinstance(data, list):
            return ensure_rows(row for row in data if isinstance(row, dict))
        return []

    def save_rows(self, rows: Iterable[Dict[str, object]]) -> None:
        serialisable = [ensure_row_structure(row) for row in rows]
        self.path.write_text(json.dumps(serialisable, indent=2, ensure_ascii=False), encoding="utf-8")


def _parse_measurement(text: str) -> Optional[float]:
    value = text.strip().lower()
    if not value:
        return None
    replacements = {
        "″": '"',
        "“": '"',
        "”": '"',
        "’": "'",
        "´": "'",
        "`": "'",
        " in": '"',
        "inch": '"',
        "inches": '"',
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = value.replace("\u00d7", "x")
    value = value.replace(" by ", " x ")

    if "'" in value:
        parts = value.split("'", 1)
        feet_text = parts[0].strip()
        remainder = parts[1].strip()
        feet = float(feet_text) if feet_text else 0.0
        inches = 0.0
        if '"' in remainder:
            inch_text = remainder.split('"', 1)[0].strip()
            inches = float(inch_text) if inch_text else 0.0
        elif remainder:
            try:
                inches = float(remainder)
            except ValueError:
                inches = 0.0
        return feet + inches / 12.0

    if '"' in value:
        inch_text = value.replace('"', '').strip()
        try:
            inches = float(inch_text)
        except ValueError:
            return None
        return inches / 12.0

    try:
        return float(value)
    except ValueError:
        return None


def parse_size_text(size_text: str) -> Optional[Tuple[float, float]]:
    """Parse a free-form size string into width and length in feet."""
    if not size_text:
        return None
    tokens = [token.strip() for token in size_text.lower().replace("\u00d7", "x").split("x")]
    tokens = [token for token in tokens if token]
    if len(tokens) < 2:
        return None
    width = _parse_measurement(tokens[0])
    length = _parse_measurement(tokens[1])
    if width is None or length is None:
        return None
    return width, length


def normalise_size(size_text: str) -> Tuple[str, str]:
    """Return the normalised StSize string and area text for a raw size input."""
    parsed = parse_size_text(size_text)
    if not parsed:
        return "", ""
    width, length = parsed
    if width <= 0 or length <= 0:
        return "", ""
    rounded_width = max(1, _round_nearest(width))
    rounded_length = max(1, _round_nearest(length))
    st_size = f"{rounded_width}x{rounded_length}"
    area_value = round(width * length, 2)
    area_text = f"{area_value:.2f}"
    return st_size, area_text


def apply_pricing(cost_text: str, pricing: Dict[str, float]) -> Dict[str, str]:
    """Calculate pricing related fields based on the raw cost string."""
    cost_text = cost_text.strip()
    if not cost_text:
        return {"Rate": "", "Amount": "", "SP": "", "Retail": "", "MSRP": "", "Cost": ""}
    try:
        cost_value = float(cost_text)
    except ValueError:
        raise ValueError("invalid_cost") from None
    formatted_cost = f"{cost_value:.2f}"
    rate_value = _round_nearest(cost_value)
    results = {
        "Cost": formatted_cost,
        "Rate": str(rate_value),
        "Amount": str(rate_value),
    }
    for key, column in (
        ("sp_multiplier", "SP"),
        ("retail_multiplier", "Retail"),
        ("msrp_multiplier", "MSRP"),
    ):
        multiplier = pricing.get(key, DEFAULT_PRICING[key])
        price_value = _round_nearest(cost_value * float(multiplier))
        results[column] = str(price_value)
    return results

__all__ = [
    "RINVEN_IMPORT_COLUMNS",
    "DEFAULT_PRICING",
    "RinvenImportStorage",
    "apply_pricing",
    "ensure_row_structure",
    "ensure_rows",
    "make_empty_row",
    "normalise_size",
]
