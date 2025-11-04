"""Utility for mapping arbitrary data sources to Wayfair export columns.

This module provides a flexible :func:`map_wayfair_fields` helper that accepts
any :class:`pandas.DataFrame` and attempts to map its columns to the Wayfair
export schema.  The mapper performs fuzzy matching against a curated synonym
list, optionally prompting the user for clarification when a match is
uncertain.  Mappings can be persisted to ``mappings.yaml`` so future runs can
apply the same selections automatically.

Only the columns that can be matched are populated; all other Wayfair columns
remain blank in the returned DataFrame.  No validation or error reporting is
performed—issues are logged to ``mapper.log`` for later inspection.
"""

from __future__ import annotations

import difflib
import logging
import re
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
import yaml

__all__ = [
    "map_wayfair_fields",
    "suggest_mappings",
    "apply_mappings",
    "save_mapping_yaml",
]

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

LOGGER_NAME = "wayfair_field_mapper"
LOG_PATH = Path("mapper.log")
DEFAULT_MAPPING_PATH = Path("mappings.yaml")
AUTO_ACCEPT_THRESHOLD = 0.86
PROMPT_THRESHOLD = 0.6

WAYFAIR_TARGETS: List[str] = [
    "core::supplierPartNumber",
    "core::universalProductCode",
    "core::productName",
    "core::collectionName",
    "price::wholesalePrice",
    "price::minimumAdvertizedPrice",
    "price::manufacturerSuggestedRetailPrice",
    "featureDescription::romanceCopy",
    "featureDescription::genericFeatures",
    "featureDescription::genericFeatures.1",
    "featureDescription::genericFeatures.2",
    "featureDescription::genericFeatures.3",
    "featureDescription::genericFeatures.4",
    "shippingAndFulfillment::weight",
    "shippingAndFulfillment::weight.1",
    "shippingAndFulfillment::height.1",
    "shippingAndFulfillment::width.1",
    "shippingAndFulfillment::depth.1",
    "shippingAndFulfillment::shipType",
    "shippingAndFulfillment::freightClass",
    "core::countryOfManufacturer",
    "media::image1",
    "media::image2",
    "media::image3",
    "media::image4",
    "media::image5",
    "attr::shape",
    "attr::material",
    "attr::technique",
    "attr::color",
    "attr::careInstructions",
    "propSixtyFive::warningRequired",
]

IMAGE_TARGETS = [
    "media::image1",
    "media::image2",
    "media::image3",
    "media::image4",
    "media::image5",
]

WAYFAIR_SYNONYMS: Mapping[str, Sequence[str]] = {
    "core::supplierPartNumber": [
        "sku",
        "rugno",
        "rug no",
        "item number",
        "supplier part number",
        "product code",
    ],
    "core::universalProductCode": ["upc", "barcode", "bar code", "ean"],
    "core::productName": ["product name", "title", "name"],
    "core::collectionName": ["collection", "collection name"],
    "price::wholesalePrice": ["wholesale", "wholesale price", "base cost", "base price"],
    "price::minimumAdvertizedPrice": ["map", "minimum advertised price", "minimum advertised"],
    "price::manufacturerSuggestedRetailPrice": ["msrp", "suggested retail", "retail price"],
    "featureDescription::romanceCopy": [
        "description",
        "longdesc",
        "long description",
        "romance copy",
        "product description",
    ],
    "featureDescription::genericFeatures": [
        "feature bullet 1",
        "feature 1",
        "bullet 1",
        "key feature 1",
    ],
    "featureDescription::genericFeatures.1": [
        "feature bullet 2",
        "feature 2",
        "bullet 2",
        "key feature 2",
    ],
    "featureDescription::genericFeatures.2": [
        "feature bullet 3",
        "feature 3",
        "bullet 3",
        "key feature 3",
    ],
    "featureDescription::genericFeatures.3": [
        "feature bullet 4",
        "feature 4",
        "bullet 4",
        "key feature 4",
    ],
    "featureDescription::genericFeatures.4": [
        "feature bullet 5",
        "feature 5",
        "bullet 5",
        "key feature 5",
    ],
    "shippingAndFulfillment::weight": [
        "product weight",
        "item weight",
        "weight",
        "weight lbs",
    ],
    "shippingAndFulfillment::weight.1": [
        "carton 1 weight",
        "carton weight",
        "box weight",
        "package weight",
    ],
    "shippingAndFulfillment::height.1": [
        "carton 1 height",
        "carton height",
        "box height",
        "package height",
    ],
    "shippingAndFulfillment::width.1": [
        "carton 1 width",
        "carton width",
        "box width",
        "package width",
    ],
    "shippingAndFulfillment::depth.1": [
        "carton 1 depth",
        "carton 1 length",
        "carton depth",
        "box length",
        "package depth",
    ],
    "shippingAndFulfillment::shipType": ["ship type", "shipping type", "ship method"],
    "shippingAndFulfillment::freightClass": ["freight class", "freight", "nmfc"],
    "core::countryOfManufacturer": [
        "country of manufacturer",
        "country of origin",
        "made in",
        "country",
    ],
    "media::image1": [
        "imageurls",
        "image url",
        "image urls",
        "image links",
        "images",
        "primary image",
    ],
    "media::image2": ["image 2", "secondary image", "image url 2"],
    "media::image3": ["image 3", "image url 3"],
    "media::image4": ["image 4", "image url 4"],
    "media::image5": ["image 5", "image url 5"],
    "attr::shape": ["shape", "rug shape"],
    "attr::material": ["material", "rug material", "fiber"],
    "attr::technique": ["technique", "rug technique", "weave", "construction"],
    "attr::color": ["color", "colour", "rug color", "primary color"],
    "attr::careInstructions": ["care", "care instructions", "care guide", "cleaning"],
    "propSixtyFive::warningRequired": ["prop65", "prop 65", "proposition 65"],
}


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------


def _get_logger() -> logging.Logger:
    """Return a dedicated logger that writes to ``mapper.log``."""

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)

    return logger


def _normalize_header(name: str) -> str:
    """Normalize a column header by stripping punctuation and case."""

    return re.sub(r"[^a-z0-9]", "", name.lower())


def _clean_value(value: object) -> str:
    """Convert raw cell values into clean text."""

    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, str):
        return value.strip()
    try:
        return str(value).strip()
    except Exception:  # pragma: no cover - defensive
        return ""


def _split_semicolon(value: str, limit: Optional[int] = None) -> List[str]:
    """Split a semicolon-separated string into a list of trimmed items."""

    if not value:
        return []
    pieces = [segment.strip() for segment in value.split(";")]
    cleaned = [segment for segment in pieces if segment]
    if limit is None:
        return cleaned
    return cleaned[:limit]


def _load_mapping_yaml(path: Optional[Path]) -> Dict[str, List[str]]:
    """Load previously persisted mappings from ``path``."""

    if path is None or not path.exists():
        return {}

    try:
        with path.open("r", encoding="utf-8") as handle:
            raw = yaml.safe_load(handle) or {}
    except Exception:
        return {}

    loaded: Dict[str, List[str]] = {}
    for key, value in raw.items():
        if isinstance(value, str):
            loaded[key] = [value]
        elif isinstance(value, Iterable):
            loaded[key] = [str(item) for item in value]
    return loaded


# ---------------------------------------------------------------------------
# Core helpers
# ---------------------------------------------------------------------------


def suggest_mappings(headers: Sequence[str]) -> Dict[str, List[Tuple[float, str]]]:
    """Return candidate source columns for each Wayfair field.

    ``headers`` may contain arbitrary strings; they are normalized before
    comparison.  The returned dictionary maps the Wayfair target column to a
    list of ``(score, header_name)`` tuples sorted by descending similarity.
    """

    normalized_headers: List[Tuple[str, str]] = []
    for header in headers:
        header_text = str(header)
        if not header_text.strip():
            continue
        normalized_headers.append((header_text, _normalize_header(header_text)))

    suggestions: Dict[str, List[Tuple[float, str]]] = {}

    for target, synonyms in WAYFAIR_SYNONYMS.items():
        candidates: List[Tuple[float, str]] = []
        expanded_synonyms = list(synonyms) + [target, target.split("::")[-1]]
        normalized_synonyms = [_normalize_header(s) for s in expanded_synonyms if s]

        for original, normalized in normalized_headers:
            if not normalized:
                continue

            best_score = 0.0
            for synonym in normalized_synonyms:
                if not synonym:
                    continue
                if normalized == synonym:
                    best_score = 1.0
                    break
                score = difflib.SequenceMatcher(None, normalized, synonym).ratio()
                if score > best_score:
                    best_score = score
            if best_score > 0:
                candidates.append((best_score, original))

        if candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            suggestions[target] = candidates

    return suggestions


def _prompt_user_selection(target: str, candidates: Sequence[Tuple[float, str]]) -> Optional[str]:
    """Prompt the user to select a column for ``target`` from ``candidates``."""

    if not candidates:
        return None

    print()  # pragma: no cover - interactive helper
    print(f"Select a source column for {target} (0 = skip):")  # pragma: no cover
    limit = min(len(candidates), 10)
    for index, (score, name) in enumerate(candidates[:limit], start=1):  # pragma: no cover
        print(f"  {index}) {name} [score: {score:.2f}]")  # pragma: no cover

    while True:  # pragma: no cover - interactive loop
        choice = input("Choice: ").strip()
        if not choice or choice == "0":
            return None
        if choice.isdigit():
            index = int(choice)
            if 1 <= index <= limit:
                return candidates[index - 1][1]
        print("Invalid selection. Enter a number from the list or 0 to skip.")


def apply_mappings(df: pd.DataFrame, mapping: Mapping[str, Sequence[str]]) -> pd.DataFrame:
    """Produce a Wayfair-formatted DataFrame using ``mapping``.

    The returned DataFrame contains a column for every target defined in
    :data:`WAYFAIR_TARGETS`.  Columns without a corresponding source remain
    blank.  ``propSixtyFive::warningRequired`` defaults to ``"No"`` when no
    mapping is supplied.
    """

    output = pd.DataFrame("", index=df.index, columns=WAYFAIR_TARGETS)

    # Handle the special image column behaviour first so downstream mappings
    # can override individual columns when desired.
    image_source: Optional[str] = None
    if "media::image1" in mapping and mapping["media::image1"]:
        candidate = mapping["media::image1"][0]
        if candidate in df.columns:
            image_source = candidate
            split_values = (
                df[candidate]
                .map(_clean_value)
                .map(lambda value: _split_semicolon(value, limit=len(IMAGE_TARGETS)))
            )
            for offset, target in enumerate(IMAGE_TARGETS):
                explicit = mapping.get(target)
                if explicit and explicit[0] != candidate:
                    continue
                output[target] = split_values.map(
                    lambda items, index=offset: items[index] if index < len(items) else ""
                )

    for target, sources in mapping.items():
        if not sources:
            continue
        source = sources[0]
        if target in IMAGE_TARGETS and image_source and source == image_source:
            continue
        if source not in df.columns:
            continue
        output[target] = df[source].map(_clean_value)

    if "propSixtyFive::warningRequired" in output.columns:
        column = output["propSixtyFive::warningRequired"].map(_clean_value)
        if column.replace("", pd.NA).isna().all():
            output["propSixtyFive::warningRequired"] = "No"
        else:
            output["propSixtyFive::warningRequired"] = column

    return output


def save_mapping_yaml(mapping: Mapping[str, Sequence[str]], path: Optional[Path] = None) -> None:
    """Persist ``mapping`` to ``path`` in YAML format."""

    if path is None:
        path = DEFAULT_MAPPING_PATH

    try:
        serializable: Dict[str, List[str]] = {
            key: [value for value in values if value]
            for key, values in mapping.items()
            if values
        }
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(serializable, handle, sort_keys=True, allow_unicode=True)
    except Exception as exc:  # pragma: no cover - logging only
        logger = _get_logger()
        logger.info("Unable to save mappings to %s: %s", path, exc)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def map_wayfair_fields(
    df: pd.DataFrame,
    *,
    mapping_path: Optional[Path] = None,
    interactive: bool = True,
) -> pd.DataFrame:
    """Map ``df`` columns to Wayfair export fields.

    ``mapping_path`` controls where mapping selections are stored.  When
    ``interactive`` is ``False`` low-confidence matches are skipped instead of
    prompting the user.
    """

    logger = _get_logger()
    try:
        headers = [str(column) for column in df.columns]
        saved_mappings = _load_mapping_yaml(mapping_path or DEFAULT_MAPPING_PATH)
        suggestions = suggest_mappings(headers)

        resolved: Dict[str, List[str]] = {}
        for target in WAYFAIR_TARGETS:
            resolved[target] = []

            saved = [column for column in saved_mappings.get(target, []) if column in df.columns]
            if saved:
                resolved[target] = saved
                logger.info("Using saved mapping for %s -> %s", target, ", ".join(saved))
                continue

            candidates = suggestions.get(target, [])
            if not candidates:
                logger.info("No candidates found for %s", target)
                continue

            best_score, best_column = candidates[0]
            if best_score >= AUTO_ACCEPT_THRESHOLD:
                resolved[target] = [best_column]
                logger.info(
                    "Auto-mapped %s -> %s (score %.2f)",
                    target,
                    best_column,
                    best_score,
                )
                continue

            if interactive and best_score >= PROMPT_THRESHOLD:
                selection = _prompt_user_selection(target, candidates)
                if selection:
                    resolved[target] = [selection]
                    logger.info("User selected mapping for %s -> %s", target, selection)
                else:
                    logger.info("User skipped mapping for %s", target)
            else:
                logger.info(
                    "Skipping low-confidence mapping for %s (best candidate %s, score %.2f)",
                    target,
                    best_column,
                    best_score,
                )

        result = apply_mappings(df, resolved)
        save_mapping_yaml(resolved, mapping_path)
        return result
    except Exception as exc:  # pragma: no cover - caller should not see errors
        logger.info("Wayfair mapping failed: %s", exc)
        blank = pd.DataFrame("", index=df.index, columns=WAYFAIR_TARGETS)
        if "propSixtyFive::warningRequired" in blank.columns:
            blank["propSixtyFive::warningRequired"] = "No"
        return blank
