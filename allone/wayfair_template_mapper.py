"""Automated mapper for Wayfair's "Product Addition Template" workbooks.

This module reads a local master data file together with the official Wayfair
template, discovers the template columns by prefix, applies fuzzy matching to
map fields, normalizes values against the "Valid Values" sheet, and writes the
result into a copy of the template while keeping the original formatting.
"""

from __future__ import annotations

import argparse
import difflib
import logging
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ---------------------------------------------------------------------------
# Logging helpers
# ---------------------------------------------------------------------------


def setup_logger(log_path: Path) -> logging.Logger:
    """Create a logger that writes verbose information to ``log_path``."""

    logger = logging.getLogger("wayfair_mapper")
    logger.setLevel(logging.INFO)

    # Avoid duplicate handlers when the module is imported repeatedly.
    if not logger.handlers:
        handler = logging.FileHandler(log_path, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------


def _clean_value(value: object) -> str:
    """Return a trimmed string representation of ``value`` suitable for Excel."""

    if value is None:
        return ""

    if isinstance(value, float) and pd.isna(value):
        return ""

    if isinstance(value, str):
        return value.strip()

    return str(value).strip()


def _split_semicolon(value: str, limit: int) -> List[str]:
    """Split ``value`` by semicolon into at most ``limit`` trimmed items."""

    if not value:
        return []

    parts = [segment.strip() for segment in value.split(";")]
    return [segment for segment in parts if segment][:limit]


def _find_sheet(workbook, target_name: str) -> Optional[Worksheet]:
    """Return the worksheet that best matches ``target_name`` (case insensitive)."""

    lower_target = target_name.lower()
    for sheet_name in workbook.sheetnames:
        if sheet_name.lower() == lower_target:
            return workbook[sheet_name]

    # Fallback to fuzzy matching with difflib for renamed sheets.
    closest = difflib.get_close_matches(lower_target, [s.lower() for s in workbook.sheetnames], n=1, cutoff=0.6)
    if closest:
        index = [s.lower() for s in workbook.sheetnames].index(closest[0])
        return workbook[workbook.sheetnames[index]]

    return None


# ---------------------------------------------------------------------------
# Template column discovery
# ---------------------------------------------------------------------------


def detect_template_columns(sheet: Worksheet) -> Dict[str, int]:
    """Discover template columns by reading the header row of ``sheet``.

    Returns a dictionary that maps the *lowercased* column header to the column
    index (1-based).
    """

    headers: Dict[str, int] = {}

    try:
        first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    except StopIteration:
        return headers

    for idx, raw_header in enumerate(first_row, start=1):
        if raw_header is None:
            continue
        header = str(raw_header).strip()
        if not header:
            continue
        headers[header.lower()] = idx

    return headers


def _match_column(headers: Mapping[str, int], target: str) -> Optional[int]:
    """Return the column index for ``target`` using case-insensitive fuzzy matching."""

    normalized_target = target.lower()

    if normalized_target in headers:
        return headers[normalized_target]

    if "::" in normalized_target:
        prefix, suffix = normalized_target.split("::", 1)
        # Try to find exact suffix match regardless of prefix differences.
        suffix_matches = [name for name in headers if name.endswith(suffix)]
        if suffix_matches:
            best_suffix = difflib.get_close_matches(normalized_target, suffix_matches, n=1, cutoff=0.6)
            if best_suffix:
                return headers[best_suffix[0]]

    # Fallback to general fuzzy matching among all headers.
    best = difflib.get_close_matches(normalized_target, list(headers.keys()), n=1, cutoff=0.6)
    if best:
        return headers[best[0]]

    return None


# ---------------------------------------------------------------------------
# Valid value normalization
# ---------------------------------------------------------------------------


def normalize_with_valid_values(
    field_name: str,
    value: str,
    valid_map: Mapping[str, Iterable[str]],
    logger: logging.Logger,
) -> Tuple[str, bool, bool, bool]:
    """Normalize ``value`` according to the ``Valid Values`` sheet.

    Returns a tuple ``(normalized_value, was_normalized, is_valid, had_input)``.
    When ``is_valid`` is ``False`` the returned value will always be an empty
    string and the condition is logged.
    """

    cleaned_value = _clean_value(value)
    had_input = bool(cleaned_value)
    if not cleaned_value:
        return "", False, True, False

    field_key = field_name.lower()
    valid_values = valid_map.get(field_key)
    if not valid_values:
        return cleaned_value, False, True, True

    normalized_map = {str(item).strip().lower(): str(item).strip() for item in valid_values if str(item).strip()}

    if cleaned_value.lower() in normalized_map:
        normalized = normalized_map[cleaned_value.lower()]
        return normalized, normalized != cleaned_value, True, True

    # Try fuzzy matching on valid values.
    candidates = difflib.get_close_matches(cleaned_value.lower(), list(normalized_map.keys()), n=1, cutoff=0.75)
    if candidates:
        return normalized_map[candidates[0]], True, True, True

    logger.warning("Value '%s' for field '%s' does not match valid options", cleaned_value, field_name)
    return "", False, False, True


def _load_valid_values(workbook) -> Dict[str, List[str]]:
    """Read the ``Valid Values`` sheet and build a dictionary for normalization."""

    sheet = _find_sheet(workbook, "Valid Values")
    if sheet is None:
        return {}

    data = list(sheet.iter_rows(values_only=True))
    if not data:
        return {}

    valid_map: Dict[str, List[str]] = {}

    for row in data:
        if not row:
            continue
        field_name = _clean_value(row[0])
        if not field_name:
            continue
        values = [_clean_value(item) for item in row[1:] if _clean_value(item)]
        if values:
            valid_map[field_name.lower()] = values

    return valid_map


# ---------------------------------------------------------------------------
# Mapping logic
# ---------------------------------------------------------------------------


def apply_mapping(
    master_df: pd.DataFrame,
    sheet: Worksheet,
    headers: Mapping[str, int],
    valid_map: Mapping[str, Iterable[str]],
    logger: logging.Logger,
) -> Dict[str, int]:
    """Populate ``sheet`` with data from ``master_df`` using ``headers`` mapping."""

    stats: Dict[str, int] = {"products": 0, "filled": 0, "skipped": 0, "normalized": 0}

    mapping_rules: List[Tuple[str, str]] = [
        ("RugNo", "core::supplierPartNumber"),
        ("Title", "core::productName"),
        ("Collection", "core::collectionName"),
        ("UPC", "core::universalProductCode"),
        ("Wholesale", "price::wholesalePrice"),
        ("MAP", "price::minimumAdvertizedPrice"),
        ("MSRP", "price::manufacturerSuggestedRetailPrice"),
        ("LongDesc", "featureDescription::romanceCopy"),
        ("PackageWeight", "shippingAndFulfillment::weight"),
        ("PackageHeight", "shippingAndFulfillment::height"),
        ("PackageWidth", "shippingAndFulfillment::width"),
        ("PackageDepth", "shippingAndFulfillment::depth"),
    ]

    feature_targets = [
        "featureDescription::genericFeatures",
        "featureDescription::genericFeatures.1",
        "featureDescription::genericFeatures.2",
        "featureDescription::genericFeatures.3",
        "featureDescription::genericFeatures.4",
        "featureDescription::genericFeatures.5",
    ]

    prop65_target = "propSixtyFive::warningRequired"

    # Determine the column indices for each target field ahead of time.
    column_indices: Dict[str, int] = {}
    for _source, target in mapping_rules + [("Features", ft) for ft in feature_targets] + [("", prop65_target)]:
        col_idx = _match_column(headers, target)
        if col_idx is not None:
            column_indices[target] = col_idx

    features_column_indices = [column_indices[ft] for ft in feature_targets if ft in column_indices]
    prop65_column = column_indices.get(prop65_target)

    # Determine the starting row (assume row 1 is header).
    row_index = 2

    # Ensure master dataframe uses trimmed string columns to avoid pandas warnings.
    cleaned_df = master_df.where(pd.notna(master_df), None)

    for _, row in cleaned_df.iterrows():
        sku = _clean_value(row.get("RugNo"))
        if not sku:
            logger.warning("Skipping row without RugNo identifier")
            stats["skipped"] += 1
            continue

        stats["products"] += 1

        for source_column, target_header in mapping_rules:
            target_column = column_indices.get(target_header)
            if target_column is None:
                logger.warning("Template column '%s' missing; skipping", target_header)
                stats["skipped"] += 1
                continue

            raw_value = row.get(source_column)
            normalized_value, was_normalized, _is_valid, had_input = normalize_with_valid_values(
                target_header, raw_value, valid_map, logger
            )
            sheet.cell(row=row_index, column=target_column, value=normalized_value or None)
            if normalized_value:
                stats["filled"] += 1
                if was_normalized:
                    stats["normalized"] += 1
            elif had_input:
                stats["skipped"] += 1

        # Handle feature descriptions.
        features_raw = _clean_value(row.get("Features"))
        features = _split_semicolon(features_raw, limit=5)
        for idx, target_column in enumerate(features_column_indices):
            value = features[idx] if idx < len(features) else ""
            sheet.cell(row=row_index, column=target_column, value=value or None)
            if value:
                stats["filled"] += 1
            else:
                stats["skipped"] += 1

        # Prop 65 default handling.
        if prop65_column is not None:
            cell = sheet.cell(row=row_index, column=prop65_column)
            existing_value = _clean_value(cell.value)
            if not existing_value:
                sheet.cell(row=row_index, column=prop65_column, value="No")
            else:
                sheet.cell(row=row_index, column=prop65_column, value=existing_value)

        row_index += 1

    return stats


# ---------------------------------------------------------------------------
# Additional images sheet handling
# ---------------------------------------------------------------------------


def write_additional_images(
    workbook,
    master_df: pd.DataFrame,
    logger: logging.Logger,
) -> Tuple[int, int]:
    """Populate the "Additional Images" worksheet.

    Returns a tuple ``(written_rows, skipped_urls)``.
    """

    sheet = _find_sheet(workbook, "Additional Images")
    if sheet is None:
        logger.warning("'Additional Images' sheet not found in template")
        return 0, 0

    headers = detect_template_columns(sheet)

    sku_column = _match_column(headers, "core::supplierPartNumber")
    if sku_column is None:
        sku_column = _match_column(headers, "Supplier Part Number")
    image_url_column = _match_column(headers, "Image URL") or _match_column(headers, "Image Url")
    is_primary_column = _match_column(headers, "Is Primary")
    sequence_column = _match_column(headers, "Sequence")

    if sku_column is None or image_url_column is None:
        logger.warning("Additional Images sheet is missing required columns; skipping population")
        return 0, 0

    next_row = sheet.max_row + 1 if sheet.max_row else 2
    written = 0
    skipped = 0

    cleaned_df = master_df.where(pd.notna(master_df), None)

    for _, row in cleaned_df.iterrows():
        sku = _clean_value(row.get("RugNo"))
        image_urls_raw = _clean_value(row.get("ImageURLs"))
        if not sku or not image_urls_raw:
            continue

        urls = [url for url in _split_semicolon(image_urls_raw, limit=50) if url]
        if not urls:
            skipped += 1
            continue

        for order, url in enumerate(urls, start=1):
            if sku_column:
                sheet.cell(row=next_row, column=sku_column, value=sku)
            if image_url_column:
                sheet.cell(row=next_row, column=image_url_column, value=url)
            if is_primary_column:
                sheet.cell(row=next_row, column=is_primary_column, value="Yes" if order == 1 else "No")
            if sequence_column:
                sheet.cell(row=next_row, column=sequence_column, value=order)
            next_row += 1
            written += 1

    if written == 0 and skipped == 0:
        logger.info("No additional image URLs were written")

    return written, skipped


# ---------------------------------------------------------------------------
# Workbook saving
# ---------------------------------------------------------------------------


def save_filled_workbook(workbook, template_path: Path, output_path: Optional[Path] = None) -> Path:
    """Save ``workbook`` to ``output_path`` (defaults to ``*_filled.xlsx``)."""

    if output_path is None:
        output_path = template_path.with_name(f"{template_path.stem}_filled{template_path.suffix}")

    workbook.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------


def _load_master_dataframe(master_path: Path) -> pd.DataFrame:
    """Load the master dataset using pandas in a memory conscious manner."""

    suffix = master_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(master_path, dtype=str)

    if suffix in {".csv", ".txt"}:
        return pd.read_csv(master_path, dtype=str, keep_default_na=False)

    raise ValueError(f"Unsupported master file format: {master_path.suffix}")


# ---------------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------------


def run_mapper(master_path: Path, template_path: Path, output_path: Optional[Path] = None) -> None:
    """Run the full mapping pipeline."""

    log_path = Path("wayfair_mapper.log")
    logger = setup_logger(log_path)

    logger.info("Starting Wayfair template mapper")
    logger.info("Master data: %s", master_path)
    logger.info("Template workbook: %s", template_path)

    try:
        master_df = _load_master_dataframe(master_path)
    except Exception as exc:  # pragma: no cover - defensive logging
        logger.exception("Failed to read master data: %s", exc)
        print("Error: Unable to read master data. See log for details.")
        return

    try:
        workbook = load_workbook(template_path)
    except Exception as exc:  # pragma: no cover - defensive logging
        logger.exception("Failed to load template workbook: %s", exc)
        print("Error: Unable to load template workbook. See log for details.")
        return

    main_sheet = _find_sheet(workbook, "7524 - One-of-a-Kind Ru")
    if main_sheet is None:
        logger.error("Main sheet '7524 - One-of-a-Kind Ru' not found in template")
        print("Error: Main template sheet missing. See log for details.")
        return

    headers = detect_template_columns(main_sheet)
    if not headers:
        logger.error("Unable to detect column headers in main sheet")
        print("Error: Failed to read template headers. See log for details.")
        return

    valid_map = _load_valid_values(workbook)

    stats = apply_mapping(master_df, main_sheet, headers, valid_map, logger)
    additional_written, additional_skipped = write_additional_images(workbook, master_df, logger)

    output = save_filled_workbook(workbook, template_path, output_path)

    logger.info("Workbook saved to %s", output)
    logger.info("Products processed: %s", stats.get("products", 0))
    logger.info("Fields filled: %s", stats.get("filled", 0))
    logger.info("Values normalized: %s", stats.get("normalized", 0))
    logger.info("Values skipped: %s", stats.get("skipped", 0))
    logger.info("Additional images rows written: %s", additional_written)
    logger.info("Additional images rows skipped: %s", additional_skipped)

    summary = (
        f"Products processed: {stats.get('products', 0)} | "
        f"Fields filled: {stats.get('filled', 0)} | "
        f"Normalized: {stats.get('normalized', 0)} | "
        f"Skipped: {stats.get('skipped', 0)} | "
        f"Additional image rows: {additional_written} | "
        f"Additional image skips: {additional_skipped}"
    )
    print(summary)


def _parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill a Wayfair Product Addition Template with master data.")
    parser.add_argument("master", type=Path, help="Path to the master data file (Excel or CSV)")
    parser.add_argument("template", type=Path, help="Path to the Wayfair template workbook")
    parser.add_argument("--output", type=Path, default=None, help="Optional output path for the filled workbook")
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> None:
    args = _parse_args(argv)
    run_mapper(args.master, args.template, args.output)


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    main(sys.argv[1:])

