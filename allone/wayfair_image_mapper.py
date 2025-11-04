"""Utilities for mapping image links into Wayfair product templates."""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Hashable, Iterable, List, Mapping, Optional, Sequence, Tuple

import pandas as pd

from backend_logic import normalize_rug_number

LOGGER_NAME = "allone.image_mapper"
LOG_PATH = Path(__file__).resolve().parent / "image_mapper.log"


def _get_logger() -> logging.Logger:
    logger = logging.getLogger(LOGGER_NAME)
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(handler)
    return logger


class WayfairImageMappingError(Exception):
    """Raised when Wayfair image mapping cannot be completed."""


@dataclass
class WayfairImageMappingResult:
    """Summary information about a completed mapping run."""

    output_path: Path
    matched_count: int
    missing_count: int
    missing_rugnos: List[str]
    total_rows: int
    extra_image_rows: List[Dict[str, str]]

    @property
    def processed_count(self) -> int:
        return self.total_rows


def display_column_name(column: Hashable) -> str:
    """Return a user friendly display name for DataFrame columns."""

    if isinstance(column, tuple):
        parts = [str(part).strip() for part in column if part is not None and str(part).strip()]
        return " - ".join(parts) if parts else "(Unnamed)"
    text = str(column).strip()
    return text or "(Unnamed)"


def load_wayfair_columns(path: str) -> Tuple[str, Mapping[str, Hashable]]:
    """Load the first worksheet of the Wayfair file and return a mapping of display names."""

    logger = _get_logger()
    if not path:
        raise WayfairImageMappingError("Wayfair file path is empty.")
    try:
        workbook = pd.read_excel(path, sheet_name=None, dtype=str)
    except FileNotFoundError as exc:  # pragma: no cover - filesystem errors
        logger.error("Wayfair file not found: %s", path)
        raise WayfairImageMappingError(str(exc)) from exc
    except Exception as exc:  # pragma: no cover - pandas parsing errors
        logger.exception("Failed to read Wayfair workbook: %%s", path)
        raise WayfairImageMappingError(str(exc)) from exc

    if not workbook:
        raise WayfairImageMappingError("Wayfair workbook has no sheets.")

    sheet_name, dataframe = next(iter(workbook.items()))
    if not isinstance(dataframe, pd.DataFrame):
        dataframe = pd.DataFrame(dataframe)
    dataframe = dataframe.fillna("")

    column_map: Dict[str, Hashable] = {}
    for column in dataframe.columns:
        display_name = display_column_name(column)
        # Ensure uniqueness by appending counts if necessary
        if display_name in column_map:
            suffix = 2
            new_name = f"{display_name} ({suffix})"
            while new_name in column_map:
                suffix += 1
                new_name = f"{display_name} ({suffix})"
            display_name = new_name
        column_map[display_name] = column

    return sheet_name, column_map


def load_image_link_columns(path: str) -> Mapping[str, Hashable]:
    """Load the image link file and return column display names."""

    logger = _get_logger()
    if not path:
        raise WayfairImageMappingError("Image link file path is empty.")

    extension = os.path.splitext(path)[1].lower()
    try:
        if extension in {".xlsx", ".xls", ".xlsm", ".xlsb"}:
            dataframe = pd.read_excel(path, dtype=str)
        else:
            dataframe = pd.read_csv(path, dtype=str, keep_default_na=False)
    except FileNotFoundError as exc:  # pragma: no cover - filesystem errors
        logger.error("Image link file not found: %s", path)
        raise WayfairImageMappingError(str(exc)) from exc
    except Exception as exc:  # pragma: no cover - pandas parsing errors
        logger.exception("Failed to read image link file: %s", path)
        raise WayfairImageMappingError(str(exc)) from exc

    if not isinstance(dataframe, pd.DataFrame):
        dataframe = pd.DataFrame(dataframe)
    dataframe = dataframe.fillna("")

    column_map: Dict[str, Hashable] = {}
    for column in dataframe.columns:
        display_name = display_column_name(column)
        if display_name in column_map:
            suffix = 2
            new_name = f"{display_name} ({suffix})"
            while new_name in column_map:
                suffix += 1
                new_name = f"{display_name} ({suffix})"
            display_name = new_name
        column_map[display_name] = column

    return column_map


def _normalize_key(value: object) -> str:
    text = normalize_rug_number(value)
    if not text:
        return ""
    return re.sub(r"\s+", "", text).lower()


def _extract_suffix(value: str) -> Optional[int]:
    match = re.search(r"(?:-|_)(\d+)(?=\.[^.]+$|$)", value)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None


def _sort_links(links: Sequence[str]) -> List[str]:
    def sort_key(link: str) -> Tuple[int, str]:
        suffix = _extract_suffix(link)
        if suffix is not None:
            return (0, suffix, link.lower())  # type: ignore[return-value]
        return (1, link.lower())  # type: ignore[return-value]

    # type: ignore[arg-type] - heterogenous tuple for ordering
    return sorted(dict.fromkeys([link.strip() for link in links if link and link.strip()]), key=sort_key)  # type: ignore[arg-type]


def _has_value(value: object) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):  # type: ignore[arg-type]
            return False
    except Exception:  # pragma: no cover - defensive
        pass
    text = str(value).strip()
    if not text:
        return False
    lower_text = text.lower()
    return lower_text not in {"nan", "none"}


def _find_image_columns(columns: Iterable[Hashable]) -> Dict[int, Hashable]:
    desired: Dict[int, Hashable] = {}
    normalized_map: Dict[str, Hashable] = {}
    for column in columns:
        text = display_column_name(column).lower()
        normalized_map[text] = column

    for index in range(1, 6):
        candidates = [
            f"images - image {index} file name or url",
            f"image {index} file name or url",
            f"image {index}",
        ]
        for candidate in candidates:
            if candidate in normalized_map:
                desired[index] = normalized_map[candidate]
                break
    return desired


def map_wayfair_images(
    wayfair_path: str,
    data_sheet: str,
    wayfair_rugno_column: Hashable,
    image_link_path: str,
    image_rugno_column: Hashable,
    image_link_column: Hashable,
    preserve_existing: bool,
    output_path: Path,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    logger: Optional[logging.Logger] = None,
) -> WayfairImageMappingResult:
    """Map image links onto the Wayfair workbook and export the result."""

    logger = logger or _get_logger()
    logger.info(
        "Starting image mapping: wayfair=%s (sheet=%s), links=%s", wayfair_path, data_sheet, image_link_path
    )

    try:
        workbook = pd.read_excel(wayfair_path, sheet_name=None, dtype=str)
    except Exception as exc:  # pragma: no cover - defensive
        logger.exception("Failed to read Wayfair workbook during mapping: %s", wayfair_path)
        raise WayfairImageMappingError(str(exc)) from exc

    if data_sheet not in workbook:
        logger.error("Sheet '%s' not found in Wayfair workbook.", data_sheet)
        raise WayfairImageMappingError(f"Sheet '{data_sheet}' not found in Wayfair workbook.")

    data_frame = workbook[data_sheet]
    if not isinstance(data_frame, pd.DataFrame):
        data_frame = pd.DataFrame(data_frame)
    data_frame = data_frame.fillna("")

    extension = os.path.splitext(image_link_path)[1].lower()
    try:
        if extension in {".xlsx", ".xls", ".xlsm", ".xlsb"}:
            image_frame = pd.read_excel(image_link_path, dtype=str)
        else:
            image_frame = pd.read_csv(image_link_path, dtype=str, keep_default_na=False)
    except Exception as exc:  # pragma: no cover - defensive
        logger.exception("Failed to read image link file during mapping: %s", image_link_path)
        raise WayfairImageMappingError(str(exc)) from exc

    if not isinstance(image_frame, pd.DataFrame):
        image_frame = pd.DataFrame(image_frame)
    image_frame = image_frame.fillna("")

    link_map: Dict[str, List[str]] = {}

    for _, row in image_frame.iterrows():
        raw_key = row.get(image_rugno_column, "")
        normalized_key = _normalize_key(raw_key)
        if not normalized_key:
            continue
        link_value = row.get(image_link_column, "")
        if not link_value or not str(link_value).strip():
            continue
        link_map.setdefault(normalized_key, []).append(str(link_value).strip())

    for key, links in list(link_map.items()):
        if not links:
            link_map.pop(key, None)
            continue
        link_map[key] = _sort_links(links)

    image_columns = _find_image_columns(data_frame.columns)

    total_rows = len(data_frame.index)
    matched_count = 0
    missing_rugnos: List[str] = []
    extra_rows: List[Dict[str, str]] = []

    for index, row in data_frame.iterrows():
        raw_value = row.get(wayfair_rugno_column, "")
        normalized_value = _normalize_key(raw_value)
        if not normalized_value:
            continue
        links = link_map.get(normalized_value)
        if not links:
            missing_rugnos.append(str(raw_value).strip())
        else:
            matched_count += 1
            for position in range(5):
                column = image_columns.get(position + 1)
                if column is None or position >= len(links):
                    continue
                if preserve_existing and _has_value(row.get(column)):
                    continue
                data_frame.at[index, column] = links[position]
            if len(links) > 5:
                for extra_index, link in enumerate(links[5:], start=6):
                    extra_rows.append(
                        {
                            "RugNo": str(raw_value).strip(),
                            "Image Index": str(extra_index),
                            "Image Link": link,
                        }
                    )
        if progress_callback:
            try:
                progress_callback(index + 1, total_rows)
            except Exception:  # pragma: no cover - defensive
                logger.debug("Progress callback failed", exc_info=True)

    missing_count = len([value for value in missing_rugnos if value])

    workbook[data_sheet] = data_frame

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, sheet_df in workbook.items():
                if not isinstance(sheet_df, pd.DataFrame):
                    sheet_df = pd.DataFrame(sheet_df)
                sheet_df = sheet_df.fillna("")
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            if extra_rows:
                extra_df = pd.DataFrame(extra_rows)
                extra_df.to_excel(writer, sheet_name="Additional Images", index=False)
    except Exception as exc:  # pragma: no cover - filesystem errors
        logger.exception("Failed to write output workbook: %s", output_path)
        raise WayfairImageMappingError(str(exc)) from exc

    logger.info(
        "Mapping completed: matched=%s missing=%s total=%s output=%s",
        matched_count,
        missing_count,
        total_rows,
        output_path,
    )

    return WayfairImageMappingResult(
        output_path=Path(output_path),
        matched_count=matched_count,
        missing_count=missing_count,
        missing_rugnos=[value for value in missing_rugnos if value],
        total_rows=total_rows,
        extra_image_rows=extra_rows,
    )
