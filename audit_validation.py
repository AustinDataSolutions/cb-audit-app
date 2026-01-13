from __future__ import annotations

from typing import List, Tuple

from io import BytesIO

import pandas as pd
from openpyxl import load_workbook


REQUIRED_HEADERS = ("#", "Sentences", "Category")
# Headers for our output format (partial audit files)
# New format includes ID column; old format did not
OUTPUT_FORMAT_HEADERS_NEW = ("ID", "Sentence", "Topic", "Audit", "Explanation")
OUTPUT_FORMAT_HEADERS_OLD = ("Sentence", "Topic", "Audit", "Explanation")


def _normalize_header(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip().casefold()


def _find_header_row(header_rows: pd.DataFrame) -> Tuple[int, List[int], bool]:
    """
    Find the header row and return (row_idx, header_indices, is_output_format).

    For input format: indices are [#_col, sentence_col, category_col]
    For output format (new): indices are [id_col, sentence_col, topic_col]
    For output format (old): indices are [sentence_col, topic_col] (no ID column)
    """
    required = [h.casefold() for h in REQUIRED_HEADERS]
    output_required_new = [h.casefold() for h in OUTPUT_FORMAT_HEADERS_NEW[:3]]  # ID, Sentence, Topic
    output_required_old = [h.casefold() for h in OUTPUT_FORMAT_HEADERS_OLD[:2]]  # Sentence, Topic

    for row_idx in range(len(header_rows.index)):
        row = header_rows.iloc[row_idx].tolist()
        normalized = [_normalize_header(cell) for cell in row]

        # First check for input format
        if all(header in normalized for header in required):
            indices = [normalized.index(header) for header in required]
            return row_idx, indices, False

        # Then check for new output format (with ID)
        if all(header in normalized for header in output_required_new):
            indices = [normalized.index(header) for header in output_required_new]
            return row_idx, indices, True

        # Finally check for old output format (without ID)
        if all(header in normalized for header in output_required_old):
            indices = [normalized.index(header) for header in output_required_old]
            return row_idx, indices, True

    raise ValueError(
        "Sentences sheet must include headers '#', 'Sentences', and 'Category' "
        "(for input files) or 'ID', 'Sentence' and 'Topic' (for partial audit files) "
        "in the same row within the first two rows."
    )


def _cell_is_merged_non_top_left(ws, row: int, col: int) -> bool:
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return (row, col) != (merged.min_row, merged.min_col)
    return False


def validate_audit_sentences_sheet(file_bytes: bytes) -> Tuple[str, int, List[int], List[str], bool]:
    """
    Validate audit sentences sheet and detect format.

    Returns (sheet_name, header_row_idx, header_indices, warnings, is_output_format).

    For input format: header_indices are [#_col, sentence_col, category_col]
    For output format: header_indices are [sentence_col, topic_col] (no ID column)
    """
    excel_file = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = excel_file.sheet_names
    sentences_sheet = None
    for name in sheet_names:
        if name.casefold() == "sentences":
            sentences_sheet = name
            break
    if sentences_sheet is None:
        if not sheet_names:
            raise ValueError("Audit file has no worksheets.")
        sentences_sheet = sheet_names[0]

    header_rows = pd.read_excel(excel_file, sheet_name=sentences_sheet, header=None, nrows=2)
    header_row_idx, header_indices, is_output_format = _find_header_row(header_rows)

    full_sheet = pd.read_excel(excel_file, sheet_name=sentences_sheet, header=None)
    header_row_offset = header_row_idx + 1

    if is_output_format:
        col_sentence = header_indices[0]
        col_category = header_indices[1]
    else:
        col_sentence = header_indices[1]
        col_category = header_indices[2]

    data_row_found = False

    for row_idx in range(header_row_offset, len(full_sheet.index)):
        row = full_sheet.iloc[row_idx]
        sentence = row[col_sentence] if col_sentence < len(row) else None
        category = row[col_category] if col_category < len(row) else None

        sentence_text = "" if pd.isna(sentence) else str(sentence).strip()
        category_text = "" if pd.isna(category) else str(category).strip()

        if sentence_text or category_text:
            data_row_found = True

    if not data_row_found:
        raise ValueError("Sentences sheet must include at least one data row.")

    return sentences_sheet, header_row_idx, header_indices, [], is_output_format
