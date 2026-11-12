#!/usr/bin/env python3
"""Isolated, batch-oriented XLSX modification worker for ExcelEditPlan.

This module contains spreadsheet operations only. It is invoked once for a
whole workbook plan and never imports either product's business code.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import time
import traceback
import zipfile
from copy import copy
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree
from xml.parsers import expat
from xml.sax.saxutils import escape as xml_escape

SCHEMA_VERSION = 1
# 与 Rust 侧 openpyxl_backend.rs 的 OpenpyxlWorkerConfig::PROTOCOL_VERSION
# 必须保持一致；修改任一侧时必须同步修改另一侧。
WORKER_PROTOCOL_VERSION = 1
# These two legacy PyCode features are explicitly Openpyxl-based after the
# user confirmed that their outputs are opened in Excel/WPS for recalculation.
# Preserve their cached results in the candidate and request a full recalc on
# open, but do not block the write on the generic formula-dependency proof.
PYC_CODE_OPENPYXL_PARITY_REQUESTS = {
    "pytools-split-village-bank",
    "pytools-adjust-rural-loan",
}
SUPPORTED_OPERATIONS = {
    "set_cell_value",
    "set_formula_text",
    "delete_rows",
    "insert_rows",
    "set_comment",
    "set_column_width",
    "set_row_height",
    "rename_sheet",
    "create_sheet",
    "create_sheet_at",
    "set_freeze_panes",
    "set_auto_filter",
    "delete_sheet",
    "set_cell_fill_color",
    "copy_range",
    "import_range",
    "set_workbook_calculation_properties",
    "set_page_settings",
    "replace_header_text",
}


class WorkerError(Exception):
    def __init__(self, kind: str, message: str, *, unsupported: bool = False) -> None:
        super().__init__(message)
        self.kind = kind
        self.unsupported = unsupported


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _formula_cache_cells(path: Path) -> List[Tuple[str, str]]:
    """Return coordinates with non-empty formula caches, resolving sheet rels.

    openpyxl drops cached formula results when it saves a workbook. The guard
    therefore needs sheet names as well as cell addresses so it can distinguish
    planned destination writes from caches that must remain untouched.
    """
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    document_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with zipfile.ZipFile(str(path), "r") as package:
        workbook = ElementTree.fromstring(package.read("xl/workbook.xml"))
        relationships = ElementTree.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relationship.attrib["Id"]: relationship.attrib["Target"]
            for relationship in relationships.findall("{%s}Relationship" % package_rel_ns)
            if relationship.attrib.get("TargetMode") != "External"
        }
        cached_cells = []
        sheets = workbook.find("{%s}sheets" % main_ns)
        if sheets is None:
            raise WorkerError("invalid_workbook", "workbook sheet list is missing", unsupported=True)
        for sheet in sheets.findall("{%s}sheet" % main_ns):
            sheet_name = sheet.attrib.get("name")
            relationship_id = sheet.attrib.get("{%s}id" % document_rel_ns)
            target = targets.get(relationship_id or "")
            if not sheet_name or not target:
                raise WorkerError(
                    "unsupported_workbook_structure",
                    "cannot resolve worksheet relationship for formula-cache inspection",
                    unsupported=True,
                )
            part = posixpath.normpath(
                target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
            )
            if part.startswith("../") or not part.startswith("xl/") or part not in package.namelist():
                raise WorkerError(
                    "unsupported_workbook_structure",
                    "worksheet relationship points outside the supported XLSX package layout",
                    unsupported=True,
                )
            root = ElementTree.fromstring(package.read(part))
            for cell in root.iter("{%s}c" % main_ns):
                formula = cell.find("{%s}f" % main_ns)
                cached = cell.find("{%s}v" % main_ns)
                address = cell.attrib.get("r")
                if formula is not None and cached is not None and cached.text not in (None, ""):
                    if not address:
                        raise WorkerError(
                            "unsupported_workbook_structure",
                            "formula cell with cached value has no coordinate",
                            unsupported=True,
                        )
                    cached_cells.append((sheet_name, address.replace("$", "").upper()))
        return cached_cells


def _range_write_covers_cell(sheet: str, address: str, operation: Dict[str, Any]) -> bool:
    """Check only operations that deterministically write every cell in a range."""
    if operation.get("operation") not in ("import_range", "copy_range"):
        return False
    if operation.get("target_sheet") != sheet:
        return False
    # `_copy_cells` skips blank source cells unless materialize_blank_cells is
    # set. Formula copies also replace an existing cached result when present.
    # Style/comment-only copies do not replace formula values.
    if not operation.get("materialize_blank_cells") or not (
        operation.get("copy_values") or operation.get("copy_formulas")
    ):
        return False
    try:
        from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
        source_min_col, source_min_row, source_max_col, source_max_row = range_boundaries(
            operation["source_range"]
        )
        target_row, target_col = coordinate_to_tuple(operation["target_start"].replace("$", ""))
        row, column = coordinate_to_tuple(address)
    except (KeyError, TypeError, ValueError):
        return False
    if not all((source_min_col, source_min_row, source_max_col, source_max_row)):
        return False
    target_max_row = target_row + source_max_row - source_min_row
    target_max_col = target_col + source_max_col - source_min_col
    return target_row <= row <= target_max_row and target_col <= column <= target_max_col


def _parse_cell_address(address: str) -> Tuple[int, int]:
    normalized = address.replace("$", "").upper()
    match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", normalized)
    if not match:
        raise WorkerError(
            "unsupported_workbook_structure",
            "invalid formula-cache coordinate " + address,
            unsupported=True,
        )
    column = 0
    for character in match.group(1):
        column = column * 26 + ord(character) - ord("A") + 1
    return int(match.group(2)), column


def _format_cell_address(row: int, column: int) -> str:
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters + str(row)


def _formula_cache_destination(
    sheet: str, address: str, operations: List[Dict[str, Any]]
) -> Optional[Tuple[str, str]]:
    """Map a saved formula cache through the plan, or omit it if overwritten/deleted."""
    current_sheet = sheet
    row, column = _parse_cell_address(address)
    for operation in operations:
        kind = operation.get("operation")
        if kind == "delete_sheet" and operation.get("name") == current_sheet:
            return None
        if kind == "rename_sheet" and operation.get("old_name") == current_sheet:
            current_sheet = operation.get("new_name", current_sheet)
            continue
        if kind in ("delete_rows", "insert_rows") and operation.get("sheet") == current_sheet:
            first_row = int(operation.get("first_row", 0))
            count = int(operation.get("count", 0))
            if first_row < 1 or count < 1:
                raise WorkerError(
                    "invalid_operation", "invalid row operation while preserving formula caches"
                )
            if kind == "delete_rows":
                last_deleted = first_row + count - 1
                if first_row <= row <= last_deleted:
                    return None
                if row > last_deleted:
                    row -= count
            elif row >= first_row:
                row += count
        if kind in ("set_cell_value", "set_formula_text"):
            if operation.get("sheet") == current_sheet:
                target_row, target_column = _parse_cell_address(operation.get("cell", ""))
                if (target_row, target_column) == (row, column):
                    return None
        elif _range_write_covers_cell(
            current_sheet, _format_cell_address(row, column), operation
        ):
            return None
    return current_sheet, _format_cell_address(row, column)


def _uncovered_formula_cache_cells(path: Path, operations: List[Dict[str, Any]]) -> List[Tuple[str, str]]:
    return [
        (sheet, address)
        for sheet, address in _formula_cache_cells(path)
        if _formula_cache_destination(sheet, address, operations) is not None
    ]


def _is_add_new_sheets_only_plan(operations: List[Dict[str, Any]]) -> bool:
    """Whether the plan only creates sheets and imports values into those sheets.

    This narrow plan shape cannot invalidate cached formulas in the input
    workbook. We can therefore restore those caches after openpyxl saves.
    """
    created_sheets = set()
    saw_import = False
    for operation in operations:
        kind = operation.get("operation")
        if kind in ("create_sheet", "create_sheet_at"):
            name = operation.get("name")
            if not isinstance(name, str) or not name:
                return False
            created_sheets.add(name)
        elif kind == "import_range":
            target = operation.get("target_sheet")
            if (
                target not in created_sheets
                or not operation.get("copy_values")
                or operation.get("copy_formulas")
            ):
                return False
            saw_import = True
        else:
            return False
    return bool(created_sheets) and saw_import


def _formula_cache_values(path: Path) -> Dict[Tuple[str, str], Tuple[Optional[str], str]]:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(str(path), "r") as package:
        title_to_part = _worksheet_parts_by_title(package)
        values = {}
        for title, part in title_to_part.items():
            if part not in package.namelist():
                continue
            root = ElementTree.fromstring(package.read(part))
            for cell in root.iter("{%s}c" % main_ns):
                if cell.find("{%s}f" % main_ns) is None:
                    continue
                cached = cell.find("{%s}v" % main_ns)
                address = cell.attrib.get("r")
                if cached is None or cached.text in (None, "") or not address:
                    continue
                values[(title, address.replace("$", "").upper())] = (
                    cell.attrib.get("t"),
                    cached.text,
                )
        return values


def _formula_cache_formulas(path: Path) -> Dict[Tuple[str, str], str]:
    """Read formula text for cells whose saved result is non-empty."""
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(str(path), "r") as package:
        values = {}
        for title, part in _worksheet_parts_by_title(package).items():
            if part not in package.namelist():
                continue
            root = ElementTree.fromstring(package.read(part))
            for cell in root.iter("{%s}c" % main_ns):
                formula = cell.find("{%s}f" % main_ns)
                cached = cell.find("{%s}v" % main_ns)
                address = cell.attrib.get("r")
                if formula is None or cached is None or cached.text in (None, "") or not address:
                    continue
                values[(title, address.replace("$", "").upper())] = formula.text or ""
        return values


def _all_formula_texts(path: Path) -> Dict[Tuple[str, str], str]:
    """Read every formula cell, including cells without a saved result."""
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(str(path), "r") as package:
        values = {}
        for title, part in _worksheet_parts_by_title(package).items():
            if part not in package.namelist():
                continue
            root = ElementTree.fromstring(package.read(part))
            for cell in root.iter("{%s}c" % main_ns):
                formula = cell.find("{%s}f" % main_ns)
                address = cell.attrib.get("r")
                if formula is None or not address:
                    continue
                values[(title, address.replace("$", "").upper())] = formula.text or ""
        return values


def _parse_formula_reference(value: str, formula_sheet: str) -> Optional[Tuple[str, int, int, int, int]]:
    """Parse one local A1 reference token into (sheet, rows, columns).

    A name, external link, 3-D reference, or other non-A1 expression returns
    None so the caller can keep the cache-preservation decision conservative.
    """
    from openpyxl.utils.cell import range_boundaries

    reference = value.strip()
    sheet = formula_sheet
    if "!" in reference:
        qualifier, reference = reference.rsplit("!", 1)
        qualifier = qualifier.strip()
        if qualifier.startswith("'") and qualifier.endswith("'"):
            qualifier = qualifier[1:-1].replace("''", "'")
        elif qualifier.startswith("[") or "[" in qualifier or "]" in qualifier:
            return None
        if not qualifier or ":" in qualifier:
            return None
        sheet = qualifier
    if "[" in reference or "]" in reference:
        return None
    try:
        min_col, min_row, max_col, max_row = range_boundaries(reference)
    except (TypeError, ValueError):
        return None
    if min_col is None or max_col is None:
        return None
    # Whole-column/whole-row references are valid A1 references. Expand only
    # the missing axis to Excel's documented sheet limits.
    if min_row is None:
        min_row, max_row = 1, 1_048_576
    if max_row is None:
        max_row = 1_048_576
    if min_col is None:
        min_col, max_col = 1, 16_384
    if max_col is None:
        max_col = 16_384
    return sheet, min_row, max_row, min_col, max_col


def _formula_dependencies(
    formula_sheet: str,
    formula_address: str,
    formula: str,
    cached_workbook: Any,
) -> Optional[List[Tuple[str, int, int, int, int]]]:
    """Resolve supported formula references; None means dependencies unknown.

    Direct A1 references are parsed with openpyxl's formula tokenizer. The
    two INDIRECT patterns used by existing Excel workbooks are resolved from
    the saved selector cell and the formula cell's own address. Other dynamic
    references remain unsupported for cache-preserving value edits.
    """
    from openpyxl.formula import Tokenizer
    from openpyxl.utils.cell import coordinate_to_tuple

    compact = re.sub(r"\s+", "", formula or "")
    upper = compact.upper()
    indirect_references = None
    selector = None
    selector_reference = None
    direct = re.fullmatch(
        r'INDIRECT\((\$?[A-Z]{1,3}\$?[1-9][0-9]*)&"(![A-Z]{1,3}[1-9][0-9]*)"\)',
        compact,
        flags=re.IGNORECASE,
    )
    address_fixed = re.fullmatch(
        r'INDIRECT\((\$?[A-Z]{1,3}\$?[1-9][0-9]*)&"!"&ADDRESS\(([1-9][0-9]*),([1-9][0-9]*)\)\)',
        compact,
        flags=re.IGNORECASE,
    )
    address_self = re.fullmatch(
        r'INDIRECT\((\$?[A-Z]{1,3}\$?[1-9][0-9]*)&"!"&ADDRESS\(ROW\(\),COLUMN\(\)\)\)',
        compact,
        flags=re.IGNORECASE,
    )
    if direct or address_fixed or address_self:
        match = direct or address_fixed or address_self
        selector_reference = match.group(1).replace("$", "").upper()
        if direct:
            target_address = direct.group(2)[1:]
        elif address_fixed:
            target_row, target_column = int(address_fixed.group(2)), int(address_fixed.group(3))
            from openpyxl.utils.cell import get_column_letter
            target_address = get_column_letter(target_column) + str(target_row)
        else:
            target_address = formula_address.replace("$", "").upper()
        if formula_sheet not in cached_workbook.sheetnames:
            return None
        selector = cached_workbook[formula_sheet][selector_reference].value
        if not isinstance(selector, str) or selector not in cached_workbook.sheetnames:
            return None
        formula_row, formula_column = coordinate_to_tuple(formula_address.replace("$", "").upper())
        target = _parse_formula_reference(
            "{}!{}".format(selector, target_address), formula_sheet
        )
        selector_dep = _parse_formula_reference(selector_reference, formula_sheet)
        if target is None or selector_dep is None:
            return None
        indirect_references = [selector_dep, target]
    elif "INDIRECT(" in upper or "OFFSET(" in upper:
        return None

    try:
        tokens = Tokenizer("=" + (formula or "")).items
    except Exception:
        return None
    references: List[Tuple[str, int, int, int, int]] = []
    for token in tokens:
        if token.type == "FUNC" and token.subtype == "OPEN":
            name = token.value[:-1].upper()
            if name in ("INDIRECT", "OFFSET"):
                if indirect_references is None:
                    return None
        if token.type == "OPERAND" and token.subtype == "RANGE":
            # The selector in a supported INDIRECT expression was already
            # handled above; the tokenizer also exposes it as a direct range.
            parsed = _parse_formula_reference(token.value, formula_sheet)
            if parsed is None:
                return None
            if indirect_references is None or token.value.replace("$", "").upper() != selector_reference:
                references.append(parsed)
    if indirect_references is not None:
        references.extend(indirect_references)
    return references


def _formula_cache_write_conflict(
    path: Path, operations: List[Dict[str, Any]], request_id: Optional[str] = None
) -> Optional[str]:
    """Return why cached results cannot be proven safe for this write plan."""
    if request_id in PYC_CODE_OPENPYXL_PARITY_REQUESTS:
        # Match the selected PyCode Openpyxl route. The worker still restores
        # saved caches and marks the workbook for a full calculation on open;
        # users may also invoke 4-4 when they need an immediate recalculation.
        return None
    from openpyxl.utils.cell import coordinate_to_tuple

    formulas = _formula_cache_formulas(path)
    formula_sensitive = any(
        operation.get("operation")
        in ("set_cell_value", "set_formula_text", "delete_rows", "insert_rows", "copy_range", "import_range")
        for operation in operations
    )
    if not formula_sensitive or _is_add_new_sheets_only_plan(operations):
        return None

    row_delete_only = bool(operations) and all(
        operation.get("operation") == "delete_rows" for operation in operations
    )
    if row_delete_only:
        conflict = _row_delete_formula_dependency_conflict(path, operations)
        if conflict:
            return conflict
        # The dependency check proves the formulas do not refer to a modified
        # sheet. Formula caches on every other sheet retain their coordinates;
        # the normal cache restoration path can therefore preserve them.
        return None

    if not formulas:
        return None

    if any(
        operation.get("operation") not in ("set_cell_value", "set_formula_text")
        for operation in operations
        if operation.get("operation")
        in ("set_cell_value", "set_formula_text", "delete_rows", "insert_rows", "copy_range", "import_range")
    ):
        remaining = _uncovered_formula_cache_cells(path, operations)
        if remaining:
            return "formula-sensitive row/range operations cannot be proven independent of {} saved formula cache(s)".format(
                len(remaining)
            )
        return None

    writes = []
    for operation in operations:
        if operation.get("operation") not in ("set_cell_value", "set_formula_text"):
            continue
        sheet = operation.get("sheet")
        address = operation.get("cell")
        if not isinstance(sheet, str) or not isinstance(address, str):
            return "a formula-sensitive cell write has no valid sheet/cell location"
        try:
            row, column = coordinate_to_tuple(address.replace("$", "").upper())
        except ValueError:
            return "a formula-sensitive cell write has an invalid address: " + address
        writes.append((sheet, row, column))

    # 2026-11-10 用户口径：单元格写入与公式依赖相交时不再拒绝整个计划。
    # 依赖这些写入的公式缓存改由 _preserve_formula_caches 跳过恢复（输出里
    # 不带旧缓存值），执行路径已设置 fullCalcOnLoad/forceFullCalc，Excel/WPS
    # 打开时统一重算；其他读取方看到的是「无缓存」而不是「过期数值」。
    # 行/区间操作（delete_rows/insert_rows/copy_range/import_range）仍然
    # 保持严格拒绝：坐标位移下「跳过恢复」不足以证明缓存安全。
    return None


def _row_delete_formula_dependency_conflict(
    path: Path, operations: List[Dict[str, Any]]
) -> Optional[str]:
    """Allow row deletion only when all formula/name dependencies are provably unrelated.

    Openpyxl does not rewrite arbitrary formulas and defined names when rows are
    deleted. This narrow path therefore rejects formulas on edited sheets,
    cross-sheet references into edited sheets, dynamic/unknown references, and
    names whose dependencies cannot be resolved. Unrelated formula caches can
    then be restored at their original coordinates after the save.
    """
    modified_sheets = set()
    for operation in operations:
        sheet = operation.get("sheet")
        if not isinstance(sheet, str) or not sheet:
            return "a row deletion has no valid worksheet name"
        modified_sheets.add(sheet.casefold())

    formula_texts = _all_formula_texts(path)
    if any(sheet.casefold() in modified_sheets for sheet, _ in formula_texts):
        return "row deletion is not allowed on a worksheet that contains formulas"

    class EmptyWorkbook:
        sheetnames: List[str] = []

    for (formula_sheet, address), formula in formula_texts.items():
        upper = formula.upper()
        if "INDIRECT(" in upper or "OFFSET(" in upper:
            return "cannot prove row deletion is independent of a dynamic formula reference"
        if not formula.strip():
            return "cannot resolve a shared or empty formula while checking row-deletion dependencies"
        dependencies = _formula_dependencies(
            formula_sheet, address, formula, EmptyWorkbook()
        )
        if dependencies is None:
            return "cannot prove whether formula {}!{} depends on a worksheet being edited".format(
                formula_sheet, address
            )
        for dependency in dependencies:
            if dependency[0].casefold() in modified_sheets:
                return "formula {}!{} depends on a worksheet being edited".format(
                    formula_sheet, address
                )

    try:
        import openpyxl
        workbook = openpyxl.load_workbook(
            str(path), read_only=True, data_only=True, keep_links=True
        )
    except Exception as error:
        return "could not inspect defined names before row deletion: " + str(error)
    try:
        for defined_name in workbook.defined_names.values():
            reference = getattr(defined_name, "attr_text", None)
            if not isinstance(reference, str) or not reference.strip():
                continue
            scope_index = getattr(defined_name, "localSheetId", None)
            if isinstance(scope_index, int) and 0 <= scope_index < len(workbook.sheetnames):
                formula_sheet = workbook.sheetnames[scope_index]
            elif workbook.sheetnames:
                formula_sheet = workbook.sheetnames[0]
            else:
                return "cannot resolve the scope of a defined name before row deletion"
            if "INDIRECT(" in reference.upper() or "OFFSET(" in reference.upper():
                return "cannot prove row deletion is independent of a dynamic defined name"
            dependencies = _formula_dependencies(
                formula_sheet, "A1", reference.lstrip("="), EmptyWorkbook()
            )
            if dependencies is None:
                return "cannot prove whether defined name {!r} depends on a worksheet being edited".format(
                    getattr(defined_name, "name", "<unnamed>")
                )
            for dependency in dependencies:
                if dependency[0].casefold() in modified_sheets:
                    return "defined name {!r} depends on a worksheet being edited".format(
                        getattr(defined_name, "name", "<unnamed>")
                    )
    finally:
        workbook.close()
    return None


def _restore_formula_cache_cell(
    worksheet_xml: bytes, address: str, cell_type: Optional[str], cached_value: str
) -> bytes:
    escaped_address = re.escape(address.encode("ascii"))
    cell_pattern = re.compile(
        rb"<c\b(?=[^>]*\br=[\"']" + escaped_address + rb"[\"'])[^>]*>.*?</c>",
        re.DOTALL,
    )
    matches = list(cell_pattern.finditer(worksheet_xml))
    if len(matches) != 1:
        raise WorkerError(
            "formula_cache_restore_failed",
            "could not uniquely locate formula cache cell " + address,
        )
    match = matches[0]
    cell_xml = match.group(0)
    if re.search(rb"<f(?:\s|>)", cell_xml) is None:
        raise WorkerError(
            "formula_cache_restore_failed",
            "formula cell is missing after openpyxl save: " + address,
        )
    opening = re.match(rb"<c\b[^>]*>", cell_xml)
    if opening is None:
        raise WorkerError("formula_cache_restore_failed", "invalid formula cell " + address)
    start_tag = opening.group(0)
    type_pattern = re.compile(rb'\s+t="[^"]*"')
    if cell_type is None:
        start_tag = type_pattern.sub(b"", start_tag)
    elif type_pattern.search(start_tag):
        start_tag = type_pattern.sub(b' t="' + cell_type.encode("ascii") + b'"', start_tag, count=1)
    else:
        start_tag = start_tag[:-1] + b' t="' + cell_type.encode("ascii") + b'">'
    cell_xml = start_tag + cell_xml[opening.end():]

    from xml.sax.saxutils import escape

    escaped_value = escape(cached_value).encode("utf-8")
    value_pattern = re.compile(rb"<v(?:\s[^>]*)?>(.*?)</v>", re.DOTALL)
    if value_pattern.search(cell_xml):
        cell_xml = value_pattern.sub(
            lambda value_match: value_match.group(0)[: value_match.start(1) - value_match.start(0)]
            + escaped_value
            + value_match.group(0)[value_match.end(1) - value_match.start(0) :],
            cell_xml,
            count=1,
        )
    else:
        empty_value_pattern = re.compile(rb"<v(?:\s[^>]*)?\s*/>")
        if empty_value_pattern.search(cell_xml):
            cell_xml = empty_value_pattern.sub(b"<v>" + escaped_value + b"</v>", cell_xml, count=1)
        else:
            closing = cell_xml.rfind(b"</c>")
            if closing < 0:
                raise WorkerError(
                    "formula_cache_restore_failed",
                    "formula cell has no closing tag: " + address,
                )
            cell_xml = cell_xml[:closing] + b"<v>" + escaped_value + b"</v>" + cell_xml[closing:]
    return worksheet_xml[: match.start()] + cell_xml + worksheet_xml[match.end() :]


def _planned_write_cells(operations: List[Dict[str, Any]]) -> List[Tuple[str, int, int]]:
    """Collect set_cell_value / set_formula_text targets as (sheet, row, column)."""
    from openpyxl.utils.cell import coordinate_to_tuple

    writes: List[Tuple[str, int, int]] = []
    for operation in operations:
        if operation.get("operation") not in ("set_cell_value", "set_formula_text"):
            continue
        sheet = operation.get("sheet")
        address = operation.get("cell")
        if not isinstance(sheet, str) or not isinstance(address, str):
            continue
        try:
            row, column = coordinate_to_tuple(address.replace("$", "").upper())
        except ValueError:
            continue
        writes.append((sheet, row, column))
    return writes


def _cache_depends_on_planned_writes(
    formula_sheet: str,
    formula_address: str,
    formula: str,
    writes: List[Tuple[str, int, int]],
    dependency_context: Any,
) -> bool:
    """Whether restoring this cache would leave a value invalidated by the plan.

    Returns True (skip the restore) when dependencies cannot be proven, so the
    output never carries a stale cached value; Excel/WPS recalculates it on
    open because the execute path sets fullCalcOnLoad/forceFullCalc.
    """
    if "INDIRECT(" in formula.upper() or "OFFSET(" in formula.upper():
        # Dynamic references are only provably safe when the tokenizer resolves
        # them against the saved workbook; anything else stays conservative.
        dependencies = _formula_dependencies(
            formula_sheet, formula_address, formula, dependency_context
        )
        if dependencies is None:
            return True
    else:
        class EmptyWorkbook:
            sheetnames: List[str] = []

        dependencies = _formula_dependencies(
            formula_sheet, formula_address, formula, EmptyWorkbook()
        )
        if dependencies is None:
            return True
    for write_sheet, write_row, write_column in writes:
        for dependency_sheet, min_row, max_row, min_column, max_column in dependencies:
            if (
                write_sheet == dependency_sheet
                and min_row <= write_row <= max_row
                and min_column <= write_column <= max_column
            ):
                return True
    return False


def _preserve_formula_caches(input_path: Path, output_path: Path, operations: List[Dict[str, Any]]) -> None:
    cached_values = {}
    writes = _planned_write_cells(operations)
    formulas = _formula_cache_formulas(input_path)
    cached_workbook = None
    if any("INDIRECT(" in formula.upper() or "OFFSET(" in formula.upper() for formula in formulas.values()):
        try:
            import openpyxl
            cached_workbook = openpyxl.load_workbook(
                str(input_path), read_only=True, data_only=True, keep_links=True
            )
        except Exception:
            cached_workbook = None
    skipped_stale = 0
    try:
        for key, value in _formula_cache_values(input_path).items():
            destination = _formula_cache_destination(key[0], key[1], operations)
            if destination is None:
                continue
            formula = formulas.get(key)
            if writes and formula and _cache_depends_on_planned_writes(
                key[0], key[1], formula, writes, cached_workbook
            ):
                # The plan writes into this formula's dependency range: the old
                # cached value would be stale. Leave the cell uncached instead;
                # Excel/WPS recalculates it on open (fullCalcOnLoad is set).
                skipped_stale += 1
                continue
            if destination in cached_values:
                raise WorkerError(
                    "formula_cache_restore_failed",
                    "multiple input formula caches map to "
                    + destination[0]
                    + "!"
                    + destination[1],
                )
            cached_values[destination] = value
    finally:
        if cached_workbook is not None:
            cached_workbook.close()
    if not cached_values:
        return
    replacement = output_path.with_name(output_path.name + ".formula-cache.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "formula-cache temporary output already exists")
    try:
        with zipfile.ZipFile(str(output_path), "r") as package:
            title_to_part = _worksheet_parts_by_title(package)
            updated_parts: Dict[str, bytes] = {}
            for (title, address), (cell_type, cached_value) in cached_values.items():
                part = title_to_part.get(title)
                if not part or part not in package.namelist():
                    raise WorkerError(
                        "formula_cache_restore_failed",
                        "worksheet is missing after openpyxl save: " + title,
                    )
                worksheet_xml = updated_parts.get(part, package.read(part))
                updated_parts[part] = _restore_formula_cache_cell(
                    worksheet_xml, address, cell_type, cached_value
                )
        with zipfile.ZipFile(str(output_path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                target.writestr(item, updated_parts.get(item.filename, source.read(item.filename)))
        os.replace(str(replacement), str(output_path))
    except WorkerError:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise WorkerError("save_failed", "could not preserve formula caches: " + str(error))


def _column_dimensions_modified(
    source_sheet: str, operations: List[Dict[str, Any]]
) -> Tuple[Optional[str], bool]:
    """Map a source sheet through the edit plan and report planned width edits."""
    current_sheet = source_sheet
    for operation in operations:
        kind = operation.get("operation")
        if kind == "delete_sheet" and operation.get("name") == current_sheet:
            return None, False
        if kind == "rename_sheet" and operation.get("old_name") == current_sheet:
            current_sheet = operation.get("new_name", current_sheet)
            continue
        if kind == "set_column_width" and operation.get("sheet") == current_sheet:
            return current_sheet, True
        if (
            kind == "import_range"
            and operation.get("target_sheet") == current_sheet
            and (operation.get("copy_column_widths") or operation.get("copy_all_column_dimensions"))
        ):
            return current_sheet, True
    return current_sheet, False


def _column_dimension_fragment(data: bytes) -> Optional[bytes]:
    prefix = rb"(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)"
    paired = re.search(
        rb"<" + prefix + rb"cols\b[^>]*>.*?</(?P=prefix)cols\s*>",
        data,
        re.DOTALL,
    )
    if paired:
        return paired.group(0)
    empty = re.search(rb"<" + prefix + rb"cols\b[^>]*/\s*>", data)
    return empty.group(0) if empty else None


def _restore_unmodified_column_dimensions(
    input_path: Path, output_path: Path, operations: List[Dict[str, Any]]
) -> None:
    """Restore exact OOXML column records where the edit plan does not change widths.

    Openpyxl can serialize a style-less ``<col>`` as an explicit default style
    index. That changes the source column semantics and can also introduce a
    style reference unrelated to the requested cell edit. Keep the original
    ``<cols>`` fragment byte-for-byte on untouched sheets; planned width edits
    remain under Openpyxl's normal serializer and Rust readback validation.
    """
    with zipfile.ZipFile(str(input_path), "r") as original, zipfile.ZipFile(
        str(output_path), "r"
    ) as candidate:
        original_parts = _worksheet_parts_by_title(original)
        candidate_parts = _worksheet_parts_by_title(candidate)
        updates: Dict[str, bytes] = {}
        expectations: Dict[str, Optional[bytes]] = {}
        for source_sheet, source_part in original_parts.items():
            output_sheet, modified = _column_dimensions_modified(source_sheet, operations)
            if output_sheet is None or modified:
                continue
            output_part = candidate_parts.get(output_sheet)
            if not output_part or output_part not in candidate.namelist():
                raise WorkerError(
                    "column_dimension_restore_failed",
                    "worksheet is missing after openpyxl save: " + output_sheet,
                )
            original_xml = original.read(source_part)
            output_xml = candidate.read(output_part)
            expected = _column_dimension_fragment(original_xml)
            actual = _column_dimension_fragment(output_xml)
            if expected == actual:
                continue
            if actual is not None:
                match = re.search(
                    rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)cols\b[^>]*>.*?</(?P=prefix)cols\s*>",
                    output_xml,
                    re.DOTALL,
                )
                if match is None:
                    match = re.search(
                        rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)cols\b[^>]*/\s*>",
                        output_xml,
                    )
                if match is None:
                    raise WorkerError(
                        "column_dimension_restore_failed",
                        "could not locate serialized column dimensions on worksheet " + output_sheet,
                    )
                if expected is None:
                    restored_xml = output_xml[: match.start()] + output_xml[match.end() :]
                else:
                    restored_xml = output_xml[: match.start()] + expected + output_xml[match.end() :]
            elif expected is not None:
                sheet_data = re.search(
                    rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)sheetData\b",
                    output_xml,
                )
                if sheet_data is None:
                    raise WorkerError(
                        "column_dimension_restore_failed",
                        "could not locate sheetData on worksheet " + output_sheet,
                    )
                restored_xml = output_xml[: sheet_data.start()] + expected + output_xml[sheet_data.start() :]
            else:
                continue
            updates[output_part] = restored_xml
            expectations[output_part] = expected

    if not updates:
        return
    replacement = output_path.with_name(output_path.name + ".column-dimensions.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "column-dimension temporary output already exists")
    try:
        with zipfile.ZipFile(str(output_path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                target.writestr(item, updates.get(item.filename, source.read(item.filename)))
        os.replace(str(replacement), str(output_path))
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise WorkerError("save_failed", "could not restore column dimensions: " + str(error))

    with zipfile.ZipFile(str(output_path), "r") as result:
        for part, expected in expectations.items():
            if _column_dimension_fragment(result.read(part)) != expected:
                raise WorkerError(
                    "column_dimension_restore_failed",
                    "column dimensions did not survive package rewrite in " + part,
                )


def _header_footer_fragment(data: bytes) -> Optional[bytes]:
    prefix = rb"(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)"
    paired = re.search(
        rb"<" + prefix + rb"headerFooter\b[^>]*>.*?</(?P=prefix)headerFooter\s*>",
        data,
        re.DOTALL,
    )
    if paired:
        return paired.group(0)
    empty = re.search(rb"<" + prefix + rb"headerFooter\b[^>]*/\s*>", data)
    return empty.group(0) if empty else None


def _header_footer_target(
    source_sheet: str, operations: List[Dict[str, Any]]
) -> Tuple[Optional[str], List[Tuple[str, str]]]:
    """Resolve a source sheet through the edit plan and collect its text edits."""
    current_sheet = source_sheet
    replacements: List[Tuple[str, str]] = []
    for operation in operations:
        kind = operation.get("operation")
        if kind == "rename_sheet" and operation.get("old_name") == current_sheet:
            current_sheet = operation.get("new_name", current_sheet)
        elif kind == "delete_sheet" and operation.get("name") == current_sheet:
            return None, []
        elif (
            kind == "set_page_settings"
            and operation.get("sheet") == current_sheet
            and operation.get("settings", {}).get("footer") is not None
        ):
            # This operation explicitly replaces the odd footer through
            # Openpyxl; keep its requested value instead of restoring source XML.
            return None, []
        elif kind == "replace_header_text" and operation.get("sheet") == current_sheet:
            replacements.append((operation.get("search", ""), operation.get("replacement", "")))
    return current_sheet, replacements


def _replace_header_footer_fragment(
    worksheet_xml: bytes,
    fragment: Optional[bytes],
    output_sheet: str,
) -> bytes:
    prefix = rb"(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)"
    match = re.search(
        rb"<" + prefix + rb"headerFooter\b[^>]*>.*?</(?P=prefix)headerFooter\s*>",
        worksheet_xml,
        re.DOTALL,
    )
    if match is None:
        match = re.search(rb"<" + prefix + rb"headerFooter\b[^>]*/\s*>", worksheet_xml)
    if match is not None:
        if fragment is None:
            return worksheet_xml[: match.start()] + worksheet_xml[match.end() :]
        return worksheet_xml[: match.start()] + fragment + worksheet_xml[match.end() :]
    if fragment is None:
        return worksheet_xml

    # In SpreadsheetML headerFooter follows page settings and precedes row/column
    # breaks and extension records. Insert before the first later optional part.
    insertion = re.search(
        rb"<(?:[A-Za-z_][\w.-]*:)?(?:rowBreaks|colBreaks|extLst)\b",
        worksheet_xml,
    )
    if insertion is None:
        closing = re.search(rb"</(?:[A-Za-z_][\w.-]*:)?worksheet\s*>", worksheet_xml)
        if closing is None:
            raise WorkerError(
                "header_footer_restore_failed",
                "could not locate worksheet end for " + output_sheet,
            )
        position = closing.start()
    else:
        position = insertion.start()
    return worksheet_xml[:position] + fragment + worksheet_xml[position:]


def _replace_header_text_in_fragment(
    fragment: bytes, search: str, replacement: str
) -> bytes:
    """Replace a phrase in odd/even/first headers, leaving all footers intact."""
    if not search:
        return fragment
    pattern = re.compile(
        rb"(?P<open><(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)"
        rb"(?P<tag>oddHeader|evenHeader|firstHeader)\b[^>]*>)"
        rb"(?P<text>.*?)"
        rb"(?P<close></(?P=prefix)(?P=tag)\s*>)",
        re.DOTALL,
    )

    def xml_character_pattern(character: str) -> bytes:
        """Match a character in literal or numeric-reference XML form."""
        literal = re.escape(xml_escape(character).encode("utf-8"))
        decimal = b"&#0*" + str(ord(character)).encode("ascii") + b";"
        hexadecimal = "".join(
            "[{}{}]".format(digit.lower(), digit.upper())
            if digit.isalpha()
            else digit
            for digit in format(ord(character), "x")
        ).encode("ascii")
        hexadecimal = b"&#[xX]0*" + hexadecimal + b";"
        return b"(?:" + literal + b"|" + decimal + b"|" + hexadecimal + b")"

    search_pattern = re.compile(
        b"".join(xml_character_pattern(character) for character in search),
        re.DOTALL,
    )
    replacement_bytes = xml_escape(replacement).encode("utf-8")

    def replace_match(match: "re.Match[bytes]") -> bytes:
        return (
            match.group("open")
            + search_pattern.sub(lambda _: replacement_bytes, match.group("text"))
            + match.group("close")
        )

    return pattern.sub(replace_match, fragment)


def _restore_header_footer_settings(
    input_path: Path, output_path: Path, operations: List[Dict[str, Any]]
) -> None:
    """Restore exact header/footer OOXML, applying only planned text replacements.

    Openpyxl encodes line feeds differently and materializes empty even/first
    headers and footers. Keep the source markup (including whitespace and
    xml:space) and patch only requested text, so unrelated worksheet headers
    and footers remain intact.
    """
    with zipfile.ZipFile(str(input_path), "r") as original, zipfile.ZipFile(
        str(output_path), "r"
    ) as candidate:
        original_parts = _worksheet_parts_by_title(original)
        candidate_parts = _worksheet_parts_by_title(candidate)
        updates: Dict[str, bytes] = {}
        expectations: Dict[str, Optional[bytes]] = {}
        for source_sheet, source_part in original_parts.items():
            output_sheet, replacements = _header_footer_target(source_sheet, operations)
            if output_sheet is None:
                continue
            output_part = candidate_parts.get(output_sheet)
            if not output_part or output_part not in candidate.namelist():
                raise WorkerError(
                    "header_footer_restore_failed",
                    "worksheet is missing after openpyxl save: " + output_sheet,
                )
            source_xml = original.read(source_part)
            output_xml = candidate.read(output_part)
            expected = _header_footer_fragment(source_xml)
            if expected is not None:
                for search, replacement in replacements:
                    expected = _replace_header_text_in_fragment(expected, search, replacement)
            restored_xml = _replace_header_footer_fragment(output_xml, expected, output_sheet)
            if restored_xml != output_xml:
                updates[output_part] = restored_xml
                expectations[output_part] = expected

    if not updates:
        return
    replacement = output_path.with_name(output_path.name + ".header-footers.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "header/footer temporary output already exists")
    try:
        with zipfile.ZipFile(str(output_path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                target.writestr(item, updates.get(item.filename, source.read(item.filename)))
        os.replace(str(replacement), str(output_path))
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        if isinstance(error, WorkerError):
            raise
        raise WorkerError("save_failed", "could not restore header/footer settings: " + str(error))

    with zipfile.ZipFile(str(output_path), "r") as result:
        for part, expected in expectations.items():
            if _header_footer_fragment(result.read(part)) != expected:
                raise WorkerError(
                    "header_footer_restore_failed",
                    "header/footer settings did not survive package rewrite in " + part,
                )


def _restore_calculation_chain(
    input_path: Path, output_path: Path, operations: List[Dict[str, Any]]
) -> None:
    """Keep calcChain only when edits cannot stale formula results or dependencies.

    Any cell-value write may affect formula precedents, so Openpyxl's removed
    calculation chain must stay removed for those plans.
    """
    safe_operations = {
        "set_comment",
        "set_column_width",
        "set_row_height",
        "set_cell_fill_color",
        "set_page_settings",
        "set_freeze_panes",
        "set_auto_filter",
        "replace_header_text",
        "set_workbook_calculation_properties",
    }
    if any(operation.get("operation") not in safe_operations for operation in operations):
        return
    formula_cells = set()
    with zipfile.ZipFile(str(input_path), "r") as original:
        for title, part in _worksheet_parts_by_title(original).items():
            if part not in original.namelist():
                continue
            worksheet = ElementTree.fromstring(original.read(part))
            for cell in worksheet.iter(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
            ):
                address = cell.attrib.get("r")
                if address and cell.find(
                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f"
                ) is not None:
                    formula_cells.add((title, address.replace("$", "").upper()))
    for operation in operations:
        if operation.get("operation") == "set_cell_value":
            address = operation.get("cell", "").replace("$", "").upper()
            if (operation.get("sheet", ""), address) in formula_cells:
                return

    relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    content_types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    calc_chain_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
    )
    with zipfile.ZipFile(str(input_path), "r") as original:
        names = set(original.namelist())
        if "xl/_rels/workbook.xml.rels" not in names or "[Content_Types].xml" not in names:
            return
        original_relationships_xml = original.read("xl/_rels/workbook.xml.rels")
        original_relationships = ElementTree.fromstring(original_relationships_xml)
        source_relationships = [
            relationship
            for relationship in original_relationships.findall(
                "{%s}Relationship" % relationship_ns
            )
            if relationship.attrib.get("Type") == calc_chain_type
        ]
        if not source_relationships:
            return
        if len(source_relationships) != 1:
            raise WorkerError(
                "calculation_chain_restore_failed",
                "source workbook has multiple calculation-chain relationships",
            )
        source_relationship = source_relationships[0]
        target = source_relationship.attrib.get("Target", "")
        calc_chain_part = posixpath.normpath(
            target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
        )
        if calc_chain_part not in names:
            raise WorkerError(
                "calculation_chain_restore_failed",
                "calculation-chain relationship points to a missing package part",
            )
        calc_chain_xml = original.read(calc_chain_part)
        original_content_types = ElementTree.fromstring(original.read("[Content_Types].xml"))
        source_overrides = [
            override
            for override in original_content_types.findall("{%s}Override" % content_types_ns)
            if override.attrib.get("PartName") == "/" + calc_chain_part
        ]
        if len(source_overrides) != 1:
            raise WorkerError(
                "calculation_chain_restore_failed",
                "calculation-chain package part has no unique content type override",
            )
        source_content_type = source_overrides[0]

    updates: Dict[str, bytes] = {calc_chain_part: calc_chain_xml}
    additions: Dict[str, bytes] = {}
    with zipfile.ZipFile(str(output_path), "r") as candidate:
        output_names = set(candidate.namelist())
        output_relationships = ElementTree.fromstring(candidate.read("xl/_rels/workbook.xml.rels"))
        existing_calc_chain = [
            relationship
            for relationship in output_relationships.findall(
                "{%s}Relationship" % relationship_ns
            )
            if relationship.attrib.get("Type") == calc_chain_type
        ]
        if existing_calc_chain:
            if len(existing_calc_chain) != 1 or existing_calc_chain[0].attrib.get("Target") != target:
                raise WorkerError(
                    "calculation_chain_restore_failed",
                    "saved calculation-chain relationship conflicts with the source package",
                )
        else:
            source_id = source_relationship.attrib.get("Id", "")
            if not source_id or any(
                relationship.attrib.get("Id") == source_id
                for relationship in output_relationships.findall(
                    "{%s}Relationship" % relationship_ns
                )
            ):
                raise WorkerError(
                    "calculation_chain_restore_failed",
                    "source calculation-chain relationship id is already in use",
                )
            relationship = ElementTree.Element("{%s}Relationship" % relationship_ns)
            relationship.attrib.update(source_relationship.attrib)
            output_relationships.append(relationship)
        ElementTree.register_namespace("", relationship_ns)
        updates["xl/_rels/workbook.xml.rels"] = ElementTree.tostring(
            output_relationships, encoding="utf-8", xml_declaration=True
        )

        output_content_types = ElementTree.fromstring(candidate.read("[Content_Types].xml"))
        existing_overrides = [
            override
            for override in output_content_types.findall("{%s}Override" % content_types_ns)
            if override.attrib.get("PartName") == "/" + calc_chain_part
        ]
        if existing_overrides:
            if len(existing_overrides) != 1 or existing_overrides[0].attrib.get("ContentType") != source_content_type.attrib.get("ContentType"):
                raise WorkerError(
                    "calculation_chain_restore_failed",
                    "saved calculation-chain content type conflicts with the source package",
                )
        else:
            override = ElementTree.Element("{%s}Override" % content_types_ns)
            override.attrib.update(source_content_type.attrib)
            output_content_types.append(override)
        ElementTree.register_namespace("", content_types_ns)
        updates["[Content_Types].xml"] = ElementTree.tostring(
            output_content_types, encoding="utf-8", xml_declaration=True
        )
        if calc_chain_part in output_names:
            if candidate.read(calc_chain_part) != calc_chain_xml:
                updates[calc_chain_part] = calc_chain_xml
        else:
            additions[calc_chain_part] = calc_chain_xml

    replacement = output_path.with_name(output_path.name + ".calc-chain.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "calculation-chain temporary output already exists")
    try:
        with zipfile.ZipFile(str(output_path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target_package:
            for item in source.infolist():
                target_package.writestr(
                    item, updates.get(item.filename, source.read(item.filename))
                )
            for part_name, data in additions.items():
                target_package.writestr(part_name, data, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(str(replacement), str(output_path))
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise WorkerError("save_failed", "could not restore calculation chain: " + str(error))

    with zipfile.ZipFile(str(output_path), "r") as result:
        relationships = ElementTree.fromstring(result.read("xl/_rels/workbook.xml.rels"))
        if not any(
            relationship.attrib.get("Type") == calc_chain_type
            and relationship.attrib.get("Target") == target
            for relationship in relationships.findall("{%s}Relationship" % relationship_ns)
        ):
            raise WorkerError(
                "calculation_chain_restore_failed",
                "calculation-chain relationship did not survive package rewrite",
            )
        if result.read(calc_chain_part) != calc_chain_xml:
            raise WorkerError(
                "calculation_chain_restore_failed",
                "calculation-chain XML did not survive package rewrite",
            )


def _restore_planned_empty_strings(
    output_path: Path,
    operations: List[Dict[str, Any]],
) -> None:
    """Keep explicit empty text writes as inline strings after Openpyxl saves.

    Openpyxl reloads an empty inline string as ``None``. The OOXML marker still
    distinguishes an explicit empty string from a truly blank cell, which can
    matter to formulas such as COUNTA and ISBLANK.
    """
    cells: Dict[Tuple[str, str], bool] = {}
    for operation in operations:
        kind = operation.get("operation")
        if kind in ("delete_rows", "insert_rows", "copy_range", "import_range"):
            # These operations can move or replace cell coordinates. Leave the
            # package under the normal operation verifier instead of restoring
            # an empty-string marker at a potentially stale address.
            return
        if kind == "set_cell_value":
            value = operation.get("value", {})
            key = (
                operation.get("sheet", ""),
                operation.get("cell", "").replace("$", "").upper(),
            )
            cells[key] = value.get("type") == "text" and value.get("value") == ""
        elif kind == "set_formula_text":
            cells.pop(
                (
                    operation.get("sheet", ""),
                    operation.get("cell", "").replace("$", "").upper(),
                ),
                None,
            )
        elif kind == "rename_sheet":
            old_name = operation.get("old_name", "")
            new_name = operation.get("new_name", "")
            renamed = {
                (new_name if sheet == old_name else sheet, address): empty
                for (sheet, address), empty in cells.items()
            }
            cells = renamed
        elif kind == "delete_sheet":
            deleted_name = operation.get("name", "")
            cells = {
                key: empty for key, empty in cells.items() if key[0] != deleted_name
            }
    targets: Dict[str, Set[str]] = {}
    for (sheet, address), is_empty in cells.items():
        if is_empty:
            targets.setdefault(sheet, set()).add(address)
    if not targets:
        return

    def set_empty_inline_string(xml_data: bytes, address: str) -> bytes:
        coordinate = re.escape(address.encode("ascii"))
        pattern = re.compile(
            rb"<c\b(?=[^>]*\br=(?:\""
            + coordinate
            + rb"\"|'"
            + coordinate
            + rb"'))[^>]*?(?:/>|>.*?</(?:[A-Za-z_][\w.-]*:)?c>)",
            re.DOTALL,
        )
        matches = list(pattern.finditer(xml_data))
        if len(matches) != 1:
            raise WorkerError(
                "empty_string_restore_failed",
                "could not uniquely locate empty-text cell " + address,
            )
        match = matches[0]
        element = match.group(0)
        opening_end = element.find(b">") + 1
        if opening_end == 0:
            raise WorkerError(
                "empty_string_restore_failed", "invalid cell XML at " + address
            )
        opening = element[:opening_end]
        self_closing = opening.rstrip().endswith(b"/>")
        if self_closing:
            opening = opening.rstrip()[:-2] + b">"
        opening = re.sub(
            rb"\s+t=(?:\"[^\"]*\"|'[^']*')", b"", opening, count=1
        )
        opening = opening[:-1] + b' t="inlineStr">'
        tag_match = re.match(rb"<(?P<tag>(?:[A-Za-z_][\w.-]*:)?c)\b", opening)
        if tag_match is None:
            raise WorkerError(
                "empty_string_restore_failed", "invalid cell element at " + address
            )
        closing_tag = b"</" + tag_match.group("tag") + b">"
        content = b"" if self_closing else element[opening_end:]
        if not self_closing:
            closing_start = content.rfind(closing_tag)
            if closing_start < 0:
                raise WorkerError(
                    "empty_string_restore_failed",
                    "cell closing tag is missing at " + address,
                )
            content = content[:closing_start]
            content = re.sub(
                rb"<(?:[A-Za-z_][\w.-]*:)?(?:is|v)\b[^>]*/>|"
                rb"<(?:[A-Za-z_][\w.-]*:)?(?:is|v)\b[^>]*>.*?</(?:[A-Za-z_][\w.-]*:)?(?:is|v)\s*>",
                b"",
                content,
                flags=re.DOTALL,
            )
        replacement = opening + b"<is><t></t></is>" + closing_tag
        return xml_data[:match.start()] + replacement + xml_data[match.end():]

    replacement = output_path.with_name(output_path.name + ".empty-strings.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "empty-string temporary output already exists")
    try:
        with zipfile.ZipFile(str(output_path), "r") as source:
            title_to_part = _worksheet_parts_by_title(source)
            updates: Dict[str, bytes] = {}
            for sheet, addresses in targets.items():
                part = title_to_part.get(sheet)
                if part is None or part not in source.namelist():
                    raise WorkerError(
                        "empty_string_restore_failed",
                        "worksheet is missing for empty-text writes: " + sheet,
                    )
                xml_data = source.read(part)
                for address in addresses:
                    xml_data = set_empty_inline_string(xml_data, address)
                updates[part] = xml_data
        with zipfile.ZipFile(str(output_path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                target.writestr(item, updates.get(item.filename, source.read(item.filename)))
        os.replace(str(replacement), str(output_path))
    except WorkerError:
        raise
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise WorkerError(
            "save_failed", "could not restore explicit empty strings: " + str(error)
        )

    with zipfile.ZipFile(str(output_path), "r") as result:
        for sheet, addresses in targets.items():
            part = _worksheet_parts_by_title(result).get(sheet)
            if part is None:
                raise WorkerError(
                    "empty_string_restore_failed",
                    "worksheet disappeared while restoring empty text: " + sheet,
                )
            root = ElementTree.fromstring(result.read(part))
            cells_by_address = {
                cell.attrib.get("r", "").replace("$", "").upper(): cell
                for cell in root.iter(
                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
                )
            }
            for address in addresses:
                cell = cells_by_address.get(address)
                if cell is None or cell.attrib.get("t") != "inlineStr":
                    raise WorkerError(
                        "empty_string_restore_failed",
                        "explicit empty-string marker did not survive at "
                        + sheet
                        + "!"
                        + address,
                    )


def _metadata_cell_attributes(path: Path) -> Dict[Tuple[str, str], Dict[str, str]]:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(str(path), "r") as package:
        result: Dict[Tuple[str, str], Dict[str, str]] = {}
        for title, part in _worksheet_parts_by_title(package).items():
            if part not in package.namelist():
                continue
            root = ElementTree.fromstring(package.read(part))
            for cell in root.iter("{%s}c" % main_ns):
                attributes = {
                    name: cell.attrib[name]
                    for name in ("cm", "vm", "ph")
                    if name in cell.attrib
                }
                address = cell.attrib.get("r")
                if address and attributes:
                    result[(title, address.replace("$", "").upper())] = attributes
        return result


def _workbook_root_extension(data: bytes) -> Optional[bytes]:
    match = re.search(
        rb"<(?P<name>(?:[A-Za-z_][\w.-]*:)?extLst)\b.*?</(?P=name)\s*>",
        data,
        re.DOTALL,
    )
    return match.group(0) if match else None


def _restore_workbook_root_extensions(original: bytes, output: bytes) -> bytes:
    original_root = re.search(rb"<workbook\b[^>]*>", original)
    output_root = re.search(rb"<workbook\b[^>]*>", output)
    if original_root is None or output_root is None:
        raise WorkerError("metadata_restore_failed", "could not locate workbook XML root")

    original_attrs = re.findall(
        rb"\s+(?P<name>xmlns(?::[A-Za-z_][\w.-]*)?|mc:Ignorable)=(?P<quote>[\"'])(?P<value>.*?)(?P=quote)",
        original_root.group(0),
    )
    root_text = output_root.group(0)
    additions = []
    for name, _quote, value in original_attrs:
        name_pattern = re.escape(name) + rb"=(?:[\"'])(.*?)(?:[\"'])"
        existing = re.search(name_pattern, root_text)
        if existing:
            current_value = existing.group(1)
            if name == b"mc:Ignorable":
                merged = list(dict.fromkeys(current_value.split() + value.split()))
                replacement = b" ".join(merged)
                if replacement != current_value:
                    root_text = re.sub(
                        name_pattern,
                        name + b'="' + replacement + b'"',
                        root_text,
                        count=1,
                    )
            elif current_value != value:
                raise WorkerError(
                    "metadata_restore_failed",
                    "workbook namespace prefix conflicts with the source package: " + name.decode("ascii"),
                )
        else:
            additions.append(b" " + name + b'="' + value + b'"')
    if additions:
        root_text = root_text[:-1] + b"".join(additions) + b">"
    output = output[: output_root.start()] + root_text + output[output_root.end() :]

    original_extension = _workbook_root_extension(original)
    output_extension_match = re.search(
        rb"<(?P<name>(?:[A-Za-z_][\w.-]*:)?extLst)\b.*?</(?P=name)\s*>",
        output,
        re.DOTALL,
    )
    if output_extension_match:
        output = (
            output[: output_extension_match.start()]
            + (original_extension or b"")
            + output[output_extension_match.end() :]
        )
    elif original_extension:
        closing = re.search(rb"</workbook\s*>", output)
        if closing is None:
            raise WorkerError("metadata_restore_failed", "could not locate workbook XML closing tag")
        output = output[: closing.start()] + original_extension + output[closing.start() :]
    return output


def _restore_cell_metadata_attributes(data: bytes, address: str, attributes: Dict[str, str]) -> bytes:
    encoded_address = re.escape(address.encode("ascii"))
    pattern = re.compile(
        rb"<c\b(?=[^>]*\br=[\"']" + encoded_address + rb"[\"'])[^>]*>",
        re.DOTALL,
    )
    matches = list(pattern.finditer(data))
    if len(matches) != 1:
        raise WorkerError(
            "metadata_restore_failed",
            "could not uniquely locate cell-metadata cell " + address,
        )
    match = matches[0]
    start_tag = match.group(0)
    for name, value in attributes.items():
        if not value.isdigit():
            raise WorkerError("metadata_restore_failed", "cell metadata index is not numeric")
        attribute = name.encode("ascii")
        attribute_pattern = re.compile(rb"\s+" + attribute + rb"=[\"'][^\"']*[\"']")
        replacement = b" " + attribute + b'="' + value.encode("ascii") + b'"'
        if attribute_pattern.search(start_tag):
            start_tag = attribute_pattern.sub(replacement, start_tag, count=1)
        else:
            insertion = start_tag.rfind(b"/>")
            if insertion < 0:
                insertion = start_tag.rfind(b">")
            if insertion < 0:
                raise WorkerError("metadata_restore_failed", "invalid cell XML start tag")
            start_tag = start_tag[:insertion] + replacement + start_tag[insertion:]
    return data[: match.start()] + start_tag + data[match.end() :]


def _restore_cell_metadata_package_parts(
    input_path: Path,
    output_path: Path,
    operations: List[Dict[str, Any]],
) -> None:
    metadata_part = "xl/metadata.xml"
    workbook_part = "xl/workbook.xml"
    relationships_part = "xl/_rels/workbook.xml.rels"
    content_types_part = "[Content_Types].xml"
    relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    content_types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    sheet_metadata_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
    )
    sheet_metadata_content_type = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
    )

    with zipfile.ZipFile(str(input_path), "r") as original:
        original_names = set(original.namelist())
        if metadata_part not in original_names:
            return
        metadata_bytes = original.read(metadata_part)
        original_workbook_xml = original.read(workbook_part)
        original_relationships_xml = original.read(relationships_part)
        original_content_types_xml = original.read(content_types_part)
        original_relationships = ElementTree.fromstring(original_relationships_xml)
        original_metadata_relationships = [
            relationship
            for relationship in original_relationships.findall("{%s}Relationship" % relationship_ns)
            if relationship.attrib.get("Type") == sheet_metadata_type
        ]
        if len(original_metadata_relationships) != 1:
            raise WorkerError(
                "metadata_restore_failed",
                "cell metadata part does not have exactly one workbook relationship",
            )
        original_metadata_relationship = original_metadata_relationships[0]
        target = original_metadata_relationship.attrib.get("Target", "")
        resolved_target = posixpath.normpath(
            target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
        )
        if resolved_target != metadata_part:
            raise WorkerError(
                "metadata_restore_failed",
                "unsupported cell metadata package target: " + target,
                unsupported=True,
            )

        original_content_types = ElementTree.fromstring(original_content_types_xml)
        metadata_overrides = [
            item
            for item in original_content_types.findall("{%s}Override" % content_types_ns)
            if item.attrib.get("PartName") == "/xl/metadata.xml"
        ]
        if len(metadata_overrides) != 1 or metadata_overrides[0].attrib.get("ContentType") != sheet_metadata_content_type:
            raise WorkerError(
                "metadata_restore_failed",
                "cell metadata part has an unexpected content type",
            )
        expected_cell_attributes = _metadata_cell_attributes(input_path)

    updates: Dict[str, bytes] = {}
    additions: Dict[str, bytes] = {}
    with zipfile.ZipFile(str(output_path), "r") as candidate:
        output_names = set(candidate.namelist())
        output_sheet_parts = _worksheet_parts_by_title(candidate)
        output_relationships = ElementTree.fromstring(candidate.read(relationships_part))
        existing_metadata_relationship = next(
            (
                relationship
                for relationship in output_relationships.findall("{%s}Relationship" % relationship_ns)
                if relationship.attrib.get("Type") == sheet_metadata_type
            ),
            None,
        )
        if existing_metadata_relationship is None:
            source_relationship = original_metadata_relationship
            used_ids = {
                relationship.attrib.get("Id", "")
                for relationship in output_relationships.findall("{%s}Relationship" % relationship_ns)
            }
            source_id = source_relationship.attrib.get("Id", "")
            if source_id and source_id not in used_ids:
                relationship_id = source_id
            else:
                sequence = 1
                while "rId{}".format(sequence) in used_ids:
                    sequence += 1
                relationship_id = "rId{}".format(sequence)
            relationship = ElementTree.Element("{%s}Relationship" % relationship_ns)
            relationship.attrib.update(source_relationship.attrib)
            relationship.attrib["Id"] = relationship_id
            output_relationships.append(relationship)
        ElementTree.register_namespace("", relationship_ns)
        updates[relationships_part] = ElementTree.tostring(
            output_relationships, encoding="utf-8", xml_declaration=True
        )

        output_content_types = ElementTree.fromstring(candidate.read(content_types_part))
        if not any(
            item.attrib.get("PartName") == "/xl/metadata.xml"
            for item in output_content_types.findall("{%s}Override" % content_types_ns)
        ):
            source_override = metadata_overrides[0]
            override = ElementTree.Element("{%s}Override" % content_types_ns)
            override.attrib.update(source_override.attrib)
            output_content_types.append(override)
        ElementTree.register_namespace("", content_types_ns)
        updates[content_types_part] = ElementTree.tostring(
            output_content_types, encoding="utf-8", xml_declaration=True
        )

        updates[workbook_part] = _restore_workbook_root_extensions(
            original_workbook_xml, candidate.read(workbook_part)
        )
        for (sheet, address), attributes in expected_cell_attributes.items():
            destination = _formula_cache_destination(sheet, address, operations)
            if destination is None:
                continue
            output_sheet, output_address = destination
            output_sheet_part = output_sheet_parts.get(output_sheet)
            if not output_sheet_part or output_sheet_part not in output_names:
                raise WorkerError(
                    "metadata_restore_failed",
                    "worksheet with cell metadata is missing after save: " + output_sheet,
                )
            worksheet_xml = updates.get(output_sheet_part, candidate.read(output_sheet_part))
            updates[output_sheet_part] = _restore_cell_metadata_attributes(
                worksheet_xml, output_address, attributes
            )
        if metadata_part in output_names:
            updates[metadata_part] = metadata_bytes
        else:
            additions[metadata_part] = metadata_bytes

    replacement = output_path.with_name(output_path.name + ".metadata.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "cell-metadata temporary output already exists")
    try:
        with zipfile.ZipFile(str(output_path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                target.writestr(item, updates.get(item.filename, source.read(item.filename)))
            for part_name, data in additions.items():
                target.writestr(part_name, data, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(str(replacement), str(output_path))
    except WorkerError:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise WorkerError("save_failed", "could not preserve cell metadata: " + str(error))

    with zipfile.ZipFile(str(output_path), "r") as result:
        names = set(result.namelist())
        if metadata_part not in names or result.read(metadata_part) != metadata_bytes:
            raise WorkerError("metadata_restore_failed", "cell metadata part differs after restore")
        relationships = ElementTree.fromstring(result.read(relationships_part))
        if not any(
            item.attrib.get("Type") == sheet_metadata_type
            and posixpath.normpath(
                item.attrib.get("Target", "").lstrip("/")
                if item.attrib.get("Target", "").startswith("/")
                else posixpath.join("xl", item.attrib.get("Target", ""))
            ) == metadata_part
            for item in relationships.findall("{%s}Relationship" % relationship_ns)
        ):
            raise WorkerError("metadata_restore_failed", "cell metadata relationship is missing")
        content_types = ElementTree.fromstring(result.read(content_types_part))
        if not any(
            item.attrib.get("PartName") == "/xl/metadata.xml"
            and item.attrib.get("ContentType") == sheet_metadata_content_type
            for item in content_types.findall("{%s}Override" % content_types_ns)
        ):
            raise WorkerError("metadata_restore_failed", "cell metadata content type is missing")
        output_sheet_parts = _worksheet_parts_by_title(result)
        for (sheet, address), attributes in expected_cell_attributes.items():
            destination = _formula_cache_destination(sheet, address, operations)
            if destination is None:
                continue
            output_part = output_sheet_parts.get(destination[0])
            if not output_part or output_part not in names:
                raise WorkerError("metadata_restore_failed", "cell-metadata worksheet is missing")
            xml = ElementTree.fromstring(result.read(output_part))
            cell = xml.find(
                ".//{%s}c[@r='%s']" % (
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    destination[1],
                )
            )
            if cell is None or any(cell.attrib.get(name) != value for name, value in attributes.items()):
                raise WorkerError(
                    "metadata_restore_failed",
                    "cell metadata readback mismatch for {}!{}".format(*destination),
                )
        original_extension = _workbook_root_extension(original_workbook_xml)
        if original_extension and original_extension not in result.read(workbook_part):
            raise WorkerError("metadata_restore_failed", "workbook extension metadata was not restored")


def _preflight(request: Dict[str, Any]) -> Tuple[Path, Path, List[Dict[str, Any]]]:
    if request.get("schema_version") != SCHEMA_VERSION:
        raise WorkerError("protocol_version", "unsupported request schema", unsupported=True)
    request_id = request.get("request_id")
    if not isinstance(request_id, str) or not request_id.strip():
        raise WorkerError("invalid_request", "request_id must be a non-empty string", unsupported=True)
    if request.get("format") != "xlsx":
        raise WorkerError("unsupported_format", "openpyxl worker accepts XLSX only", unsupported=True)
    input_path = Path(request.get("input", "")).resolve(strict=True)
    output_path = Path(request.get("output", "")).resolve(strict=False)
    if input_path.suffix.lower() != ".xlsx" or output_path.suffix.lower() != ".xlsx":
        raise WorkerError("unsupported_format", "input and output paths must end in .xlsx", unsupported=True)
    if input_path == output_path:
        raise WorkerError("invalid_paths", "input and output paths must be different", unsupported=True)
    if not input_path.is_file():
        raise WorkerError("input_missing", "input workbook does not exist", unsupported=True)
    if output_path.exists() and output_path.stat().st_size != 0:
        raise WorkerError("output_exists", "worker output must be a new or empty candidate", unsupported=True)
    if not zipfile.is_zipfile(str(input_path)):
        raise WorkerError("unsupported_format", "input is not a valid XLSX ZIP package", unsupported=True)
    plan = request.get("plan")
    if not isinstance(plan, dict) or plan.get("request_id") != request_id:
        raise WorkerError("invalid_request", "embedded ExcelEditPlan does not match request_id", unsupported=True)
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise WorkerError("invalid_request", "ExcelEditPlan has no operations", unsupported=True)
    unsupported = sorted({
        str(operation.get("operation", "<missing>")) if isinstance(operation, dict) else "<invalid>"
        for operation in operations
        if not isinstance(operation, dict) or operation.get("operation") not in SUPPORTED_OPERATIONS
    })
    if unsupported:
        raise WorkerError(
            "unsupported_operation",
            "openpyxl worker does not support the complete batch: " + ", ".join(unsupported),
            unsupported=True,
        )
    has_row_shift = any(
        operation["operation"] in ("delete_rows", "insert_rows")
        for operation in operations
    )
    has_range_copy = any(
        operation["operation"] in ("copy_range", "import_range")
        for operation in operations
    )
    if has_row_shift and has_range_copy:
        raise WorkerError(
            "unsupported_operation",
            "openpyxl worker cannot combine row insertion/deletion with range copy in one plan",
            unsupported=True,
        )
    workbook = plan.get("workbook") or {}
    if not workbook.get("inventory_complete", False):
        raise WorkerError("unknown_package_parts", "OOXML inventory is incomplete", unsupported=True)
    formula_cache_conflict = _formula_cache_write_conflict(input_path, operations, request_id)
    if formula_cache_conflict:
        raise WorkerError(
            "formula_cache_not_preserved",
            formula_cache_conflict,
            unsupported=True,
        )
    for operation in operations:
        if operation["operation"] == "import_range":
            value = operation.get("source_workbook", "")
            source = Path(value).resolve(strict=True)
            if source.suffix.lower() != ".xlsx" or not source.is_file() or not zipfile.is_zipfile(str(source)):
                raise WorkerError("unsupported_import_source", "ImportRange currently accepts valid XLSX sources only", unsupported=True)
    return input_path, output_path, operations


def _cell_value(value: Dict[str, Any]) -> Any:
    kind = value.get("type")
    content = value.get("value")
    if kind == "text":
        return content
    if kind == "decimal":
        return Decimal(str(content))
    if kind == "date_time":
        return Decimal(str(content["serial"]))
    if kind == "boolean":
        return bool(content)
    if kind == "error":
        return content
    if kind == "blank":
        return None
    raise WorkerError("invalid_operation", "unknown CellValue variant: " + str(kind))


def _preserve_significant_string_whitespace(path: Path) -> None:
    """Restore OOXML's xml:space marker after openpyxl serializes string cells.

    openpyxl preserves leading/trailing spaces when it reads an inline string,
    but its writer can omit xml:space="preserve". That changes OOXML's declared
    whitespace semantics and can make Rust readers observe a different value.
    Re-serialize only worksheet/shared-string XML parts that contain such text.
    """
    xml_space = "http://www.w3.org/XML/1998/namespace}space"
    main_text = "http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
    replacement = path.with_name(path.name + ".whitespace.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "whitespace-preservation temporary output already exists")

    def with_namespaces(data: bytes) -> bytes:
        parser = expat.ParserCreate(namespace_separator="}")
        patches = []
        active_text = []

        def start_element(name: str, attributes: Dict[str, str]) -> None:
            if name == main_text:
                active_text.append((parser.CurrentByteIndex, attributes, []))

        def character_data(value: str) -> None:
            if active_text:
                active_text[-1][2].append(value)

        def end_element(name: str) -> None:
            if name != main_text:
                return
            start, attributes, chunks = active_text.pop()
            value = "".join(chunks)
            if not value or (value[0] not in " \t\r\n" and value[-1] not in " \t\r\n"):
                return
            existing = attributes.get(xml_space)
            if existing == "preserve":
                return
            if existing is not None:
                raise WorkerError(
                    "unsupported_whitespace_semantics",
                    "cannot safely change a non-preserve xml:space attribute",
                    unsupported=True,
                )
            quote = 0
            end = None
            for index in range(start, len(data)):
                byte = data[index]
                if quote:
                    if byte == quote:
                        quote = 0
                elif byte in (ord('"'), ord("'")):
                    quote = byte
                elif byte == ord(">"):
                    end = index
                    break
            if end is None:
                raise WorkerError("invalid_xml", "could not locate OOXML text element end")
            patches.append((end, b' xml:space="preserve"'))

        parser.StartElementHandler = start_element
        parser.CharacterDataHandler = character_data
        parser.EndElementHandler = end_element
        parser.Parse(data, True)
        if not patches:
            return data
        result = data
        for position, insertion in reversed(patches):
            result = result[:position] + insertion + result[position:]
        return result

    try:
        changed_parts = set()
        with zipfile.ZipFile(str(path), "r") as source:
            for item in source.infolist():
                if not (
                    item.filename.startswith("xl/worksheets/")
                    or item.filename == "xl/sharedStrings.xml"
                ):
                    continue
                data = source.read(item.filename)
                if with_namespaces(data) != data:
                    changed_parts.add(item.filename)
        if not changed_parts:
            return
        with zipfile.ZipFile(str(path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename in changed_parts:
                    data = with_namespaces(data)
                target.writestr(item, data)
        os.replace(str(replacement), str(path))
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        if isinstance(error, WorkerError):
            raise
        raise WorkerError("save_failed", "could not preserve significant string whitespace: " + str(error))


def _translate_formula(value: Any, source: Any, target: Any) -> Any:
    if not isinstance(value, str) or not value.startswith("="):
        return value
    from openpyxl.formula.translate import Translator
    origin = source if isinstance(source, str) else source.coordinate
    return Translator(value, origin=origin).translate_formula(target.coordinate)


def _copy_style(source: Any, target: Any, style_mode: str) -> None:
    if style_mode == "none":
        return
    if style_mode == "full":
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)
        return
    if style_mode == "basic":
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        return
    raise WorkerError("invalid_operation", "unknown range style mode: " + str(style_mode))


def _copy_comment(source: Any, target: Any) -> None:
    if source.comment is None:
        target.comment = None
        return
    from openpyxl.comments import Comment
    target.comment = Comment(source.comment.text, source.comment.author)


def _range_snapshot(source: Any, data_only_source: Any, address: str) -> Tuple[Any, Any, Any, Any]:
    from openpyxl.utils.cell import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(address)
    if not min_col or not min_row or not max_col or not max_row:
        raise WorkerError("invalid_range", "range must have finite cell boundaries")
    if max_row < min_row or max_col < min_col:
        raise WorkerError("invalid_range", "range end precedes range start")
    return min_row, min_col, max_row, max_col


def _copy_cells(source: Any, cached: Any, target: Any, bounds: Tuple[int, int, int, int],
                target_start: str, copy_values: bool, copy_formulas: bool,
                style_mode: str, copy_comments: bool, materialize_blanks: bool,
                hide_zero: bool, materialized: Optional[Dict[int, Set[str]]] = None) -> None:
    from openpyxl.utils.cell import coordinate_to_tuple
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
    min_row, min_col, max_row, max_col = bounds
    target_row, target_col = coordinate_to_tuple(target_start.replace("$", ""))
    # Snapshot first so overlapping ranges on the same worksheet behave like
    # copy/paste rather than reading values already written by this operation.
    snapshots = []
    for row in source.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for src in row:
            value = src.value
            if isinstance(value, (ArrayFormula, DataTableFormula)):
                raise WorkerError("unsupported_formula_kind", "array/data-table formula copying is not supported", unsupported=True)
            if value is None and style_mode == "none" and not copy_comments and not materialize_blanks:
                continue
            cached_value = cached.cell(src.row, src.column).value
            snapshots.append((
                src.coordinate,
                value,
                cached_value,
                (copy(src.font), copy(src.fill), copy(src.border), copy(src.alignment),
                 src.number_format, copy(src.protection)) if style_mode != "none" else None,
                None if src.comment is None else (src.comment.text, src.comment.author),
            ))
    for source_coordinate, value, cached_value, style, comment in snapshots:
        source_row, source_col = coordinate_to_tuple(source_coordinate)
        dst = target.cell(target_row + source_row - min_row, target_col + source_col - min_col)
        if style is not None:
            dst.font, dst.fill, dst.border, dst.alignment, dst.number_format, dst.protection = style
        if copy_formulas and isinstance(value, str) and value.startswith("="):
            dst.value = _translate_formula(value, source_coordinate, dst)
        elif copy_values:
            copied = cached_value if isinstance(value, str) and value.startswith("=") else value
            from decimal import Decimal
            if hide_zero and isinstance(copied, (int, float, Decimal)) and copied == 0:
                copied = None
            dst.value = copied
            if isinstance(copied, str) and copied.startswith("#"):
                dst.data_type = "e"
        if copy_comments:
            if comment is None:
                dst.comment = None
            else:
                from openpyxl.comments import Comment
                dst.comment = Comment(comment[0], comment[1])
        if materialized is not None:
            # openpyxl 的写入器会跳过“无值且样式为默认”的单元格；记录显式
            # 具体化的坐标，保存后由 _materialize_tracked_cells 补写，
            # 保证读取端能看到与计划一致的存在性。
            materialized.setdefault(id(target), set()).add(dst.coordinate)


def _copy_merges(source: Any, target: Any, bounds: Tuple[int, int, int, int],
                 target_start: str) -> None:
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
    min_row, min_col, max_row, max_col = bounds
    target_row, target_col = coordinate_to_tuple(target_start.replace("$", ""))
    for merged in source.merged_cells.ranges:
        if not (merged.min_row <= max_row and merged.max_row >= min_row
                and merged.min_col <= max_col and merged.max_col >= min_col):
            continue
        if not (min_row <= merged.min_row and merged.max_row <= max_row
                and min_col <= merged.min_col and merged.max_col <= max_col):
            raise WorkerError("partial_merge_copy", "source range cuts through a merged range", unsupported=True)
        offset_row = target_row - min_row
        offset_col = target_col - min_col
        new_range = "{}{}:{}{}".format(
            get_column_letter(merged.min_col + offset_col), merged.min_row + offset_row,
            get_column_letter(merged.max_col + offset_col), merged.max_row + offset_row,
        )
        if any(existing.coord == new_range for existing in target.merged_cells.ranges):
            continue
        # openpyxl 的 merge_cells 会把非锚点单元格替换为无样式的 MergedCell，
        # 抹掉 _copy_cells 刚从源复制来的成员样式；先快照、合并后恢复，
        # 保证输出与源工作簿的成员样式语义一致。
        anchor_coordinate = "{}{}".format(
            get_column_letter(merged.min_col + offset_col),
            merged.min_row + offset_row,
        )
        member_styles: Dict[str, Any] = {}
        for row in target[new_range]:
            for cell in row:
                if cell.coordinate != anchor_coordinate:
                    member_styles[cell.coordinate] = cell._style
        target.merge_cells(new_range)
        for row in target[new_range]:
            for cell in row:
                if cell.coordinate != anchor_coordinate and cell.coordinate in member_styles:
                    cell._style = member_styles[cell.coordinate]


def _copy_dimensions(source: Any, target: Any, bounds: Tuple[int, int, int, int],
                     target_start: str, operation: Dict[str, Any]) -> None:
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, column_index_from_string
    min_row, min_col, max_row, max_col = bounds
    target_row, target_col = coordinate_to_tuple(target_start.replace("$", ""))
    row_offset = target_row - min_row
    col_offset = target_col - min_col
    all_rows = operation.get("copy_all_row_dimensions", False)
    all_columns = operation.get("copy_all_column_dimensions", False)
    if operation.get("copy_row_heights") or operation.get("copy_row_hidden"):
        for row, dimension in source.row_dimensions.items():
            if not all_rows and not min_row <= row <= max_row:
                continue
            destination = target.row_dimensions[row + row_offset]
            if operation.get("copy_row_heights"):
                destination.height = dimension.height
            if operation.get("copy_row_hidden"):
                destination.hidden = dimension.hidden
            if operation.get("copy_outline_levels"):
                destination.outlineLevel = dimension.outlineLevel
    if operation.get("copy_column_widths") or operation.get("copy_column_hidden"):
        for key, dimension in source.column_dimensions.items():
            first = dimension.min or column_index_from_string(key)
            last = dimension.max or first
            if not all_columns and (last < min_col or first > max_col):
                continue
            first = max(first, min_col) if not all_columns else first
            last = min(last, max_col) if not all_columns else last
            for column in range(first, last + 1):
                destination_col = column + col_offset
                destination = target.column_dimensions[get_column_letter(destination_col)]
                if operation.get("copy_column_widths"):
                    destination.width = dimension.width
                if operation.get("copy_column_hidden"):
                    destination.hidden = dimension.hidden
                if operation.get("copy_outline_levels"):
                    destination.outlineLevel = dimension.outlineLevel


def _apply_operation(workbook: Any, cached_workbook: Any, operation: Dict[str, Any],
                     materialized: Optional[Dict[int, Set[str]]] = None) -> None:
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill
    from openpyxl.utils.cell import get_column_letter, coordinate_to_tuple, range_boundaries

    kind = operation["operation"]
    payload = operation
    if kind == "delete_rows":
        workbook[payload["sheet"]].delete_rows(
            int(payload["first_row"]), int(payload["count"])
        )
    elif kind == "insert_rows":
        workbook[payload["sheet"]].insert_rows(
            int(payload["first_row"]), amount=int(payload["count"])
        )
    elif kind == "set_cell_value":
        cell = workbook[payload["sheet"]][payload["cell"]]
        value = _cell_value(payload["value"])
        cell.value = value
        if payload["value"].get("type") == "date_time":
            cell.number_format = payload["value"]["value"]["number_format"]
        if payload["value"].get("type") == "error":
            cell.data_type = "e"
    elif kind == "set_formula_text":
        formula = payload["formula"]
        workbook[payload["sheet"]][payload["cell"]].value = formula if formula.startswith("=") else "=" + formula
    elif kind == "set_comment":
        cell = workbook[payload["sheet"]][payload["cell"]]
        text = payload.get("text")
        cell.comment = Comment(text, payload.get("author") or "pytools") if text is not None else None
    elif kind == "set_column_width":
        workbook[payload["sheet"]].column_dimensions[get_column_letter(payload["column_index"])].width = payload["width"]
    elif kind == "set_row_height":
        workbook[payload["sheet"]].row_dimensions[payload["row_index"]].height = payload["height"]
    elif kind == "rename_sheet":
        workbook[payload["old_name"]].title = payload["new_name"]
    elif kind == "create_sheet":
        workbook.create_sheet(payload["name"])
    elif kind == "create_sheet_at":
        workbook.create_sheet(payload["name"], payload["index"])
    elif kind == "set_freeze_panes":
        workbook[payload["sheet"]].freeze_panes = payload.get("cell")
    elif kind == "set_auto_filter":
        workbook[payload["sheet"]].auto_filter.ref = payload.get("range")
    elif kind == "delete_sheet":
        if len(workbook.worksheets) <= 1:
            raise WorkerError("invalid_operation", "cannot delete the only worksheet")
        del workbook[payload["name"]]
    elif kind == "set_cell_fill_color":
        color = payload["argb"].strip().lstrip("#")
        if len(color) == 6:
            color = "FF" + color
        if len(color) != 8 or any(char not in "0123456789abcdefABCDEF" for char in color):
            raise WorkerError("invalid_operation", "fill color must be RGB or ARGB hexadecimal")
        workbook[payload["sheet"]][payload["cell"]].fill = PatternFill(fill_type="solid", fgColor=color.upper())
    elif kind in ("copy_range", "import_range"):
        if kind == "copy_range":
            source_workbook = workbook
            cached_source = cached_workbook
            source_name = payload["source_sheet"]
            source_range = payload["source_range"]
            style_mode = "full" if payload.get("copy_styles") else "none"
            copy_comments = payload.get("copy_comments", False)
            copy_merges = False
            dimension_options: Dict[str, Any] = {}
            materialize = False
            hide_zero = False
        else:
            source_workbook = __import__("openpyxl").load_workbook(payload["source_workbook"], data_only=False, keep_links=True)
            cached_source = __import__("openpyxl").load_workbook(payload["source_workbook"], data_only=True, keep_links=True)
            source_name = payload["source_sheet"]
            source_range = payload["source_range"]
            style_mode = payload.get("style_mode", "none")
            copy_comments = payload.get("copy_comments", False)
            copy_merges = payload.get("copy_merges", False)
            dimension_options = payload
            materialize = payload.get("materialize_blank_cells", False)
            hide_zero = payload.get("hide_zero", False)
        source = source_workbook[source_name]
        cached_source_sheet = cached_source[source_name]
        target = workbook[payload["target_sheet"]]
        bounds = _range_snapshot(source, cached_source_sheet, source_range)
        _copy_cells(
            source, cached_source_sheet, target, bounds, payload["target_start"],
            payload.get("copy_values", False), payload.get("copy_formulas", False),
            style_mode, copy_comments, materialize, hide_zero, materialized,
        )
        if copy_merges:
            _copy_merges(source, target, bounds, payload["target_start"])
        if dimension_options:
            _copy_dimensions(source, target, bounds, payload["target_start"], dimension_options)
        if kind == "import_range":
            source_workbook.close()
            cached_source.close()
    elif kind == "set_page_settings":
        sheet = workbook[payload["sheet"]]
        settings = payload["settings"]
        if settings.get("orientation") is not None:
            sheet.page_setup.orientation = settings["orientation"]
        if settings.get("fit_to_page") is not None:
            if sheet.sheet_properties.pageSetUpPr is None:
                from openpyxl.worksheet.properties import PageSetupProperties
                sheet.sheet_properties.pageSetUpPr = PageSetupProperties()
            sheet.sheet_properties.pageSetUpPr.fitToPage = settings["fit_to_page"]
        if settings.get("fit_to_width") is not None:
            sheet.page_setup.fitToWidth = settings["fit_to_width"]
        if settings.get("fit_to_height") is not None:
            sheet.page_setup.fitToHeight = settings["fit_to_height"]
        if settings.get("clear_scale"):
            sheet.page_setup.scale = None
        if settings.get("paper_size") is not None:
            sheet.page_setup.paperSize = str(settings["paper_size"])
        if settings.get("print_area") is not None:
            sheet.print_area = settings["print_area"]
        if settings.get("horizontal_centered") is not None:
            sheet.print_options.horizontalCentered = settings["horizontal_centered"]
        if settings.get("vertical_centered") is not None:
            sheet.print_options.verticalCentered = settings["vertical_centered"]
        if settings.get("footer") is not None:
            sheet.oddFooter.center.text = settings["footer"]
    elif kind == "set_workbook_calculation_properties":
        calculation = workbook.calculation
        if calculation is None:
            from openpyxl.workbook.properties import CalcProperties
            calculation = CalcProperties()
            workbook.calculation = calculation
        if payload.get("full_calc_on_load") is not None:
            calculation.fullCalcOnLoad = payload["full_calc_on_load"]
        if payload.get("force_full_calc") is not None:
            calculation.forceFullCalc = payload["force_full_calc"]
    elif kind == "replace_header_text":
        sheet = workbook[payload["sheet"]]
        search, replacement = payload["search"], payload["replacement"]
        for collection in (sheet.oddHeader, sheet.evenHeader, sheet.firstHeader):
            for section in (collection.left, collection.center, collection.right):
                if section.text:
                    section.text = section.text.replace(search, replacement)
    else:
        raise WorkerError("unsupported_operation", "operation is not implemented: " + kind, unsupported=True)


def _verify_style_table_integrity(output_path: Path) -> None:
    """Reject a saved workbook whose cell style references exceed its style table.

    A styles table with fewer xf entries than a worksheet cell, row, or column
    style reference makes Excel/umya fall back to the default style and
    silently drop formatting semantics. Component-style assignment has been
    observed to produce such output on a real template; the worker must fail
    loudly instead of publishing a corrupted candidate.

    Parse the OOXML parts with namespace-aware parsers. Column dimensions refer
    to a cell style through their style attribute, not s.
    """
    import zipfile as _zipfile

    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    et_style_sheet_tag = "{{{}}}styleSheet".format(namespace)
    et_cell_xfs_tag = "{{{}}}cellXfs".format(namespace)
    et_xf_tag = "{{{}}}xf".format(namespace)
    expat_namespace = namespace + "}"
    expat_worksheet_tag = expat_namespace + "worksheet"
    expat_cell_tag = expat_namespace + "c"
    expat_row_tag = expat_namespace + "row"
    expat_column_tag = expat_namespace + "col"

    def checked_style_index(value: str, part_name: str, element_name: str) -> int:
        try:
            index = int(value)
        except ValueError:
            raise WorkerError(
                "style_table_corrupted",
                "saved workbook has an invalid {} style reference in {}".format(
                    element_name, part_name
                ),
            )
        if index < 0:
            raise WorkerError(
                "style_table_corrupted",
                "saved workbook has a negative {} style reference in {}".format(
                    element_name, part_name
                ),
            )
        return index

    try:
        with _zipfile.ZipFile(str(output_path)) as archive:
            styles_root = ElementTree.fromstring(archive.read("xl/styles.xml"))
            if styles_root.tag != et_style_sheet_tag:
                raise WorkerError(
                    "style_table_corrupted",
                    "saved workbook styles part has an unexpected root element",
                )
            cell_xfs_parts = [
                child for child in styles_root if child.tag == et_cell_xfs_tag
            ]
            if len(cell_xfs_parts) != 1:
                raise WorkerError(
                    "style_table_corrupted",
                    "saved workbook has no unique cellXfs style table",
                )
            cell_xfs = cell_xfs_parts[0]
            try:
                declared_count = int(cell_xfs.attrib["count"])
            except (KeyError, ValueError):
                raise WorkerError(
                    "style_table_corrupted",
                    "saved workbook cellXfs has no valid count attribute",
                )
            if declared_count < 0:
                raise WorkerError(
                    "style_table_corrupted",
                    "saved workbook cellXfs has a negative count attribute",
                )
            xf_elements = sum(child.tag == et_xf_tag for child in cell_xfs)
            worksheet_parts = [
                part_name
                for part_name in archive.namelist()
                if posixpath.dirname(part_name) == "xl/worksheets"
                and part_name.endswith(".xml")
            ]
            if not worksheet_parts:
                raise WorkerError(
                    "style_table_corrupted",
                    "saved workbook has no worksheet XML parts to validate",
                )
            for part_name in worksheet_parts:
                parser = expat.ParserCreate(namespace_separator="}")
                root_seen = False

                def check_worksheet_element(
                    element_name: str, attributes: Dict[str, str]
                ) -> None:
                    nonlocal root_seen
                    if not root_seen:
                        root_seen = True
                        if element_name != expat_worksheet_tag:
                            raise WorkerError(
                                "style_table_corrupted",
                                "saved workbook worksheet part has an unexpected root element: "
                                + part_name,
                            )
                    if element_name in (expat_cell_tag, expat_row_tag):
                        attribute_name = "s"
                    elif element_name == expat_column_tag:
                        attribute_name = "style"
                    else:
                        return
                    value = attributes.get(attribute_name)
                    if value is None:
                        return
                    style_index = checked_style_index(value, part_name, element_name)
                    if style_index >= xf_elements:
                        raise WorkerError(
                            "style_table_corrupted",
                            "saved workbook references cell style {} beyond its {} cellXfs entries in {}".format(
                                style_index, xf_elements, part_name
                            ),
                        )

                parser.StartElementHandler = check_worksheet_element
                with archive.open(part_name) as worksheet_xml:
                    parser.ParseFile(worksheet_xml)
                if not root_seen:
                    raise WorkerError(
                        "style_table_corrupted",
                        "saved workbook worksheet part is empty: " + part_name,
                    )
    except WorkerError:
        raise
    except Exception as error:
        raise WorkerError(
            "style_table_corrupted", "style-table integrity check failed: " + str(error)
        )
    if declared_count != xf_elements:
        raise WorkerError(
            "style_table_corrupted",
            "saved workbook cellXfs declares count={} but carries {} xf entries".format(
                declared_count, xf_elements
            ),
        )


def _materialize_explicit_style_attributes(path: Path) -> None:
    """Give every worksheet cell an explicit style attribute.

    OOXML treats an omitted ``s`` attribute as style index 0, but openpyxl
    omits ``s`` whenever the cell's style entry is the default while the Rust
    readback (umya) reports an EMPTY style instead of the default xf for such
    cells. Excel itself writes an explicit style attribute on every cell;
    mirror that canonical behaviour by adding ``s="0"`` wherever openpyxl
    omitted it. This keeps every downstream reader (umya readback included)
    consistent about omitted-style semantics.
    """
    replacement = path.with_name(path.name + ".styleattr.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "style-attribute temporary output already exists")
    pattern = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*?)(/?)>')

    def patch_part(data: bytes) -> bytes:
        text = data.decode("utf-8")

        def patch(match: "re.Match[str]") -> str:
            coordinate, attributes, closing = match.groups()
            if " s=" in attributes:
                return match.group(0)
            return '<c r="{}" s="0"{}{}>'.format(coordinate, attributes, closing)

        return pattern.sub(patch, text).encode("utf-8")

    try:
        changed_parts = set()
        with zipfile.ZipFile(str(path), "r") as source:
            for item in source.infolist():
                if not (
                    item.filename.startswith("xl/worksheets/")
                    and item.filename.endswith(".xml")
                ):
                    continue
                if patch_part(source.read(item.filename)) != source.read(item.filename):
                    changed_parts.add(item.filename)
        if not changed_parts:
            return
        with zipfile.ZipFile(str(path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename in changed_parts:
                    data = patch_part(data)
                target.writestr(item, data)
        os.replace(str(replacement), str(path))
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        if isinstance(error, WorkerError):
            raise
        raise WorkerError(
            "save_failed",
            "could not materialize merged member style attributes: " + str(error),
        )


def _worksheet_parts_by_title(package: zipfile.ZipFile) -> Dict[str, str]:
    """Map worksheet titles to their package part names (workbook rels order)."""
    document_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    workbook = ElementTree.fromstring(package.read("xl/workbook.xml"))
    relationships = ElementTree.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships.findall("{%s}Relationship" % package_rel_ns)
        if relationship.attrib.get("TargetMode") != "External"
    }
    mapping: Dict[str, str] = {}
    sheets = workbook.find("{%s}sheets" % main_ns)
    if sheets is None:
        return mapping
    for sheet in sheets.findall("{%s}sheet" % main_ns):
        name = sheet.attrib.get("name")
        target = targets.get(sheet.attrib.get("{%s}id" % document_rel_ns) or "")
        if not name or not target:
            continue
        part = posixpath.normpath(
            target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
        )
        if not part.startswith("../") and part.startswith("xl/"):
            mapping[name] = part
    return mapping


def _raw_numeric_cell_values(
    path: Path, coordinates: Iterable[Tuple[str, str]]
) -> Dict[Tuple[str, str], str]:
    """Read exact numeric OOXML values for cells formatted as dates.

    openpyxl converts a numeric cell with a date number format to a Python
    datetime when reopening a saved workbook. Comparing that object to the
    planned Excel serial falsely rejects valid date writes; converting the
    datetime back can also hide sub-millisecond changes. Compare the stored
    `<v>` text instead.
    """
    requested: Dict[str, Set[str]] = {}
    for sheet, address in coordinates:
        requested.setdefault(sheet, set()).add(address.replace("$", "").upper())
    result: Dict[Tuple[str, str], str] = {}
    if not requested:
        return result

    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(str(path), "r") as package:
        parts = _worksheet_parts_by_title(package)
        for sheet, addresses in requested.items():
            part = parts.get(sheet)
            if part is None or part not in package.namelist():
                continue
            root = ElementTree.fromstring(package.read(part))
            for cell in root.iter("{%s}c" % main_ns):
                address = cell.attrib.get("r", "").replace("$", "").upper()
                if address not in addresses:
                    continue
                if cell.attrib.get("t") not in (None, "n"):
                    continue
                value = cell.find("{%s}v" % main_ns)
                if value is not None and value.text is not None:
                    result[(sheet, address)] = value.text
    return result


def _ensure_cells_present(xml_text: str, wanted: Set[str]) -> str:
    """Insert `<c r="..." s="0"/>` elements for tracked coordinates that
    openpyxl's writer skipped (empty value + default style).

    Row and cell elements are inserted at their OOXML-sorted positions so the
    part stays schema-valid; existing elements are never modified.
    """
    from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple, get_column_letter

    if not wanted:
        return xml_text
    row_pattern = re.compile(
        r'<row\b[^>]*?\br="([0-9]+)"[^>]*?(?:/>|>(?:(?!</row>).)*?</row>)',
        re.DOTALL,
    )
    # 匹配完整单元格元素（自闭合或带子元素），保证切片与插入点落在元素边界。
    cell_pattern = re.compile(
        r'<c r="([A-Z]+)([0-9]+)"[^>]*?(?:/>|>(?:(?!</c>).)*?</c>)',
        re.DOTALL,
    )

    wanted_rows: Dict[int, Set[int]] = {}
    for coordinate in wanted:
        row, column = coordinate_to_tuple(coordinate)
        wanted_rows.setdefault(row, set()).add(column)

    def row_cell_text(row_index: int, column: int) -> str:
        return '<c r="{}{}" s="0"/>'.format(get_column_letter(column), row_index)

    def force_style_zero(cell_text: str) -> str:
        """给缺省 s 属性的受跟踪单元格补 s="0"（OOXML 语义等价，
        修复 umya 读取端把缺省 s 视为空样式的偏差）。"""
        open_end = cell_text.find(">")
        if open_end < 0:
            return cell_text
        open_tag = cell_text[:open_end]
        if " s=" in open_tag:
            return cell_text
        return open_tag + ' s="0"' + cell_text[open_end:]

    def replace_row(match: "re.Match[str]") -> str:
        row_text = match.group(0)
        row_index = int(match.group(1))
        columns = wanted_rows.get(row_index)
        if not columns:
            return row_text
        coordinates = {"{}{}".format(name, digits) for name, digits in cell_pattern.findall(row_text)}
        present = {
            column_index_from_string(name)
            for name, digits in cell_pattern.findall(row_text)
            if "{}{}".format(name, digits) in wanted
        }
        missing = sorted(columns - present)
        if not missing and not (coordinates & wanted):
            return row_text
        if row_text.endswith("/>"):
            head = row_text[:-2]
            insertion = "".join(row_cell_text(row_index, column) for column in missing)
            return head + ">" + insertion + "</row>"
        open_tag_end = row_text.index(">") + 1
        content = row_text[open_tag_end:-len("</row>")]
        existing = [
            (match.start(), match.end(), match.group(1), match.group(2))
            for match in cell_pattern.finditer(content)
        ]
        pieces: List[str] = []
        cursor = 0
        missing_index = 0
        for start, end, name, digits in existing:
            while missing_index < len(missing) and missing[missing_index] < column_index_from_string(name):
                pieces.append(content[cursor:start])
                pieces.append(row_cell_text(row_index, missing[missing_index]))
                missing_index += 1
                cursor = start
            piece = content[cursor:end]
            cursor = end
            if "{}{}".format(name, digits) in wanted:
                piece = force_style_zero(piece)
            pieces.append(piece)
        while missing_index < len(missing):
            pieces.append(row_cell_text(row_index, missing[missing_index]))
            missing_index += 1
        pieces.append(content[cursor:])
        return row_text[:open_tag_end] + "".join(pieces) + "</row>"

    sheet_data_match = re.search(
        r"<sheetData\b[^>]*?/>|<sheetData\b[^>]*?>(.*)</sheetData>",
        xml_text,
        re.DOTALL,
    )
    if sheet_data_match is None:
        return xml_text
    empty_sheet_data = sheet_data_match.group(0).endswith("/>")
    body = "" if empty_sheet_data else (sheet_data_match.group(1) or "")

    body = row_pattern.sub(replace_row, body)
    present_rows = {int(match.group(1)) for match in row_pattern.finditer(body)}
    pending = sorted(row for row in wanted_rows if row not in present_rows)
    if pending:
        segments = [(match.start(), match.end(), int(match.group(1)))
                    for match in row_pattern.finditer(body)]
        pieces: List[str] = []
        cursor = 0
        pending_index = 0
        for start, end, row_index in segments:
            while pending_index < len(pending) and pending[pending_index] < row_index:
                row = pending[pending_index]
                pieces.append(body[cursor:start])
                pieces.append(
                    "<row r=\"{}\">{}</row>".format(
                        row,
                        "".join(row_cell_text(row, column) for column in sorted(wanted_rows[row])),
                    )
                )
                pending_index += 1
                cursor = start
            pieces.append(body[cursor:end])
            cursor = end
        while pending_index < len(pending):
            row = pending[pending_index]
            pieces.append(
                "<row r=\"{}\">{}</row>".format(
                    row,
                    "".join(row_cell_text(row, column) for column in sorted(wanted_rows[row])),
                )
            )
            pending_index += 1
        pieces.append(body[cursor:])
        body = "".join(pieces)

    start, end = sheet_data_match.span()
    if empty_sheet_data:
        return xml_text[:start] + "<sheetData>" + body + "</sheetData>" + xml_text[end:]
    return xml_text[:start] + "<sheetData>" + body + "</sheetData>" + xml_text[end:]


def _materialize_tracked_cells(path: Path, tracked: Dict[str, Set[str]]) -> None:
    """Write explicit (empty, default-styled) cells for coordinates that the
    plan materialized but openpyxl's writer skipped.

    openpyxl skips cells with no value and the default style; the plan's
    materialize semantics require them to exist so the Rust readback observes
    the same cell inventory as the source."""
    if not tracked:
        return
    replacement = path.with_name(path.name + ".cells.tmp")
    if replacement.exists():
        raise WorkerError("save_failed", "materialized-cells temporary output already exists")
    try:
        changed_parts = set()
        with zipfile.ZipFile(str(path), "r") as source:
            title_to_part = _worksheet_parts_by_title(source)
            wanted_by_part: Dict[str, Set[str]] = {}
            for title, coordinates in tracked.items():
                part = title_to_part.get(title)
                if not part or part not in source.namelist():
                    continue
                wanted_by_part.setdefault(part, set()).update(coordinates)
            for item in source.infolist():
                wanted = wanted_by_part.get(item.filename)
                if not wanted:
                    continue
                if _ensure_cells_present(
                    source.read(item.filename).decode("utf-8"), wanted
                ) != source.read(item.filename).decode("utf-8"):
                    changed_parts.add(item.filename)
        if not changed_parts:
            return
        with zipfile.ZipFile(str(path), "r") as source, zipfile.ZipFile(
            str(replacement), "w"
        ) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename in changed_parts:
                    data = _ensure_cells_present(
                        data.decode("utf-8"),
                        wanted_by_part[item.filename],
                    ).encode("utf-8")
                target.writestr(item, data)
        os.replace(str(replacement), str(path))
    except WorkerError:
        raise
    except Exception as error:
        try:
            replacement.unlink()
        except OSError:
            pass
        raise WorkerError(
            "save_failed", "could not materialize tracked cells: " + str(error)
        )


def _apply(request: Dict[str, Any]) -> Dict[str, Any]:
    started = time.monotonic()
    input_path, output_path, operations = _preflight(request)
    before_hash = _sha256(input_path)
    try:
        import openpyxl
    except Exception as error:
        raise WorkerError("worker_unavailable", "openpyxl import failed: " + str(error), unsupported=True)
    try:
        workbook = openpyxl.load_workbook(str(input_path), data_only=False, keep_links=True, read_only=False)
        # data_only 视图（既有公式的缓存值）只被 copy_range（同簿复制取缓存）
        # 使用；纯跨簿导入的计划跳过这一次整本加载，省去对大目标簿的一次
        # 完整读取。语义不变：没有 copy_range 时该视图本来就不会被读取。
        cached_workbook = (
            openpyxl.load_workbook(str(input_path), data_only=True, keep_links=True, read_only=False)
            if any(operation.get("operation") == "copy_range" for operation in operations)
            else None
        )
    except Exception as error:
        raise WorkerError("workbook_open_failed", str(error))
    try:
        initial_sheet_names = list(workbook.sheetnames)
        if _formula_cache_values(input_path) and any(
            operation.get("operation")
            in ("set_cell_value", "set_formula_text", "delete_rows", "insert_rows", "copy_range", "import_range")
            for operation in operations
        ):
            # The exact existing caches are restored after save. Also ask
            # Excel/WPS to refresh formulas whose inputs may have changed.
            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True
        materialized_by_sheet: Dict[int, Set[str]] = {}
        for operation in operations:
            try:
                _apply_operation(workbook, cached_workbook, operation, materialized_by_sheet)
            except WorkerError:
                raise
            except Exception as error:
                raise WorkerError("operation_failed", "{}: {}".format(operation.get("operation"), error))
        if not workbook.worksheets:
            raise WorkerError("invalid_output", "workbook must retain at least one worksheet")
        # 就地修改计划：把 id 键控的具体化坐标转换为最终工作表标题键控
        # （重命名/删除操作之后仍指向最终输出中的正确工作表）。
        tracked_cells: Dict[str, Set[str]] = {}
        for worksheet in workbook.worksheets:
            coordinates = materialized_by_sheet.get(id(worksheet))
            if coordinates:
                tracked_cells[worksheet.title] = coordinates
        try:
            workbook.save(str(output_path))
        except Exception as error:
            raise WorkerError("save_failed", str(error))
    finally:
        workbook.close()
        if cached_workbook is not None:
            cached_workbook.close()
    _preserve_significant_string_whitespace(output_path)
    _materialize_tracked_cells(output_path, tracked_cells)
    _materialize_explicit_style_attributes(output_path)
    _preserve_formula_caches(input_path, output_path, operations)
    _restore_cell_metadata_package_parts(input_path, output_path, operations)
    _restore_unmodified_column_dimensions(input_path, output_path, operations)
    _restore_header_footer_settings(input_path, output_path, operations)
    _restore_calculation_chain(input_path, output_path, operations)
    _restore_planned_empty_strings(output_path, operations)
    after_hash = _sha256(input_path)
    if after_hash != before_hash:
        raise WorkerError("source_changed", "worker input changed while applying the edit plan")
    if not output_path.is_file() or output_path.stat().st_size == 0:
        raise WorkerError("save_failed", "openpyxl did not create a non-empty output workbook")
    _verify_style_table_integrity(output_path)
    try:
        verification = openpyxl.load_workbook(str(output_path), read_only=False, data_only=False, keep_links=True)
        if not verification.worksheets:
            raise WorkerError("output_reopen_failed", "saved workbook has no worksheets")
        _verify_plan(verification, operations, initial_sheet_names, output_path)
        verification.close()
    except WorkerError:
        raise
    except Exception as error:
        raise WorkerError("output_reopen_failed", str(error))
    return {
        "schema_version": SCHEMA_VERSION,
        "request_id": request["request_id"],
        "status": "applied",
        "operation_count": len(operations),
        "verified_operation_count": len(operations),
        "input_sha256_before": before_hash,
        "input_sha256_after": after_hash,
        "output_sha256": _sha256(output_path),
        "diagnostics": {
            "python_version": sys.version.split()[0],
            "openpyxl_version": openpyxl.__version__,
            "elapsed_ms": int((time.monotonic() - started) * 1000),
        },
    }


def _final_sheet_names(initial_names: List[str], operations: List[Dict[str, Any]]) -> List[str]:
    """Apply sheet-name operations to the original order for final-state checks."""
    expected = list(initial_names)
    for operation in operations:
        kind = operation["operation"]
        if kind == "rename_sheet":
            old_name, new_name = operation["old_name"], operation["new_name"]
            try:
                index = expected.index(old_name)
            except ValueError:
                raise WorkerError(
                    "output_verification_failed",
                    "worksheet rename source missing from planned state",
                )
            expected[index] = new_name
        elif kind == "create_sheet":
            expected.append(operation["name"])
        elif kind == "create_sheet_at":
            index = operation["index"]
            if index < 0 or index > len(expected):
                raise WorkerError(
                    "output_verification_failed",
                    "worksheet insertion index is invalid in planned state",
                )
            expected.insert(index, operation["name"])
        elif kind == "delete_sheet":
            name = operation["name"]
            try:
                expected.remove(name)
            except ValueError:
                raise WorkerError(
                    "output_verification_failed",
                    "worksheet deletion target missing from planned state",
                )
    return expected


def _verify_import_dimensions(workbook: Any, operation: Dict[str, Any]) -> None:
    """Verify the row/column dimensions copied by an ImportRange operation."""
    if not any(
        operation.get(key)
        for key in (
            "copy_row_heights",
            "copy_row_hidden",
            "copy_column_widths",
            "copy_column_hidden",
        )
    ):
        return

    import openpyxl
    from openpyxl.utils.cell import (
        column_index_from_string,
        coordinate_to_tuple,
        get_column_letter,
        range_boundaries,
    )

    source_book = openpyxl.load_workbook(
        str(operation["source_workbook"]),
        data_only=False,
        keep_links=True,
        read_only=False,
    )
    try:
        source = source_book[operation["source_sheet"]]
        target = workbook[operation["target_sheet"]]
        min_col, min_row, max_col, max_row = range_boundaries(operation["source_range"])
        target_row, target_col = coordinate_to_tuple(
            operation["target_start"].replace("$", "")
        )
        row_offset = target_row - min_row
        col_offset = target_col - min_col

        row_fields = []
        if operation.get("copy_row_heights"):
            row_fields.append("height")
        if operation.get("copy_row_hidden"):
            row_fields.append("hidden")
        if operation.get("copy_outline_levels") and row_fields:
            row_fields.append("outlineLevel")
        if row_fields:
            for source_index, source_dimension in source.row_dimensions.items():
                if not operation.get("copy_all_row_dimensions") and not min_row <= source_index <= max_row:
                    continue
                target_index = source_index + row_offset
                target_dimension = target.row_dimensions[target_index]
                for field in row_fields:
                    expected = getattr(source_dimension, field)
                    actual = getattr(target_dimension, field)
                    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
                        matches = abs(float(expected) - float(actual)) <= 1e-10
                    else:
                        matches = expected == actual
                    if not matches:
                        raise WorkerError(
                            "output_verification_failed",
                            "ImportRange row dimension mismatch at row {} ({})".format(
                                target_index, field
                            ),
                        )

        column_fields = []
        if operation.get("copy_column_widths"):
            column_fields.append("width")
        if operation.get("copy_column_hidden"):
            column_fields.append("hidden")
        if operation.get("copy_outline_levels") and column_fields:
            column_fields.append("outlineLevel")
        if column_fields:
            for key, source_dimension in source.column_dimensions.items():
                first = source_dimension.min or column_index_from_string(key)
                last = source_dimension.max or first
                if not operation.get("copy_all_column_dimensions"):
                    if last < min_col or first > max_col:
                        continue
                    first = max(first, min_col)
                    last = min(last, max_col)
                for source_index in range(first, last + 1):
                    target_index = source_index + col_offset
                    target_dimension = target.column_dimensions[
                        get_column_letter(target_index)
                    ]
                    for field in column_fields:
                        expected = getattr(source_dimension, field)
                        actual = getattr(target_dimension, field)
                        if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
                            matches = abs(float(expected) - float(actual)) <= 1e-10
                        else:
                            matches = expected == actual
                        if not matches:
                            raise WorkerError(
                                "output_verification_failed",
                                "ImportRange column dimension mismatch at column {} ({})".format(
                                    get_column_letter(target_index), field
                                ),
                            )
    finally:
        source_book.close()


def _verify_plan(
    workbook: Any,
    operations: List[Dict[str, Any]],
    initial_sheet_names: List[str],
    output_path: Path,
) -> None:
    """Read back simple edits with openpyxl; Rust separately reopens the file."""
    from decimal import Decimal
    from openpyxl.utils.cell import get_column_letter

    date_time_cells = [
        (operation["sheet"], operation["cell"])
        for operation in operations
        if operation.get("operation") == "set_cell_value"
        and operation.get("value", {}).get("type") == "date_time"
    ]
    raw_date_time_values = _raw_numeric_cell_values(output_path, date_time_cells)

    expected_sheet_names = _final_sheet_names(initial_sheet_names, operations)
    if workbook.sheetnames != expected_sheet_names:
        raise WorkerError(
            "output_verification_failed",
            "worksheet names/order readback mismatch",
        )

    for operation in operations:
        kind = operation["operation"]
        payload = operation
        if kind == "set_cell_value":
            cell = workbook[payload["sheet"]][payload["cell"]]
            expected = _cell_value(payload["value"])
            if isinstance(expected, Decimal):
                try:
                    # Match openpyxl's own XLSX numeric serialization
                    # (compat.strings.safe_string uses %.16g for numbers).
                    from openpyxl.compat.strings import safe_string

                    serialized_expected = float(safe_string(expected))
                    if payload["value"].get("type") == "date_time":
                        raw_value = raw_date_time_values.get(
                            (payload["sheet"], payload["cell"].replace("$", "").upper())
                        )
                        numeric_matches = (
                            raw_value is not None
                            and float(raw_value) == serialized_expected
                        )
                    else:
                        numeric_matches = (
                            cell.value is not None
                            and float(cell.value) == serialized_expected
                        )
                except (TypeError, ValueError, OverflowError):
                    numeric_matches = False
                if not numeric_matches:
                    raise WorkerError("output_verification_failed", "numeric value readback mismatch at {}!{}".format(payload["sheet"], payload["cell"]))
            elif (
                expected == ""
                and cell.value is None
                and cell.data_type == "inlineStr"
            ):
                # Openpyxl reads explicit empty inline strings as None, but
                # _restore_planned_empty_strings has retained their OOXML type.
                pass
            elif cell.value != expected:
                raise WorkerError("output_verification_failed", "value readback mismatch at {}!{}".format(payload["sheet"], payload["cell"]))
            if payload["value"].get("type") == "date_time" and cell.number_format != payload["value"]["value"]["number_format"]:
                raise WorkerError("output_verification_failed", "date number format readback mismatch")
        elif kind == "set_formula_text":
            formula = payload["formula"]
            formula = formula if formula.startswith("=") else "=" + formula
            if workbook[payload["sheet"]][payload["cell"]].value != formula:
                raise WorkerError("output_verification_failed", "formula text readback mismatch")
        elif kind == "set_comment":
            actual = workbook[payload["sheet"]][payload["cell"]].comment
            expected = payload.get("text")
            if (actual.text if actual else None) != expected:
                raise WorkerError("output_verification_failed", "comment text readback mismatch")
            if actual and actual.author != (payload.get("author") or "pytools"):
                raise WorkerError("output_verification_failed", "comment author readback mismatch")
        elif kind == "set_column_width":
            actual = workbook[payload["sheet"]].column_dimensions[get_column_letter(payload["column_index"])].width
            if actual is None or abs(float(actual) - float(payload["width"])) > 1e-10:
                raise WorkerError("output_verification_failed", "column width readback mismatch")
        elif kind == "set_row_height":
            actual = workbook[payload["sheet"]].row_dimensions[payload["row_index"]].height
            if actual is None or abs(float(actual) - float(payload["height"])) > 1e-10:
                raise WorkerError("output_verification_failed", "row height readback mismatch")
        elif kind == "set_freeze_panes":
            actual = workbook[payload["sheet"]].freeze_panes
            actual = None if actual is None else str(actual)
            if actual != payload.get("cell"):
                raise WorkerError("output_verification_failed", "freeze panes readback mismatch")
        elif kind == "set_auto_filter":
            actual = workbook[payload["sheet"]].auto_filter.ref
            if actual != payload.get("range"):
                raise WorkerError("output_verification_failed", "AutoFilter readback mismatch")
        elif kind == "set_cell_fill_color":
            expected = payload["argb"].strip().lstrip("#").upper()
            if len(expected) == 6:
                expected = "FF" + expected
            actual = workbook[payload["sheet"]][payload["cell"]].fill.fgColor.rgb
            if not actual or str(actual).upper() != expected:
                raise WorkerError("output_verification_failed", "cell fill readback mismatch")
        elif kind == "set_page_settings":
            sheet = workbook[payload["sheet"]]
            settings = payload["settings"]
            checks = (
                ("orientation", sheet.page_setup.orientation),
                ("fit_to_width", sheet.page_setup.fitToWidth),
                ("fit_to_height", sheet.page_setup.fitToHeight),
                ("paper_size", sheet.page_setup.paperSize),
                ("horizontal_centered", sheet.print_options.horizontalCentered),
                ("vertical_centered", sheet.print_options.verticalCentered),
            )
            for key, actual in checks:
                expected = settings.get(key)
                if expected is not None and str(actual).lower() != str(expected).lower():
                    raise WorkerError("output_verification_failed", "page setting readback mismatch: " + key)
            if settings.get("fit_to_page") is not None and sheet.sheet_properties.pageSetUpPr.fitToPage != settings["fit_to_page"]:
                raise WorkerError("output_verification_failed", "fit-to-page readback mismatch")
            if settings.get("clear_scale") and sheet.page_setup.scale is not None:
                raise WorkerError("output_verification_failed", "page scale was not cleared")
            if settings.get("print_area") is not None:
                actual = str(sheet.print_area)
                if settings["print_area"].replace("$", "") not in actual.replace("$", ""):
                    raise WorkerError("output_verification_failed", "print area readback mismatch")
            if settings.get("footer") is not None and sheet.oddFooter.center.text != settings["footer"]:
                raise WorkerError("output_verification_failed", "footer readback mismatch")
        elif kind == "set_workbook_calculation_properties":
            calculation = workbook.calculation
            if payload.get("full_calc_on_load") is not None and calculation.fullCalcOnLoad != payload["full_calc_on_load"]:
                raise WorkerError("output_verification_failed", "fullCalcOnLoad readback mismatch")
            if payload.get("force_full_calc") is not None and calculation.forceFullCalc != payload["force_full_calc"]:
                raise WorkerError("output_verification_failed", "forceFullCalc readback mismatch")
        elif kind in ("copy_range", "import_range", "replace_header_text"):
            # These operations receive independent range/metadata checks in the
            # Rust validator and dedicated fixture tests; opening the saved
            # package here still verifies it can be deserialized.
            if kind == "import_range":
                _verify_import_dimensions(workbook, payload)
            continue


def _response(request_id: Optional[str], status: str, kind: str, message: str,
              diagnostics: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "request_id": request_id,
        "status": status,
        "error": {"kind": kind, "message": message},
        "diagnostics": diagnostics or {},
    }


def _run(request_path: Path, response_path: Path) -> int:
    request_id = None
    try:
        request = json.loads(request_path.read_text(encoding="utf-8"))
        if isinstance(request, dict):
            request_id = request.get("request_id")
        result = _apply(request)
        response_path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        return 0
    except WorkerError as error:
        status = "unsupported_before_write" if error.unsupported else "failed"
        result = _response(request_id, status, error.kind, str(error), {"traceback": traceback.format_exc()})
        response_path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        return 0 if error.unsupported else 1
    except Exception as error:
        result = _response(request_id, "failed", "internal_error", str(error), {"traceback": traceback.format_exc()})
        response_path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        return 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--request", required=False)
    parser.add_argument("--response", required=False)
    parser.add_argument(
        "--version",
        action="store_true",
        help="print worker protocol version and exit",
    )
    args = parser.parse_args()
    if args.version:
        print(f"pytools-openpyxl-worker protocol={WORKER_PROTOCOL_VERSION}")
        try:
            import openpyxl
        except Exception as error:  # noqa: BLE001 - 环境诊断必须捕获所有导入失败
            print(f"openpyxl=missing ({error})")
            # 退出码 3：运行环境缺少 openpyxl（Rust 侧据此区分环境问题与协议问题）。
            return 3
        print(f"openpyxl={openpyxl.__version__}")
        return 0
    if not args.request or not args.response:
        parser.error("--request and --response are required")
    return _run(Path(args.request), Path(args.response))


if __name__ == "__main__":
    raise SystemExit(main())
