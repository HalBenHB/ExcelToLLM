from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from openpyxl.utils import get_column_letter

from .constants import XML_NS


def _normalize_zip_target(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    normalized = posixpath.normpath(posixpath.join(base_dir, target))
    return normalized.lstrip("/")


def _load_relationships(zip_file: ZipFile, rels_path: str) -> dict[str, tuple[str, str]]:
    if rels_path not in zip_file.namelist():
        return {}

    rels_root = ET.fromstring(zip_file.read(rels_path))
    relationships: dict[str, tuple[str, str]] = {}
    for rel in rels_root.findall("pkgrel:Relationship", XML_NS):
        relationships[rel.attrib["Id"]] = (
            rel.attrib.get("Type", ""),
            rel.attrib.get("Target", ""),
        )
    return relationships


def _extract_shape_text(shape) -> str:
    paragraphs: list[str] = []
    for paragraph in shape.findall("xdr:txBody/a:p", XML_NS):
        fragments = [
            (node.text or "")
            for node in paragraph.findall(".//a:t", XML_NS)
            if (node.text or "").strip()
        ]
        if fragments:
            paragraphs.append("".join(fragments).strip())
    return "\n".join(paragraphs).strip()


def _anchor_to_label(anchor) -> str:
    if anchor is None:
        return "unknown position"

    col_node = anchor.find("xdr:col", XML_NS)
    row_node = anchor.find("xdr:row", XML_NS)
    if col_node is None or row_node is None:
        return "unknown position"

    try:
        col_idx = int(col_node.text or "0")
        row_idx = int(row_node.text or "0")
    except ValueError:
        return "unknown position"

    return f"{get_column_letter(col_idx + 1)}{row_idx + 1}"


def extract_sheet_floating_text(
    file_path: Path,
    sheet_names: list[str],
) -> dict[str, list[dict[str, str]]]:
    results = {sheet_name: [] for sheet_name in sheet_names}

    with ZipFile(file_path) as zip_file:
        workbook_root = ET.fromstring(zip_file.read("xl/workbook.xml"))
        workbook_rels = _load_relationships(zip_file, "xl/_rels/workbook.xml.rels")

        sheet_part_by_name: dict[str, str] = {}
        for sheet in workbook_root.find("main:sheets", XML_NS):
            sheet_name = sheet.attrib.get("name")
            rel_id = sheet.attrib.get(f"{{{XML_NS['office_rel']}}}id")
            if not sheet_name or not rel_id or rel_id not in workbook_rels:
                continue

            _, target = workbook_rels[rel_id]
            sheet_part_by_name[sheet_name] = _normalize_zip_target("xl/workbook.xml", target)

        for sheet_name, sheet_part in sheet_part_by_name.items():
            if sheet_name not in results:
                continue

            sheet_filename = posixpath.basename(sheet_part)
            sheet_rels_path = f"xl/worksheets/_rels/{sheet_filename}.rels"
            sheet_rels = _load_relationships(zip_file, sheet_rels_path)

            drawing_targets = [
                target for rel_type, target in sheet_rels.values()
                if rel_type.endswith("/drawing")
            ]

            for drawing_target in drawing_targets:
                drawing_part = _normalize_zip_target(sheet_part, drawing_target)
                if drawing_part not in zip_file.namelist():
                    continue

                drawing_root = ET.fromstring(zip_file.read(drawing_part))
                for anchor_tag in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                    for anchor in drawing_root.findall(f"xdr:{anchor_tag}", XML_NS):
                        for shape in anchor.findall("xdr:sp", XML_NS):
                            text = _extract_shape_text(shape)
                            if not text:
                                continue

                            c_nv_pr = shape.find("xdr:nvSpPr/xdr:cNvPr", XML_NS)
                            shape_name = (
                                c_nv_pr.attrib.get("name", "Unnamed shape")
                                if c_nv_pr is not None else "Unnamed shape"
                            )
                            from_anchor = anchor.find("xdr:from", XML_NS)
                            results[sheet_name].append(
                                {
                                    "name": shape_name,
                                    "anchor": _anchor_to_label(from_anchor),
                                    "text": text,
                                }
                            )

    return results
