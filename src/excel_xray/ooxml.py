"""Structural read straight from the OOXML package.

openpyxl in read-only mode silently drops merged ranges, ListObjects, external
links, connections and pivot caches -- exactly the parts that tell you what a
workbook is *for*. This module reads them from the zip directly: complete, and
roughly an order of magnitude faster than a full openpyxl load.

Cell values are NOT read here -- that is sheetmodel.py's single lxml pass.
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass, field

from lxml import etree

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}
R_ID = f"{{{NS['r']}}}id"

# Strict-namespace variants exist in the wild (files written by some non-Excel
# tools). Detect and remap rather than returning an empty workbook.
STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main"


@dataclass
class SheetRef:
    name: str
    sheet_id: str
    state: str          # visible | hidden | veryHidden
    rel_id: str
    part: str           # zip path, e.g. xl/worksheets/sheet1.xml
    position: int


@dataclass
class TableDef:
    name: str
    display_name: str
    ref: str
    header_row_count: int
    totals_row_count: int
    columns: list[str]


@dataclass
class SheetStructure:
    dimension: str | None = None
    merges: list[str] = field(default_factory=list)
    tables: list[TableDef] = field(default_factory=list)
    has_autofilter: bool = False
    conditional_formatting_ranges: int = 0
    data_validation_ranges: int = 0
    hyperlink_targets: list[str] = field(default_factory=list)


@dataclass
class WorkbookStructure:
    path: str
    sheets: list[SheetRef] = field(default_factory=list)
    structures: dict[str, SheetStructure] = field(default_factory=dict)
    defined_names: list[tuple[str, str]] = field(default_factory=list)
    external_links: list[str] = field(default_factory=list)
    connections: list[dict] = field(default_factory=list)
    pivot_cache_sources: list[str] = field(default_factory=list)
    has_vba: bool = False
    has_power_query: bool = False
    core_props: dict = field(default_factory=dict)
    app_props: dict = field(default_factory=dict)
    parse_notes: list[str] = field(default_factory=list)


def _txt(el) -> str | None:
    return el.text if el is not None else None


def _parse(zf: zipfile.ZipFile, name: str):
    try:
        with zf.open(name) as fh:
            return etree.parse(fh).getroot()
    except KeyError:
        return None
    except etree.XMLSyntaxError as e:
        raise ValueError(f"malformed XML in {name}: {e}") from e


def _rels_for(zf: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Relationship id -> resolved zip path, for the given part."""
    d, base = posixpath.split(part)
    rels_path = posixpath.join(d, "_rels", base + ".rels")
    root = _parse(zf, rels_path)
    if root is None:
        return {}
    out = {}
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target", "")
        mode = rel.get("TargetMode", "")
        if mode == "External":
            out[rid] = target
        else:
            out[rid] = posixpath.normpath(posixpath.join(d, target)).lstrip("/")
    return out


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def read_structure(path: str) -> WorkbookStructure:
    """Read every structural part of one workbook. Never raises on content."""
    ws = WorkbookStructure(path=path)
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        ws.has_vba = "xl/vbaProject.bin" in names
        ws.has_power_query = any(
            n.startswith("customXml/item") for n in names
        ) and _detect_mashup(zf, names)

        _read_props(zf, ws, names)

        wb_root = _parse(zf, "xl/workbook.xml")
        if wb_root is None:
            ws.parse_notes.append("no xl/workbook.xml -- not a valid xlsx package")
            return ws
        ns = NS
        if wb_root.nsmap.get(None) == STRICT_MAIN:
            ns = dict(NS, m=STRICT_MAIN)
            ws.parse_notes.append("strict OOXML namespace -- remapped")

        wb_rels = _rels_for(zf, "xl/workbook.xml")

        for i, sh in enumerate(wb_root.findall(".//m:sheets/m:sheet", ns)):
            rid = sh.get(R_ID)
            part = wb_rels.get(rid, "")
            ws.sheets.append(SheetRef(
                name=sh.get("name", f"Sheet{i+1}"),
                sheet_id=sh.get("sheetId", str(i + 1)),
                state=sh.get("state", "visible"),
                rel_id=rid or "",
                part=part,
                position=i,
            ))

        for dn in wb_root.findall(".//m:definedNames/m:definedName", ns):
            ws.defined_names.append((dn.get("name", ""), (dn.text or "").strip()))

        for rid, target in wb_rels.items():
            if "externalLink" in target:
                ext_rels = _rels_for(zf, target)
                for t in ext_rels.values():
                    if t.startswith(("http", "file:", "/", "\\")) or ":" in t[:3]:
                        ws.external_links.append(t)

        _read_connections(zf, ws, ns)
        _read_pivot_caches(zf, ws, names, ns)

        for sheet in ws.sheets:
            if not sheet.part or sheet.part not in names:
                ws.parse_notes.append(f"sheet part missing for {sheet.name!r}")
                continue
            ws.structures[sheet.name] = _read_sheet_structure(zf, sheet.part, ns)

    return ws


def _detect_mashup(zf: zipfile.ZipFile, names: set[str]) -> bool:
    """Power Query stores its M code in a base64 DataMashup part."""
    for n in names:
        if n.startswith("customXml/item") and n.endswith(".xml"):
            try:
                head = zf.read(n)[:400]
            except Exception:
                continue
            if b"DataMashup" in head:
                return True
    return False


def _read_props(zf: zipfile.ZipFile, ws: WorkbookStructure, names: set[str]) -> None:
    core = _parse(zf, "docProps/core.xml") if "docProps/core.xml" in names else None
    if core is not None:
        ws.core_props = {
            "creator": _txt(core.find("dc:creator", NS)),
            "last_modified_by": _txt(core.find("cp:lastModifiedBy", NS)),
            "created": _txt(core.find("dcterms:created", NS)),
            "modified": _txt(core.find("dcterms:modified", NS)),
            "title": _txt(core.find("dc:title", NS)),
        }
    app = _parse(zf, "docProps/app.xml") if "docProps/app.xml" in names else None
    if app is not None:
        ws.app_props = {
            "application": _txt(app.find("ep:Application", NS)),
            "company": _txt(app.find("ep:Company", NS)),
        }


def _read_connections(zf: zipfile.ZipFile, ws: WorkbookStructure, ns: dict) -> None:
    root = _parse(zf, "xl/connections.xml")
    if root is None:
        return
    for conn in root:
        db = None
        for child in conn:
            if _local(child.tag) == "dbPr":
                db = child
        ws.connections.append({
            "name": conn.get("name"),
            "type": conn.get("type"),
            "description": conn.get("description"),
            "connection_string": (db.get("connection") if db is not None else None),
            "command": (db.get("command") if db is not None else None),
        })


def _read_pivot_caches(
    zf: zipfile.ZipFile, ws: WorkbookStructure, names: set[str], ns: dict
) -> None:
    for n in sorted(names):
        if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", n):
            root = _parse(zf, n)
            if root is None:
                continue
            for src in root.iter():
                lt = _local(src.tag)
                if lt == "worksheetSource":
                    ref = src.get("ref") or src.get("name") or ""
                    sheet = src.get("sheet") or ""
                    ws.pivot_cache_sources.append(
                        f"{sheet}!{ref}" if sheet else ref)
                elif lt == "dbPr" or lt == "connection":
                    c = src.get("connection")
                    if c:
                        ws.pivot_cache_sources.append(c)


def _read_sheet_structure(
    zf: zipfile.ZipFile, part: str, ns: dict
) -> SheetStructure:
    """Stream one sheet's XML for structure only, stopping before sheetData.

    We use iterparse so a 100k-row sheet does not get fully materialised just to
    learn it has three merged ranges.
    """
    st = SheetStructure()
    table_rids: list[str] = []
    try:
        with zf.open(part) as fh:
            for event, el in etree.iterparse(fh, events=("end",)):
                lt = _local(el.tag)
                if lt == "dimension":
                    st.dimension = el.get("ref")
                elif lt == "mergeCell":
                    ref = el.get("ref")
                    if ref:
                        st.merges.append(ref)
                elif lt == "autoFilter":
                    st.has_autofilter = True
                elif lt == "conditionalFormatting":
                    st.conditional_formatting_ranges += 1
                elif lt == "dataValidation":
                    st.data_validation_ranges += 1
                elif lt == "tablePart":
                    rid = el.get(R_ID)
                    if rid:
                        table_rids.append(rid)
                elif lt == "hyperlink":
                    t = el.get("location") or el.get(R_ID)
                    if t:
                        st.hyperlink_targets.append(t)
                # Free memory aggressively; sheetData rows are the bulk.
                if lt in ("row", "mergeCell", "conditionalFormatting",
                          "dataValidation", "hyperlink"):
                    el.clear()
    except etree.XMLSyntaxError as e:
        st.dimension = st.dimension or None
        return st

    if table_rids:
        rels = _rels_for(zf, part)
        for rid in table_rids:
            tpath = rels.get(rid)
            if not tpath:
                continue
            troot = _parse(zf, tpath)
            if troot is None:
                continue
            cols = [c.get("name", "") for c in troot.iter()
                    if _local(c.tag) == "tableColumn"]
            st.tables.append(TableDef(
                name=troot.get("name", ""),
                display_name=troot.get("displayName", ""),
                ref=troot.get("ref", ""),
                header_row_count=int(troot.get("headerRowCount", "1") or 0),
                totals_row_count=int(troot.get("totalsRowCount", "0") or 0),
                columns=cols,
            ))
    return st
