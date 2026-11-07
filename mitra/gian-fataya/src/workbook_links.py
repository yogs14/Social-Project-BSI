"""Native in-workbook navigation, including repairs that preserve cell caches."""
from __future__ import annotations

import hashlib
from pathlib import Path
import posixpath
import re
from urllib.parse import unquote, urlsplit
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET
from openpyxl.utils.cell import quote_sheetname
from openpyxl.worksheet.hyperlink import Hyperlink

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"s": S}
CELL_RANGE = re.compile(r"\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$")
SELF_FILES = {"POS_Gian_Fataya.xlsx", "POS_Gian_Fataya_DUMMY.xlsx", "SIMULASI_Grafik_Wilayah.xlsx"}


def xml(data):
    return ET.fromstring(data, parser=ET.XMLParser(resolve_entities=False, no_network=True))


def internal_link(cell, sheet, anchor="A1"):
    # Assigning a string beginning with # creates an External relationship in
    # openpyxl. A location-only Hyperlink is the SpreadsheetML bookmark form.
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"{quote_sheetname(sheet)}!{anchor}", display=str(cell.value or ""))


def sheet_parts(archive):
    relationships = xml(archive.read("xl/_rels/workbook.xml.rels"))
    by_id = {r.get("Id"): r.get("Target") for r in relationships}
    result = {}
    for sheet in xml(archive.read("xl/workbook.xml")).findall("s:sheets/s:sheet", NS):
        target = by_id[sheet.get(f"{{{R}}}id")]
        result[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    return result


def location(value, names):
    value = unquote(value or "").lstrip("#")
    if "!" not in value:
        return None
    name, anchor = value.rsplit("!", 1)
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    if name not in names or not CELL_RANGE.fullmatch(anchor):
        return None
    return f"{quote_sheetname(name)}!{anchor}"


def financial_fingerprint(path):
    """Hash every package part except hyperlinks; include all cell/style/cache data."""
    digest = hashlib.sha256()
    with ZipFile(path) as archive:
        for part in sorted(archive.namelist()):
            payload = archive.read(part)
            if part.startswith("xl/worksheets/") and part.endswith(".xml"):
                root = xml(payload)
                for links in root.findall(f"{{{S}}}hyperlinks"):
                    root.remove(links)
                payload = ET.tostring(root, method="c14n")
            elif part.startswith("xl/worksheets/_rels/") and part.endswith(".rels"):
                root = xml(payload)
                for relation in list(root):
                    if relation.get("Type", "").endswith("/hyperlink"):
                        root.remove(relation)
                if not len(root):
                    continue
                payload = ET.tostring(root, method="c14n")
            digest.update(part.encode() + b"\0" + payload)
    return digest.hexdigest()


def repair_navigation(path):
    """Convert only bookmarks to this platform's own sheets, keeping URLs/PDFs."""
    path = Path(path)
    before = financial_fingerprint(path)
    replacements, removed, fixed = {}, set(), 0
    with ZipFile(path) as archive:
        sheets = sheet_parts(archive)
        for part in sheets.values():
            root = xml(archive.read(part))
            rel_part = posixpath.dirname(part) + "/_rels/" + posixpath.basename(part) + ".rels"
            relationships = xml(archive.read(rel_part)) if rel_part in archive.namelist() else None
            by_id = {r.get("Id"): r for r in relationships} if relationships is not None else {}
            changed = False
            discarded_ids = set()
            for link in root.findall("s:hyperlinks/s:hyperlink", NS):
                rid = link.get(f"{{{R}}}id")
                target = by_id[rid].get("Target", "") if rid else ""
                bookmark = location(link.get("location") or (target.partition("#")[2]), sheets)
                if not bookmark:
                    continue  # Web and PDF destinations are independent links.
                target_url = urlsplit(target)
                filename = posixpath.basename(unquote(target_url.path).replace("\\", "/"))
                if rid and (target_url.scheme not in {"", "file"} or filename not in SELF_FILES | {path.name, ""}):
                    continue  # Do not retarget a genuine external workbook.
                if rid or link.get("location") != bookmark:
                    link.set("location", bookmark)
                    if rid:
                        del link.attrib[f"{{{R}}}id"]
                        discarded_ids.add(rid)
                    fixed += 1
                    changed = True
            if changed:
                # lxml preserves namespace declarations, including prefixes in
                # mc:Ignorable, which must survive for Microsoft Excel to open it.
                replacements[part] = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                used_ids = set(root.xpath("//@r:id", namespaces={"r": R}))
                if relationships is not None:
                    for relation in list(relationships):
                        if relation.get("Id") in discarded_ids - used_ids:
                            relationships.remove(relation)
                    if len(relationships):
                        replacements[rel_part] = ET.tostring(relationships, encoding="UTF-8", xml_declaration=True, standalone=True)
                    else:
                        removed.add(rel_part)
        if replacements:
            staged = path.with_suffix(".navigation.xlsx")
            try:
                with ZipFile(staged, "w", ZIP_DEFLATED) as out:
                    for info in archive.infolist():
                        if info.filename not in removed:
                            out.writestr(info, replacements.get(info.filename, archive.read(info.filename)))
                assert financial_fingerprint(staged) == before, "Repair changed non-navigation workbook data"
            except BaseException:
                staged.unlink(missing_ok=True)
                raise
    if replacements:
        staged.replace(path)
    return {"repaired": fixed, "financial_fingerprint": before}


def validate_navigation(path):
    internal, external = 0, 0
    with ZipFile(path) as archive:
        sheets = sheet_parts(archive)
        for name, part in sheets.items():
            root = xml(archive.read(part))
            rel_part = posixpath.dirname(part) + "/_rels/" + posixpath.basename(part) + ".rels"
            rels = {r.get("Id"): r.get("Target", "") for r in xml(archive.read(rel_part))} if rel_part in archive.namelist() else {}
            for link in root.findall("s:hyperlinks/s:hyperlink", NS):
                where = f"{name}!{link.get('ref')}"
                rid = link.get(f"{{{R}}}id")
                if rid:
                    target = rels[rid]
                    # Gian Fataya has web sources and a sibling PDF guide, but
                    # no navigation to any external Excel file or build folder.
                    assert not location(link.get("location") or target.partition("#")[2], sheets), f"{where}: bookmark stored as external file {target}"
                    assert not re.search(r"\.xls[xm]?(?:$|#)", target, re.I), f"{where}: external workbook {target}"
                    external += 1
                else:
                    canonical = location(link.get("location"), sheets)
                    assert canonical and canonical == link.get("location"), f"{where}: invalid sheet bookmark {link.get('location')}"
                    internal += 1
    assert internal >= 36, "Missing six-sheet navigation"
    return {"internal": internal, "external": external, "financial_fingerprint": financial_fingerprint(path)}
