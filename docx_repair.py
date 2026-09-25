"""Open .docx files whose internal names are inconsistently cased.

Some submissions arrive with a Word package that is self-inconsistent: a part
stored as ``customXml/item5.xml`` but referenced as ``customXML/item5.xml``, and
zip entries whose local header spells the name differently from the central
directory. Word opens these without complaint; Python's ``zipfile`` refuses,
and the document is lost entirely rather than counted.

Nothing here changes any content. It rewrites only the internal part names so
the package refers to itself consistently, and it is used only after a plain
open has already failed.
"""

from __future__ import annotations

import posixpath
import struct
import zipfile
from io import BytesIO
from typing import Any

from lxml import etree

RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CONTENT_TYPES_PART = "[Content_Types].xml"


def _local_header_name(archive: zipfile.ZipFile, info: zipfile.ZipInfo) -> str | None:
    """The name as spelled in an entry's local header, which may differ."""
    try:
        archive.fp.seek(info.header_offset)
        head = archive.fp.read(30)
    except (OSError, ValueError, AttributeError):
        return None
    if len(head) < 30 or head[:4] != b"PK\x03\x04":
        return None
    (flags,) = struct.unpack("<H", head[6:8])
    (namelen,) = struct.unpack("<H", head[26:28])
    raw = archive.fp.read(namelen)
    if len(raw) != namelen:
        return None
    return raw.decode("utf-8" if flags & 0x800 else "cp437")


def _accept_header_names(archive: zipfile.ZipFile) -> int:
    """Let zipfile read entries whose two recorded names differ only in case.

    ``ZipFile.open`` compares the local header name with ``orig_filename`` and
    refuses on any difference. Setting ``orig_filename`` to what the header
    actually says satisfies that check while leaving ``filename`` - the key
    everything else looks parts up by - alone.
    """
    fixed = 0
    for info in archive.infolist():
        header = _local_header_name(archive, info)
        if header is not None and header != info.orig_filename:
            info.orig_filename = header
            fixed += 1
    return fixed


def _resolve(base: str, target: str) -> str:
    """Where a relationship target points, as a package part name."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target)).lstrip("/")


def _repoint(tree: etree._Element, base: str, actual: dict[str, str]) -> int:
    """Rewrite relationship targets onto the names the archive really uses."""
    changed = 0
    for rel in tree:
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target")
        if not target:
            continue
        resolved = _resolve(base, target)
        if resolved in actual.values():
            continue
        match = actual.get(resolved.lower())
        if match is None:
            continue
        rel.set(
            "Target",
            "/" + match if target.startswith("/") else posixpath.relpath(match, base),
        )
        changed += 1
    return changed


def _retype(tree: etree._Element, actual: dict[str, str]) -> int:
    """Rewrite content-type overrides onto the names the archive really uses."""
    changed = 0
    for node in tree:
        name = node.get("PartName")
        if not name:
            continue
        stripped = name.lstrip("/")
        if stripped in actual.values():
            continue
        match = actual.get(stripped.lower())
        if match is None:
            continue
        node.set("PartName", "/" + match)
        changed += 1
    return changed


def repair(source: Any) -> tuple[bytes, list[str]]:
    """Return a self-consistent copy of a .docx, and what had to be changed."""
    notes: list[str] = []

    with zipfile.ZipFile(source) as archive:
        headers_fixed = _accept_header_names(archive)
        if headers_fixed:
            notes.append(
                f"{headers_fixed} zip entry name(s) were spelled differently in the "
                "archive's index and in the entry itself."
            )

        names = [info.filename for info in archive.infolist()]
        actual = {name.lower(): name for name in names}
        blobs = {name: archive.read(name) for name in names}

    repointed = 0
    for name in names:
        if not name.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(name))
        try:
            tree = etree.fromstring(blobs[name])
        except etree.XMLSyntaxError:
            continue
        changed = _repoint(tree, base, actual)
        if changed:
            repointed += changed
            blobs[name] = etree.tostring(
                tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )
    if repointed:
        notes.append(
            f"{repointed} internal reference(s) pointed at a part whose name is "
            "stored with different capitalisation."
        )

    if CONTENT_TYPES_PART in blobs:
        try:
            tree = etree.fromstring(blobs[CONTENT_TYPES_PART])
        except etree.XMLSyntaxError:
            tree = None
        if tree is not None:
            changed = _retype(tree, actual)
            if changed:
                notes.append(f"{changed} content-type entr(ies) were repointed.")
                blobs[CONTENT_TYPES_PART] = etree.tostring(
                    tree, xml_declaration=True, encoding="UTF-8", standalone=True
                )

    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for name in names:
            out.writestr(name, blobs[name])
    return buffer.getvalue(), notes
