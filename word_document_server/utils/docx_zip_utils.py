"""Sanitizers for .docx zip packages produced by docx-editor.

The docx-editor library (v0.1.1) has two known output bugs that corrupt the
OPC package:

1. It writes its workspace state file (``meta.json``) into the unpacked tree
   and its packer zips the whole tree, so ``meta.json`` ends up as a
   top-level OPC part that ``[Content_Types].xml`` does not declare. Word
   then prompts "unreadable content" on open.
2. Its ``delete_comment`` path forgets to remove the matching entry in
   ``word/commentsExtensible.xml``, leaving an orphan
   ``<w16cex:commentExtensible>`` element.

These helpers are wrapper-layer sanitizers — they run after every
docx-editor save to repair the package in place.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path


def strip_meta_json(docx_path: str | Path) -> bool:
    """Remove a top-level ``meta.json`` entry from the .docx zip if present.

    Returns True if an entry was removed, False if there was nothing to do.
    The file is rewritten atomically; the original is replaced only after the
    new zip is fully written.
    """
    docx_path = Path(docx_path)
    if not docx_path.exists():
        return False

    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        if "meta.json" not in names:
            return False

    return _rewrite_zip_without(docx_path, skip_names={"meta.json"})


def strip_orphan_comments_extensible(docx_path: str | Path) -> int:
    """Drop ``<w16cex:commentExtensible>`` entries whose durableId no longer
    has a matching entry in ``word/commentsIds.xml``.

    Returns the number of orphan entries removed. Returns 0 if either part
    is missing from the package (nothing to reconcile).
    """
    docx_path = Path(docx_path)
    if not docx_path.exists():
        return 0

    with zipfile.ZipFile(docx_path, "r") as zin:
        names = set(zin.namelist())
        if "word/commentsExtensible.xml" not in names:
            return 0
        ext_xml = zin.read("word/commentsExtensible.xml").decode("utf-8")
        ids_xml = (
            zin.read("word/commentsIds.xml").decode("utf-8")
            if "word/commentsIds.xml" in names
            else ""
        )

    live_ids = set(re.findall(r'w16cid:durableId="([^"]+)"', ids_xml))
    ext_ids = set(re.findall(r'w16cex:durableId="([^"]+)"', ext_xml))
    orphans = ext_ids - live_ids
    if not orphans:
        return 0

    new_ext = ext_xml
    for durable_id in orphans:
        # Remove either <w16cex:commentExtensible ... durableId="X" .../> or
        # the paired open/close form. durableId is the distinguishing attr.
        pattern = (
            r"<w16cex:commentExtensible\b[^>]*?"
            + r'w16cex:durableId="' + re.escape(durable_id) + r'"'
            + r"[^>]*?/>"
        )
        new_ext = re.sub(pattern, "", new_ext)
        pattern_paired = (
            r"<w16cex:commentExtensible\b[^>]*?"
            + r'w16cex:durableId="' + re.escape(durable_id) + r'"'
            + r"[^>]*?>.*?</w16cex:commentExtensible>"
        )
        new_ext = re.sub(pattern_paired, "", new_ext, flags=re.DOTALL)

    _rewrite_zip_replacing(
        docx_path,
        replacements={"word/commentsExtensible.xml": new_ext.encode("utf-8")},
    )
    return len(orphans)


# ── internals ────────────────────────────────────────────────────────────

def _rewrite_zip_without(path: Path, skip_names: set[str]) -> bool:
    """Copy ``path`` to a tempfile, skipping entries in ``skip_names``, then
    atomically replace the original. Returns True."""
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=path.parent)
    os.close(tmp_fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                if item.filename in skip_names:
                    continue
                zout.writestr(item, zin.read(item.filename))
        shutil.move(str(tmp_path), str(path))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return True


def _rewrite_zip_replacing(path: Path, replacements: dict[str, bytes]) -> None:
    """Copy ``path`` to a tempfile, replacing listed entries' contents, then
    atomically replace the original."""
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=path.parent)
    os.close(tmp_fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
        shutil.move(str(tmp_path), str(path))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
