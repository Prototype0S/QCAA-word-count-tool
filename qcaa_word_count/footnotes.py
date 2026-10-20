"""Footnote and endnote extraction and classification."""
from __future__ import annotations

import re
import zipfile
from xml.etree import ElementTree as ET

from .model import Document, Footnote

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

BIBLIOGRAPHIC_HINTS = (
    re.compile(r"(?:19|20)\d{2}"),
    re.compile(r"\b(pp?\.|vol\.|no\.|ed\.|doi|isbn|issn)\b", re.IGNORECASE),
    re.compile(
        r"\b(press|journal|university|thesis|phd|edition|trans\.|"
        r"vols?|pub\.|publishing|http|www\.|accessed)\b",
        re.IGNORECASE,
    ),
    re.compile(r"\bet\s+al\.", re.IGNORECASE),
    # [FIX 7] Retrieval phrasing is a strong bibliographic signal.
    # Pushes "Available at http://..." and "Retrieved 5 March 2020"
    # from AMBIGUOUS up to BIBLIOGRAPHIC.
    re.compile(
        r"\b(?:retrieved|available\s+(?:at|from)|last\s+accessed|"
        r"reprinted|revised)\b",
        re.IGNORECASE,
    ),
)


def classify_footnote(text: str) -> tuple[str, float]:
    """Return (kind, confidence) where kind is bibliographic/commentary/ambiguous."""
    score = 0
    for pattern in BIBLIOGRAPHIC_HINTS:
        if pattern.search(text):
            score += 1
    prose_sentences = len(re.findall(r"[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}", text))
    if score >= 2 and prose_sentences <= 1:
        return "bibliographic", 0.85
    if score <= 1 and prose_sentences >= 2:
        return "commentary", 0.75
    return "ambiguous", 0.4


def read_footnotes_from_docx(path: str) -> list[Footnote]:
    footnotes: list[Footnote] = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            for filename, tag, kind in (
                ("word/footnotes.xml", "footnote", "footnote"),
                ("word/endnotes.xml", "endnote", "endnote"),
            ):
                if filename not in names:
                    continue
                try:
                    root = ET.fromstring(z.read(filename))
                except ET.ParseError:
                    continue
                for note in root.iter(f"{_W}{tag}"):
                    ntype = note.get(f"{_W}type")
                    if ntype in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    note_id = note.get(f"{_W}id") or ""
                    text = "".join(t.text or "" for t in note.iter(f"{_W}t")).strip()
                    if not text:
                        continue
                    footnotes.append(Footnote(
                        footnote_id=note_id, text=text, kind=kind,
                    ))
    except (zipfile.BadZipFile, OSError):
        pass
    return footnotes


def apply_footnote_suggestions(document: Document) -> None:
    for fn in document.footnotes:
        kind, conf = classify_footnote(fn.text)
        fn.suggestion = kind  # type: ignore
        fn.suggestion_confidence = conf