"""Parameters: ``{{ name }}`` placeholders in a document, listed and filled at build time.

The owner, 2026-10-29: a controlled document imported from Word (:mod:`asi_gutenberg.importer`) is
parametrised by naming its blanks. A placeholder is ``{{ name }}`` (``name`` is ``[a-z_][a-z0-9_]*``,
spaces inside the braces optional) anywhere text goes: the metadata (title, subtitle, doc_id,
revision, date, author, organisation, project, classification, approvers), a heading, a paragraph,
a list item, a table cell, column or caption, a callout, a requirement, a quote, a figure caption,
a raw block. It is plain text in the Markdown source, so it round-trips untouched.

:func:`parameters` lists the names a document uses; :func:`fill` returns a filled copy (the model is
never changed in place) and refuses, naming them, the placeholders it has no value for unless
``strict=False``, in which case they stay and the pipeline reports each as an issue.
"""

from __future__ import annotations

import re
from typing import Any

from .model import Document

PLACEHOLDER = re.compile(r"\{\{\s*([a-z_][a-z0-9_]*)\s*\}\}")


def parameters(doc: Document) -> list[str]:
    """Every placeholder name the document uses, sorted, each once."""
    found: set[str] = set()
    for text in _texts(doc.to_dict()):
        found.update(PLACEHOLDER.findall(text))
    return sorted(found)


def fill(doc: Document, values: dict[str, Any], *, strict: bool = True) -> Document:
    """A copy of ``doc`` with every ``{{ name }}`` replaced by ``str(values[name])``.

    With ``strict`` (the default) a placeholder without a value is a ``ValueError`` that names every
    missing one, so a build never silently ships a blank; with ``strict=False`` they stay as they are.
    A value that is not a string is written with ``str``; ``None`` is an empty string.
    """
    missing = sorted(set(parameters(doc)) - set(values))
    if missing and strict:
        raise ValueError("no value for parameter(s): " + ", ".join(missing))

    def sub(text: str) -> str:
        def one(m: re.Match[str]) -> str:
            name = m.group(1)
            if name not in values:
                return m.group(0)
            v = values[name]
            return "" if v is None else str(v)
        return PLACEHOLDER.sub(one, text)

    data = _map_texts(doc.to_dict(), sub)
    out = Document.from_dict(data)
    out.toc, out.front_matter, out.numbered = doc.toc, doc.front_matter, doc.numbered
    return out


def unfilled(doc: Document) -> list[str]:
    """The placeholder names still in the document: what :func:`fill` with ``strict=False`` left."""
    return parameters(doc)


# -- walking the plain-data form -----------------------------------------------------------------
def _texts(data: Any):
    if isinstance(data, str):
        yield data
    elif isinstance(data, dict):
        for v in data.values():
            yield from _texts(v)
    elif isinstance(data, (list, tuple)):
        for v in data:
            yield from _texts(v)


def _map_texts(data: Any, fn) -> Any:
    if isinstance(data, str):
        return fn(data)
    if isinstance(data, dict):
        return {k: _map_texts(v, fn) for k, v in data.items()}
    if isinstance(data, list):
        return [_map_texts(v, fn) for v in data]
    if isinstance(data, tuple):
        return tuple(_map_texts(v, fn) for v in data)
    return data
