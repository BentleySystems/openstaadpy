# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
"""Sphinx directive that renders the OpenSTAAD error class table by
introspecting ``openstaadpy.os_analytical.oserrors`` at build time.

Usage in an .rst file::

    .. autoerrortable::
"""
from __future__ import annotations

from docutils import nodes
from docutils.parsers.rst import Directive


def _all_subclasses(cls):
    """Yield every direct/indirect subclass of ``cls`` exactly once."""
    seen = set()
    stack = list(cls.__subclasses__())
    while stack:
        sub = stack.pop()
        if sub in seen:
            continue
        seen.add(sub)
        stack.extend(sub.__subclasses__())
        yield sub


def _first_doc_line(cls) -> str:
    doc = (cls.__doc__ or "").strip()
    if not doc:
        return cls.__name__
    for line in doc.splitlines():
        line = line.strip()
        if line:
            return line
    return cls.__name__


def _row(cells):
    row = nodes.row()
    for text in cells:
        entry = nodes.entry()
        entry += nodes.paragraph(text=text)
        row += entry
    return row


class AutoErrorTable(Directive):
    has_content = False
    required_arguments = 0
    optional_arguments = 0

    def run(self):
        from openstaadpy.os_analytical.oserrors import OsErrorBase

        rows = [(OsErrorBase, "(base class, no code)")]
        for cls in _all_subclasses(OsErrorBase):
            try:
                code = cls().code
            except Exception as exc:  # noqa: BLE001
                self.state.document.reporter.warning(
                    f"autoerrortable: could not instantiate {cls.__name__}: {exc}"
                )
                continue
            rows.append((cls, code))

        rows.sort(key=lambda r: (0 if r[0] is OsErrorBase else 1, r[1] if isinstance(r[1], int) else 0, r[0].__name__))

        table = nodes.table(classes=["autoerrortable"])
        tgroup = nodes.tgroup(cols=3)
        table += tgroup
        for width in (25, 55, 15):
            tgroup += nodes.colspec(colwidth=width)

        thead = nodes.thead()
        thead += _row(["Error Class", "Description", "Error Code"])
        tgroup += thead

        tbody = nodes.tbody()
        for cls, code in rows:
            tbody += _row([cls.__name__, _first_doc_line(cls), str(code)])
        tgroup += tbody

        return [table]


def setup(app):
    app.add_directive("autoerrortable", AutoErrorTable)
    return {"version": "1.0", "parallel_read_safe": True, "parallel_write_safe": True}
