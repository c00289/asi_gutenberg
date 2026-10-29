"""``python -m asi_gutenberg``: import a Word file, list a source's parameters, build a source with values,
or fill the Word original itself with them.

Gutenberg is a library; this is the smallest door for the four things a person does by hand:

    python -m asi_gutenberg import "QF-17 Safety Plan.docx" ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md
    python -m asi_gutenberg params ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md
    python -m asi_gutenberg build ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out --param program=WA485
    python -m asi_gutenberg fill ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out/qf-17.docx --pdf
    python -m asi_gutenberg fill ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out/qf-17.docx --pdf \
        --abridge-headers headings.json

Every command prints one JSON object and exits 0, or prints ``{"ok": false, "error": ...}`` and exits 1.
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import sys
from typing import Any

from .fill_docx import fill_docx
from .importer import from_docx
from .markdown import from_markdown, to_markdown
from .params import parameters
from .pipeline import build


def _parse_params(items: list[str] | None) -> dict[str, str]:
    out: dict[str, str] = {}
    for item in items or ():
        if "=" not in item:
            raise ValueError(f"--param wants name=value, got {item!r}")  # reported as JSON, like every other error
        name, value = item.split("=", 1)
        out[name.strip()] = value
    return out


def _values(args) -> dict[str, Any]:
    """The parameter values of a build or a fill: the files first, a later file or ``--param`` winning."""
    values: dict[str, Any] = {}
    for path in args.params_file or ():
        with open(path, encoding="utf-8") as f:
            loaded = json.load(f)
        if not isinstance(loaded, dict):
            raise ValueError(f"{path}: a parameters file is a JSON object of name: value")
        values.update({str(k): v for k, v in loaded.items()})
    values.update(_parse_params(args.param))
    return values


def _template_beside(md: str) -> str:
    """The one Word original under ``source/`` beside the Markdown source (the asi_standards layout)."""
    found = sorted(glob.glob(os.path.join(os.path.dirname(os.path.abspath(md)), "source", "*.docx")))
    if len(found) != 1:
        raise ValueError(f"--template: {len(found)} .docx under source/ beside {md}; name the Word original")
    return found[0]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="asi_gutenberg", description=__doc__.split("\n")[0])
    sub = parser.add_subparsers(dest="verb", required=True)

    imp = sub.add_parser("import", help="read a .docx into a Gutenberg Markdown source (plus its figures)")
    imp.add_argument("docx")
    imp.add_argument("out_md", help="the Markdown source to write; figures go to figures/ beside it")
    imp.add_argument("--kind", default="plan", help="the document kind (default plan)")
    imp.add_argument("--title", default=None, help="override the title (default: the Word title or the file stem)")

    par = sub.add_parser("params", help="list the {{ name }} placeholders a Markdown source uses")
    par.add_argument("md")

    bld = sub.add_parser("build", help="build a Markdown source into md, pdf and docx, filling its parameters")
    bld.add_argument("md")
    bld.add_argument("out_dir")
    bld.add_argument("--name", default=None, help="the output stem (default: from the title)")
    bld.add_argument("--param", action="append", metavar="NAME=VALUE", help="a parameter value (repeatable)")
    bld.add_argument("--params-file", metavar="JSON", action="append",
                     help="a JSON object of parameter values (repeatable; a later file or --param wins)")
    bld.add_argument("--allow-unfilled", action="store_true", help="build even when a placeholder has no value")
    bld.add_argument("--no-mark", action="store_true", help="do not highlight the filled values")
    bld.add_argument("--landscape", action="store_true",
                     help="lay the pages out landscape (a wide table reads better)")

    fll = sub.add_parser("fill", help="fill the Word original in place from a Markdown source: the template exactly")
    fll.add_argument("md")
    fll.add_argument("out_docx", help="the Word file to write; --pdf prints a PDF beside it through Word")
    fll.add_argument("--template", default=None,
                     help="the Word original (default: the one .docx under source/ beside the Markdown)")
    fll.add_argument("--param", action="append", metavar="NAME=VALUE", help="a parameter value (repeatable)")
    fll.add_argument("--params-file", metavar="JSON", action="append",
                     help="a JSON object of parameter values (repeatable; a later file or --param wins)")
    fll.add_argument("--allow-unfilled", action="store_true", help="fill even when a placeholder has no value")
    fll.add_argument("--mark", action="store_true", help="shade the filled values (off: the template exactly)")
    fll.add_argument("--no-repeat-header", action="store_true",
                     help="do not repeat each table's first row at the top of every page")
    fll.add_argument("--pdf", action="store_true", help="also print the PDF through Word (Windows, Word installed)")
    fll.add_argument("--abridge-headers", metavar="JSON", default=None,
                     help="a JSON object of verbatim column heading -> short form: the table shows the verbatim "
                          "header on its first page and the short forms on every later page (with --pdf)")

    args = parser.parse_args(argv)
    try:
        if args.verb == "import":
            out_md = os.path.abspath(args.out_md)
            folder = os.path.dirname(out_md)
            os.makedirs(folder, exist_ok=True)
            result = from_docx(args.docx, os.path.join(folder, "figures"), kind=args.kind, title=args.title)
            # the figure paths were made relative to the docx's folder; make them relative to the source
            for b in result.document.blocks:
                if b.kind == "figure":
                    docx_dir = os.path.dirname(os.path.abspath(args.docx))
                    b.path = os.path.relpath(os.path.join(docx_dir, b.path), folder)
                    b.path = b.path.replace(os.sep, "/")
            with open(out_md, "w", encoding="utf-8", newline="\n") as f:
                f.write(to_markdown(result.document))
            payload = {
                "ok": True, "verb": "import", "md": out_md, "title": result.document.title,
                "blocks": len(result.document.blocks),
                "tables": sum(1 for b in result.document.blocks if b.kind == "table"),
                "figures": result.figures, "parameters": parameters(result.document), "issues": result.issues,
            }
        elif args.verb == "params":
            with open(args.md, encoding="utf-8") as f:
                doc = from_markdown(f.read())
            payload = {"ok": True, "verb": "params", "md": args.md, "parameters": parameters(doc)}
        elif args.verb == "fill":
            with open(args.md, encoding="utf-8") as f:
                source = f.read()
            template = args.template or _template_beside(args.md)
            abridged = None
            if args.abridge_headers:
                with open(args.abridge_headers, encoding="utf-8") as f:
                    abridged = json.load(f)
            filled = fill_docx(template, source, _values(args), args.out_docx, strict=not args.allow_unfilled,
                               mark=args.mark, repeat_header=not args.no_repeat_header, pdf=args.pdf,
                               abridged_headers=abridged)
            payload = {"ok": True, "verb": "fill", "template": os.path.abspath(template), "docx": filled.docx,
                       "pdf": filled.pdf, "replaced": filled.replaced, "unfilled": filled.unfilled,
                       "issues": filled.issues, "abridged": filled.abridged, "pages": filled.pages}
        else:
            with open(args.md, encoding="utf-8") as f:
                doc = from_markdown(f.read())
            theme = {"orientation": "landscape"} if args.landscape else None
            result = build(doc, args.out_dir, args.name, theme=theme, params=_values(args),
                           strict_params=not args.allow_unfilled, mark_params=not args.no_mark)
            payload = {"ok": True, "verb": "build", **{k: result[k] for k in ("md", "pdf", "docx", "pages", "issues")}}
    except (ValueError, OSError) as exc:
        print(json.dumps({"ok": False, "verb": args.verb, "error": str(exc)}, indent=2))
        return 1
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
