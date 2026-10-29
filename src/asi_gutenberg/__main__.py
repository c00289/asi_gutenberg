"""``python -m asi_gutenberg``: import a Word file, list a source's parameters, build a source with values.

Gutenberg is a library; this is the smallest door for the three things a person does by hand:

    python -m asi_gutenberg import "QF-17 Safety Plan.docx" ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md
    python -m asi_gutenberg params ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md
    python -m asi_gutenberg build ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out --param program=WA485

Every command prints one JSON object and exits 0, or prints ``{"ok": false, "error": ...}`` and exits 1.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any

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
        else:
            with open(args.md, encoding="utf-8") as f:
                doc = from_markdown(f.read())
            values: dict[str, Any] = {}
            for path in args.params_file or ():
                with open(path, encoding="utf-8") as f:
                    loaded = json.load(f)
                if not isinstance(loaded, dict):
                    raise ValueError(f"{path}: a parameters file is a JSON object of name: value")
                values.update({str(k): v for k, v in loaded.items()})
            values.update(_parse_params(args.param))
            theme = {"orientation": "landscape"} if args.landscape else None
            result = build(doc, args.out_dir, args.name, theme=theme, params=values,
                           strict_params=not args.allow_unfilled, mark_params=not args.no_mark)
            payload = {"ok": True, "verb": "build", **{k: result[k] for k in ("md", "pdf", "docx", "pages", "issues")}}
    except (ValueError, OSError) as exc:
        print(json.dumps({"ok": False, "verb": args.verb, "error": str(exc)}, indent=2))
        return 1
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
