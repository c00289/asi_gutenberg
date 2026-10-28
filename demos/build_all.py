"""Render every demo document in the catalog to out/<name>.{md,pdf,docx},
with their diagrams under out/figures/.  Layout issues (dangling trace links,
degraded diagram layouts) are listed under the document they belong to."""
from __future__ import annotations

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from catalog import DEMOS  # noqa: E402

from asi_gutenberg import build  # noqa: E402


def main(out_dir: str = "out") -> int:
    failures = 0
    for name, builder in DEMOS:
        t0 = time.perf_counter()
        try:
            result = build(builder(), out_dir, name)
        except Exception as exc:  # report the failure and keep going
            failures += 1
            print(f"{name:32s} FAILED: {exc}", file=sys.stderr)
            continue
        issues = result.get("issues") or []
        diagrams = result.get("diagrams") or []
        extra = f"  {len(diagrams)} diagram(s)" if diagrams else ""
        extra += f"  {len(issues)} issue(s)" if issues else ""
        print(f"{name:32s} {result['pages']:4d} pages  {time.perf_counter() - t0:5.1f}s{extra}")
        for issue in issues:
            print(f"    - {issue}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "out"))
