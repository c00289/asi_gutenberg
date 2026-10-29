"""The version is written in three places; a half-done bump must fail the suite (as asi_escher 0.1.2 does)."""
import re
import tomllib
from pathlib import Path

from asi_gutenberg import __version__

ROOT = Path(__file__).resolve().parent.parent


def test_version_is_the_same_everywhere():
    pyproject = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    top_entry = re.search(r"^## \[(\d+\.\d+\.\d+)\] - \d{4}-\d{2}-\d{2}$", changelog, re.M).group(1)
    assert pyproject == top_entry == __version__
    assert f"[{pyproject}]: https://github.com/c00289/asi_gutenberg/compare/" in changelog  # its compare link
