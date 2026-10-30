"""The Word session (0.7.0, 0.8.0): Word itself, driven by COM through PowerShell, prints the PDF of a filled
file and, in the same session, runs the abridge pass that splits a table where its first page ends.

:func:`word_session` is what :func:`~asi_gutenberg.fill_docx.fill_docx` calls with ``pdf=True``; every failure
is an issue in words, never an exception, so a machine without Word (or not Windows) gets its Word file and
the reason. The Word process the session starts is ended before the session returns; one the person has
open is never touched. :func:`pdf_with_word` is the plain print with the same promise (0.7.1).
"""

from __future__ import annotations

import base64
import html
import json
import os
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from typing import Any

WD_EXPORT_PDF = 17  # Word's WdExportFormat.wdExportFormatPDF
ABRIDGE_NEEDS_WORD = "abridged headers need Word (pdf=True on Windows with Word installed)"


# -- the Word session: the abridge pass and the PDF ------------------------------------------------------
@dataclass
class WordSession:
    """What one Word session did: whether the PDF was printed, the tables split (by first heading), Word's
    page count, and the issues (never exceptions) in words."""

    printed: bool = False
    abridged: list[str] = field(default_factory=list)
    pages: int | None = None
    issues: list[str] = field(default_factory=list)


# One PowerShell script per session, Word driven by COM: open the file (a file marked as final opens
# read-only: the mark is lifted for the pass and put back before the save); for each planned table, check
# it is the one expected (its first cell), map every row to its last cell (Rows.Item is refused in a table
# with vertically merged cells; Range.Cells is not), find the last row on the table's first page by Word's
# own layout (Range.Information(3) = wdActiveEndPageNumber, bisected: the rows' pages never decrease), and
# when the table runs on past its first page, freeze its column widths, split it before the next row with
# the Split Table command on a cell of that row (Table.Split is refused there too; a vertical merge cut by
# the split leaves an empty cell at the top of the second part, as the page break showed it), shrink the
# paragraph Word puts between the parts to 1pt (exact, no spacing) with a page break before it, insert a
# row above the second part's first row copying the verbatim header cells' content, formatting and shading
# with the short forms as text, set the heading-row flag off on the first part and on on the second, and
# give the second part the first's column widths (Word lays a split part out afresh); save when anything
# changed (the read-only recommendation the fill lifted put back first); report one JSON line; then print
# the PDF. Repaginate is a courtesy (some files refuse it: "not available because a macro is currently
# running"; Information paginates by itself). Word sometimes drops the COM link right after the export
# (RPC_E_DISCONNECTED on Close or Quit), which is not a failed print, so those two are tried, never
# trusted, and the Word process this started is ended, never one the person has open.
_WORD_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$job = Get-Content -LiteralPath '%(job)s' -Raw -Encoding UTF8 | ConvertFrom-Json
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$report = @{ tables = @(); pages = $null }
try {
  $d = $w.Documents.Open($job.docx, $false, ($job.tables.Count -eq 0))
  $changed = $false; $shift = 0; $wasFinal = $false
  if ($job.tables.Count -gt 0) {
    try { $d.Repaginate() } catch {}
    try { $wasFinal = [bool]$d.Final } catch {}
    if ($wasFinal) { $d.Final = $false }
  }
  foreach ($s in $job.tables) {
    $t = $d.Tables.Item($s.index + $shift)
    $n = $t.Rows.Count
    $first = ($t.Cell(1, 1).Range.Text -replace '[\s\u00a0\r\a]+', ' ').Trim()
    $rec = @{ index = $s.index; first = $first; rows = $n; split = $false; error = $null }
    $rowCell = @{}; $headerCells = 0
    foreach ($c in $t.Range.Cells) { $rowCell[$c.RowIndex] = $c; if ($c.RowIndex -eq 1) { $headerCells += 1 } }
    if ($first -ne $s.first) {
      $rec.error = "Word's table $($s.index) starts '$first', not '$($s.first)'"
    } elseif ($headerCells -ne $s.short.Count) {
      $rec.error = "Word's table $($s.index) has $headerCells header cells, not $($s.short.Count)"
    } else {
      $page1 = $rowCell[1].Range.Information(3)
      $lo = 1; $hi = $n
      while ($lo -lt $hi) {
        $mid = [int][math]::Ceiling(($lo + $hi) / 2)
        if ($rowCell[$mid].Range.Information(3) -eq $page1) { $lo = $mid } else { $hi = $mid - 1 }
      }
      $rec.page = $page1; $rec.on_first_page = $lo; $rec.last_page = $rowCell[$n].Range.Information(3)
      if ($lo -gt 1 -and $lo -lt $n) {
        $tables = $d.Tables.Count
        $t.AllowAutoFit = $false
        $rowCell[$lo + 1].Range.Select()
        $w.Selection.SplitTable()
        $t = $d.Tables.Item($s.index + $shift); $t2 = $d.Tables.Item($s.index + $shift + 1)
        $t2.Cell(1, 1).Range.Select()
        $w.Selection.InsertRowsAbove(1)
        $cells = $w.Selection.Cells.Count
        if ($cells -ne $s.short.Count) {
          # a spanned row where page one ends: the new row cannot carry the header; back to one table
          $tries = 0
          while ($d.Tables.Count -ne $tables -and $tries -lt 20) { $null = $d.Undo(); $tries += 1 }
          $rec.error = "Word's table $($s.index): the row after page one has $cells cells, not $($s.short.Count)"
          $report.tables += $rec
          continue
        }
        $gap = $d.Range($t.Range.End, $t2.Range.Start)
        $gap.Font.Size = 1
        $gap.ParagraphFormat.SpaceBefore = 0; $gap.ParagraphFormat.SpaceAfter = 0
        $gap.ParagraphFormat.LineSpacingRule = 4; $gap.ParagraphFormat.LineSpacing = 1
        $gap.ParagraphFormat.PageBreakBefore = -1
        for ($j = 1; $j -le $s.short.Count; $j++) {
          $from = $t.Cell(1, $j); $to = $t2.Cell(1, $j)
          $to.Range.FormattedText = $from.Range.FormattedText
          $r = $to.Range; $null = $r.MoveEnd(1, -1); $r.Text = [string]$s.short[$j - 1]
          $to.Range.ParagraphFormat = $from.Range.Paragraphs.Item(1).Format
          $to.Shading.Texture = $from.Shading.Texture
          $to.Shading.ForegroundPatternColor = $from.Shading.ForegroundPatternColor
          $to.Shading.BackgroundPatternColor = $from.Shading.BackgroundPatternColor
          $to.VerticalAlignment = $from.VerticalAlignment
        }
        $t.Cell(1, 1).Range.Rows.HeadingFormat = 0
        $t2.Cell(1, 1).Range.Rows.HeadingFormat = -1
        # the second part is laid out afresh by Word: it takes the first part's column widths, column by
        # column (whole columns carry a vertical merge), cell by cell when Word refuses the columns
        $t2.AllowAutoFit = $false
        try { $t2.PreferredWidthType = $t.PreferredWidthType; $t2.PreferredWidth = $t.PreferredWidth } catch {}
        $widths = @(0) * ($s.short.Count + 1)
        for ($j = 1; $j -le $s.short.Count; $j++) { $widths[$j] = $t.Cell(1, $j).Width }
        try {
          for ($j = 1; $j -le $s.short.Count; $j++) { $t2.Columns.Item($j).Width = $widths[$j] }
        } catch {
          foreach ($c in $t2.Range.Cells) {
            if ($c.ColumnIndex -le $s.short.Count) { $c.Width = $widths[$c.ColumnIndex] }
          }
        }
        $rec.split = $true; $changed = $true; $shift += 1
      }
    }
    $report.tables += $rec
  }
  if ($changed) {
    if ($job.read_only_recommended) { $d.ReadOnlyRecommended = $true }
    $d.Save()
    if ($wasFinal) { $d.Final = $true; try { $d.Save() } catch {} }
    try { $d.Repaginate() } catch {}
  }
  try { $report.pages = $d.ComputeStatistics(2) } catch {}
  Write-Output ('REPORT ' + ($report | ConvertTo-Json -Compress -Depth 4))
  if ($job.pdf) { $d.ExportAsFixedFormat($job.pdf, %(pdf_format)d) }
} finally {
  try { $d.Close(0) } catch {}
  try { $w.Quit() } catch {}
  Start-Sleep -Milliseconds 500
  Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
"""


def word_session(docx_path: str, pdf_path: str | None, plan: list[dict[str, Any]] | None = None,
                 timeout: float = 180, read_only_recommended: bool = False) -> WordSession:
    """One Word session over ``docx_path`` (COM automation through PowerShell, Windows only): the abridge
    pass for the tables in ``plan`` (saved into the file when any was split; ``read_only_recommended``
    puts back the recommendation the fill lifted so Word would open the file for writing), then the print
    to ``pdf_path`` when given. Every failure is an issue in words, never an exception."""
    plan = plan or []
    done = WordSession()
    if os.name != "nt":
        done.issues.append("pdf through Word needs Windows with Word installed")
        if plan:
            done.issues.append(ABRIDGE_NEEDS_WORD)
        return done
    # Word prints to a working name beside the target, which takes the target's place only once the print
    # is there: a failed print leaves whatever ``pdf_path`` already was (a rendering of the source, say)
    printing = None
    if pdf_path:
        printing = os.path.splitext(os.path.abspath(pdf_path))[0] + ".printing.pdf"
        if os.path.isfile(printing):
            os.remove(printing)
    job = {"docx": os.path.abspath(docx_path), "pdf": printing, "tables": plan,
           "read_only_recommended": bool(read_only_recommended)}
    # Word keeps an owner file (~$ + the name less two letters) beside a file it has open for writing and
    # leaves it behind when it is ended by force after the export; one that was not there before goes
    folder, name = os.path.split(job["docx"])
    owner = os.path.join(folder, "~$" + name[2:])
    owner_before = os.path.exists(owner)
    fd, job_path = tempfile.mkstemp(prefix="gutenberg-word-", suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(job, f, ensure_ascii=False)
    script = _WORD_SCRIPT % {"job": _ps(job_path), "pdf_format": WD_EXPORT_PDF}
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    run = None
    try:
        run = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
                             capture_output=True, text=True, timeout=timeout)
    except (OSError, subprocess.TimeoutExpired) as exc:
        failure = f"{exc}"
    finally:
        try:
            os.remove(job_path)
        except OSError:
            pass
        if plan and not owner_before and os.path.exists(owner):
            try:
                os.remove(owner)
            except OSError:
                pass
    report = None
    if run is not None:
        for line in (run.stdout or "").splitlines():
            if line.startswith("REPORT "):
                try:
                    report = json.loads(line[len("REPORT "):])
                except ValueError:
                    report = None
        failure = _first_error(run.stderr) or _first_error(run.stdout) or f"exit {run.returncode}"
    if report is not None:
        tables = report.get("tables") or []
        if isinstance(tables, dict):  # PowerShell's JSON writes a one-element list as the element
            tables = [tables]
        for rec in tables:
            if rec.get("error"):
                done.issues.append(f"abridged headers: {rec['error']}")
            elif rec.get("split"):
                done.abridged.append(rec.get("first") or "")
        done.pages = report.get("pages")
    elif plan:
        done.issues.append(ABRIDGE_NEEDS_WORD)
    if printing:
        if os.path.isfile(printing):
            os.replace(printing, pdf_path)
            done.printed = True
        else:
            done.issues.insert(0, f"pdf through Word failed: {failure}")
    return done


def pdf_with_word(docx_path: str, pdf_path: str, timeout: float = 180) -> str | None:
    """Print ``docx_path`` to ``pdf_path`` with Word (COM automation through PowerShell, Windows only);
    None when it did, else the reason as one line, and ``pdf_path`` as it was (0.7.1)."""
    done = word_session(docx_path, pdf_path, None, timeout)
    return None if done.printed else done.issues[0]


def _first_error(text: str | None) -> str | None:
    """The first line of what PowerShell wrote, out of its CLIXML wrapping when it used one."""
    text = (text or "").strip()
    if text.startswith("#< CLIXML"):
        errors = [html.unescape(m).replace("_x000D__x000A_", "\n").strip()
                  for m in re.findall(r'<S S="Error">(.*?)</S>', text, re.S)]
        lines = [ln for e in errors for ln in e.split("\n") if ln.strip()]
        return lines[0] if lines else None
    lines = [ln for ln in text.splitlines() if ln.strip() and not ln.startswith("REPORT ")]
    return lines[0] if lines else None


def _ps(path: str) -> str:
    return os.path.abspath(path).replace("'", "''")
