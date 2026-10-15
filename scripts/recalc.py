"""Optional: have LibreOffice evaluate every formula and store the results.

Usage:
    python scripts/recalc.py path/to/output.xlsx

The workbooks this tool writes already ask Excel to recalculate when opened
(fullCalcOnLoad), so this step is NOT needed for normal use. It is only useful
if the file will be read by software that does not calculate formulas.

LibreOffice re-saves the whole workbook and can drop or change PivotTables,
so the original file is only replaced when the re-saved copy still has the
same sheets and PivotTables. Otherwise the original is left untouched.
"""

import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def _find_soffice():
    candidates = [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        'soffice',
        'libreoffice',
    ]
    return next((c for c in candidates if shutil.which(c) or Path(c).exists()), None)


def _structure(path: Path) -> tuple[int, int]:
    """(number of sheets, number of pivot tables) in a workbook."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    sheets = sum(1 for n in names if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'))
    pivots = sum(1 for n in names if n.startswith('xl/pivotTables/pivotTable') and n.endswith('.xml'))
    return sheets, pivots


def recalc(xlsx_path: str | Path) -> bool:
    """Recalculate with LibreOffice. Returns True when the file was replaced
    by the recalculated copy, False when it was left as it was (LibreOffice
    missing, failed, or the copy lost sheets / PivotTables)."""
    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        print(f"[recalc] File not found: {xlsx_path}", file=sys.stderr)
        return False

    soffice = _find_soffice()
    if not soffice:
        print("[recalc] LibreOffice not found — skipped. This is fine: Excel "
              "recalculates the workbook itself when it opens it.", file=sys.stderr)
        return False

    with tempfile.TemporaryDirectory() as tmp:
        cmd = [soffice, '--headless', '--norestore', '--convert-to', 'xlsx',
               '--outdir', tmp, str(xlsx_path)]
        print(f"[recalc] Running: {' '.join(str(c) for c in cmd)}")
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        except subprocess.TimeoutExpired:
            print("[recalc] LibreOffice timed out — original file kept.", file=sys.stderr)
            return False
        converted = Path(tmp) / xlsx_path.name
        if result.returncode != 0 or not converted.exists():
            print(f"[recalc] LibreOffice failed — original file kept.\n{result.stderr}",
                  file=sys.stderr)
            return False

        before, after = _structure(xlsx_path), _structure(converted)
        if after[0] < before[0] or after[1] < before[1]:
            print(f"[recalc] Recalculated copy lost sheets/pivot tables "
                  f"(before {before}, after {after}) — original file kept.",
                  file=sys.stderr)
            return False

        shutil.copy2(converted, xlsx_path)
    print("[recalc] Done. Workbook replaced by the recalculated copy.")
    return True


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python scripts/recalc.py <path_to_xlsx>")
        sys.exit(1)
    sys.exit(0 if recalc(sys.argv[1]) else 1)
