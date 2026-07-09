#!/usr/bin/env python3
"""One-command delivery gate: mechanical fixes + full audit + verdict.

This is the single command to run before delivering any DOCX. It exists because
rules written in prose do not enforce themselves — the gate does:

1. Applies every safe mechanical fix in place (``normalize_docx.py --all``):
   citation brackets, CJK punctuation, number-unit spacing, in-table shading,
   in-table OMML formula size (五号), ``w:tblLook`` flags, header-row repeat,
   ``w:updateFields``.
2. Runs ``audit_docx_format.py`` and prints its full report.
3. Prints the verdict. **DELIVERY GATE: FAIL means the document must not be
   delivered** — exit code is non-zero until every FAIL is fixed.

Usage:
    python scripts/finalize_docx.py report.docx            # fix in place, then audit
    python scripts/finalize_docx.py report.docx --no-fix   # audit only, no changes
    python scripts/finalize_docx.py report.docx --json     # audit output as JSON

Paste the gate verdict and the audit summary line into the delivery note.
"""

from __future__ import annotations

import argparse
import importlib.util
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent


def load_sibling(name: str):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS_DIR / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def main() -> int:
    parser = argparse.ArgumentParser(description="Fix + audit + delivery verdict for a DOCX.")
    parser.add_argument("docx", type=Path, help="Path to the DOCX file to finalize.")
    parser.add_argument("--no-fix", action="store_true", help="Audit only; do not modify the file.")
    parser.add_argument("--json", action="store_true", help="Emit the audit report as JSON.")
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"ERROR: file not found: {args.docx}", file=sys.stderr)
        return 2

    if not args.no_fix:
        normalize = load_sibling("normalize_docx")
        temp = args.docx.with_suffix(".finalize.tmp")
        try:
            counts = normalize.normalize_docx(
                args.docx, temp, units=True, tables=True, update_fields=True
            )
        except (ValueError, OSError) as exc:
            print(f"ERROR: mechanical fixes failed: {exc}", file=sys.stderr)
            return 2
        temp.replace(args.docx)
        applied = {key: value for key, value in counts.items() if value}
        print(f"Mechanical fixes applied: {applied if applied else 'none needed'}")

    command = [sys.executable, str(SCRIPTS_DIR / "audit_docx_format.py"), str(args.docx)]
    if args.json:
        command.append("--json")
    result = subprocess.run(command)

    if result.returncode == 0:
        print(
            "DELIVERY GATE: PASS — no FAIL-level issues. Review WARN lines above; "
            "a visual spot-check in MS Word still applies."
        )
    else:
        print(
            "DELIVERY GATE: FAIL — this document must NOT be delivered until every "
            "FAIL above is fixed. Re-run this command after fixing."
        )
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
