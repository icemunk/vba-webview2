"""Fail if any VBA source file contains a non-ASCII character.

VBA keeps module source in the ANSI code page of the VBA project (1252 on
Western Windows, 932 on Japanese Windows, ...). A character that code page
cannot hold becomes "?" when the module is imported and is lost for good.
Keeping the source pure ASCII makes it safe on every system; non-ASCII data
must be built at runtime with ChrW(&Hxxxx&).

Usage:
    python tools/check_ascii.py [dir ...]      (default: src forms)

Exit code 0 = clean, 1 = non-ASCII found.
"""
import sys
from pathlib import Path

EXTS = {".bas", ".cls", ".frm"}


def main():
    root = Path(__file__).resolve().parent.parent
    dirs = [Path(d) for d in sys.argv[1:]] or [root / "src", root / "forms"]
    bad = 0
    for d in dirs:
        for p in sorted(d.rglob("*")):
            if p.suffix.lower() not in EXTS:
                continue
            text = p.read_bytes().decode("utf-8", errors="replace")
            for n, line in enumerate(text.splitlines(), 1):
                cols = [i for i, ch in enumerate(line) if ord(ch) > 127]
                if cols:
                    bad += 1
                    chars = ", ".join("U+%04X" % ord(line[i]) for i in cols[:5])
                    print("%s:%d: non-ASCII %s" % (p, n, chars))
    if bad:
        print("\n%d line(s) with non-ASCII characters. Use ASCII in comments and "
              "ChrW(&Hxxxx&) for non-ASCII data." % bad)
        sys.exit(1)
    print("No problems: all VBA source is ASCII.")


if __name__ == "__main__":
    main()
