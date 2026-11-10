#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_book.py -- directly inspects the VBA stored inside the workbook

    python tools/check_book.py                    # the .xlsm in book/
    python tools/check_book.py path\to\book.xlsm

**Why read the workbook directly instead of using COM**
    VBIDE's CodeModule.Lines **never returns Attribute lines**.
    So via COM, an attribute line exposed as code cannot be detected.
    K-2 actually stepped on this:

      Private WithEvents m_newTabTimer As SafeTimer
      Attribute m_newTabTimer.VB_VarHelpID = -1    <- the genuine one held by the VBE (hidden)
      Attribute m_newTabTimer.VB_VarHelpID = -1    <- a duplicate inserted by AddFromString (exposed)

    The duplicate appears in the VBE code pane and causes a compile error.
    But neither is visible via CodeModule, so it lies "no problems".

    Reading the module streams inside the workbook with olevba reveals the real content.

What it looks at:
    - duplicates of member attribute lines (Attribute <name>.<attribute>)
    - member attribute lines that are not immediately after a declaration line
    - member attributes other than VB_VarHelpID (the VBE does not generate them, so they are lost on import)

Prerequisite: pip install oletools

Exit code: 0 = no problems / 1 = problems found
"""

import os
import re
import sys
import glob
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    print("oletools is required:  pip install oletools")
    sys.exit(1)

MEMBER_ATTR = re.compile(r"^Attribute\s+(\w+)\.(\w+)\s*=")


def check(path):
    problems = []
    vp = VBA_Parser(path)
    try:
        rows = []
        for (_fname, _stream, vba_name, code) in vp.extract_macros():
            lines = code.replace("\r", "").split("\n")
            hits = []
            for i, l in enumerate(lines, 1):
                m = MEMBER_ATTR.match(l)
                if m:
                    hits.append((i, l.strip(), m.group(1), m.group(2)))
            if not hits:
                continue

            dup = sum(v - 1 for v in Counter(h[1] for h in hits).values())
            rows.append((vba_name, len(hits), dup))

            if dup:
                problems.append("%s : member attribute line is duplicated %d time(s) "
                                "(one of them is exposed as code in the VBE)" % (vba_name, dup))
                seen = set()
                for i, text, name, attr in hits:
                    if text in seen:
                        problems.append("    line %d: %s  <- duplicate" % (i, text))
                    seen.add(text)

            for i, text, name, attr in hits:
                prev = lines[i - 2] if i >= 2 else ""
                if not re.search(r"\b" + re.escape(name) + r"\b", prev):
                    if text not in [h[1] for h in hits[:hits.index((i, text, name, attr))]]:
                        problems.append("%s : line %d: %s is not immediately after a declaration line"
                                        % (vba_name, i, text))
                if attr != "VB_VarHelpID":
                    problems.append("%s : line %d: %s is an attribute the VBE does not generate. "
                                    "It may be lost on import" % (vba_name, i, text))

        print("%-24s %8s %8s" % ("Module", "Attr lines", "Dups"))
        print("-" * 44)
        for nm, cnt, dup in rows:
            mark = "  ★" if dup else ""
            print("%-24s %8d %8d%s" % (nm, cnt, dup, mark))
        if not rows:
            print("(no modules with member attributes)")
        print()
    finally:
        vp.close()
    return problems


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if len(argv) > 1:
        books = argv[1:]
    else:
        books = sorted(glob.glob(os.path.join(repo, "book", "*.xlsm")))
    if not books:
        print("No .xlsm to inspect was found")
        return 1

    total = 0
    for b in books:
        print("=== %s ===" % os.path.basename(b))
        problems = check(b)
        if problems:
            print("**Problems**")
            for p in problems:
                print("  " + p)
            total += len(problems)
        else:
            print("No problems")
        print()

    if total:
        print("**Total %d**" % total)
        print()
        print("How to fix duplicates: they are invisible from CodeModule, so DeleteLines cannot remove them.")
        print("  Delete the component with tools/repair_book.ps1 and Import it again.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
