#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_cp932.py -- checks that the VBA sources in src/ contain no "characters that break when put into the workbook"

    python tools/check_cp932.py                 # check src/ and src/document/
    python tools/check_cp932.py path\to\file    # check a specified file/folder

Why it is needed:
    VBA source is held in the project's code page (CP932 in a Japanese environment).
    Characters that CP932 cannot handle are lost at import time and never come back.
    There are two failure modes.

      (1) cannot be encoded at all      -> becomes ?            e.g. gear U+2699, check mark U+2713
      (2) encodes but turns into another char -> silently replaced    e.g. U+301C -> U+FF5E

    (2) cannot be noticed by eye. So the judgment is not "can it be encoded" but
    "does it come back as the same character after a round trip".

    Note that Python's cp932 codec is more lenient than Windows; U+301C passes in Python but
    becomes ? in VBA. A round-trip check catches both modes.

Exit code:
    0 = no problems / 1 = problems found
"""

import os
import sys
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Candidates for rewriting garbled characters (confirmed to round-trip in CP932)
SUGGEST = {
    "⚙": "[設定]",      # gear
    "⟳": "[更新]",      # rotating arrow
    "⇄": "⇔",      # left-right arrow -> ⇔
    "↔": "⇔",      # ↔ -> ⇔
    "✓": "○",      # ✓ -> ○
    "✗": "×",      # ✗ -> ×
    "➡": "→",      # ➡ -> →
    "—": "―",      # — (EM DASH) -> ― (HORIZONTAL BAR)
    "〜": "～",      # 〜 (WAVE DASH) -> ～ (FULLWIDTH TILDE)
    "−": "－",      # − (MINUS SIGN) -> －
    "¢": "￠",      # ¢ -> ￠
    "£": "￡",
    "¬": "￢",
}

EXTS = (".bas", ".cls", ".frm")


def classify(ch):
    """Round-trip ch through CP932 to judge it. Returns None if there is no problem."""
    try:
        enc = ch.encode("cp932")
    except UnicodeEncodeError:
        return "cannot encode (becomes ?)"
    back = enc.decode("cp932")
    if back != ch:
        return "turns into another character -> %s (U+%04X)" % (back, ord(back))
    return None


def collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for ext in EXTS:
                files.extend(glob.glob(os.path.join(p, "**", "*" + ext), recursive=True))
        elif os.path.isfile(p):
            files.append(p)
    return sorted(set(files))


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    targets = argv[1:] if len(argv) > 1 else [os.path.join(repo, "src")]

    files = collect(targets)
    if not files:
        print("Nothing to inspect was found: %s" % ", ".join(targets))
        return 1

    problems = []
    for f in files:
        # Without newline="" CRLF is normalized to LF and the columns shift
        with open(f, encoding="utf-8", newline="") as fh:
            text = fh.read()
        for lineno, line in enumerate(text.split("\r\n"), 1):
            for col, ch in enumerate(line, 1):
                if ord(ch) < 0x80:
                    continue
                why = classify(ch)
                if why:
                    problems.append((f, lineno, col, ch, why, line.strip()))

    rel = lambda p: os.path.relpath(p, repo).replace("\\", "/")

    print("Files inspected: %d" % len(files))
    print()

    if not problems:
        print("No problems. There is not a single character that cannot round-trip through CP932.")
        return 0

    print("**%d problem(s)** There are characters that would be lost when imported into the workbook" % len(problems))
    print()
    for f, lineno, col, ch, why, src in problems:
        sug = SUGGEST.get(ch)
        tail = ("  Replacement candidate: %s" % sug) if sug else ""
        print("%s:%d:%d" % (rel(f), lineno, col))
        print("    %s (U+%04X)  %s%s" % (ch, ord(ch), why, tail))
        print("    | %s" % src[:110])
    print()

    # Tally per character
    from collections import Counter
    tally = Counter(p[3] for p in problems)
    print("Count per character:")
    for ch, n in tally.most_common():
        sug = SUGGEST.get(ch)
        tail = ("  -> %s" % sug) if sug else ""
        print("  %s (U+%04X) x%d%s" % (ch, ord(ch), n, tail))
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
