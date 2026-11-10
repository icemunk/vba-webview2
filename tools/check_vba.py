#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_vba.py -- formatting and syntax check of VBA sources (standard check no. 1)

    python tools/check_vba.py                 # all of src/
    python tools/check_vba.py src/Wv2Log.bas  # specify a file

What it looks at:
    - BOM / CRLF / bare LF / bare CR
    - trailing whitespace (the VBE drops it, so it becomes a source of diffs)
    - even/odd count of quotation marks
    - a comment on the same line to the right of a line-continuation mark (a syntax error)
    - matching open/close of Sub / Function / Property / Enum / Type
    - balance of If / Do / Select / For / With
    - whether the target of Exit matches the block that is open at that point
    - an &H literal in 0x8000-0xFFFF without the & suffix (design principle 29)
    - one-line Sub / Enum completed with a colon are also regarded as correctly closed

Exit code: 0 = no problems / 1 = problems found
"""

import os
import re
import sys
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OPENERS = {
    "sub": re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Sub\s+\w+', re.I),
    "function": re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Function\s+\w+', re.I),
    "property": re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Property\s+(?:Get|Let|Set)\s+\w+', re.I),
    "enum": re.compile(r'^\s*(?:Public\s+|Private\s+)?Enum\s+\w+', re.I),
    "type": re.compile(r'^\s*(?:Public\s+|Private\s+)?Type\s+\w+', re.I),
    "with": re.compile(r'^\s*With\s+\S', re.I),
    "select": re.compile(r'^\s*Select\s+Case\b', re.I),
}
CLOSERS = {
    "sub": re.compile(r'^\s*End\s+Sub\b', re.I),
    "function": re.compile(r'^\s*End\s+Function\b', re.I),
    "property": re.compile(r'^\s*End\s+Property\b', re.I),
    "enum": re.compile(r'^\s*End\s+Enum\b', re.I),
    "type": re.compile(r'^\s*End\s+Type\b', re.I),
    "with": re.compile(r'^\s*End\s+With\b', re.I),
    "select": re.compile(r'^\s*End\s+Select\b', re.I),
}
DECLARE_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\b', re.I)
IF_OPEN_RE = re.compile(r'^\s*(?:\}|)\s*(?:Else)?If\b.*\bThen\s*$', re.I)
IF_END_RE = re.compile(r'^\s*End\s+If\b', re.I)
FOR_OPEN_RE = re.compile(r'^\s*For\b', re.I)
FOR_END_RE = re.compile(r'^\s*Next\b', re.I)
DO_OPEN_RE = re.compile(r'^\s*Do\b', re.I)
DO_END_RE = re.compile(r'^\s*Loop\b', re.I)
EXIT_RE = re.compile(r'^\s*Exit\s+(Sub|Function|Property|Do|For)\b', re.I)


def strip_comment(line):
    """Drop everything from an apostrophe outside a string literal. Also return the parity of quotation marks."""
    out = []
    in_str = False
    quotes = 0
    i = 0
    while i < len(line):
        c = line[i]
        if c == '"':
            quotes += 1
            in_str = not in_str
            out.append(c)
        elif c == "'" and not in_str:
            break
        else:
            out.append(c)
        i += 1
    return "".join(out), quotes


ONE_LINER = {}
for _k in ("sub", "function", "property", "enum", "type", "with", "select"):
    ONE_LINER[_k] = re.compile(r":\s*End\s+" + _k + r"\b", re.I)


def is_one_liner(code, kind, pat):
    """Whether it is completed on one line with a colon, like Private Sub EntryPoint(): End Sub"""
    m = pat.match(code)
    if not m:
        return False
    return bool(ONE_LINER[kind].search(code))


def check_file(path):
    problems = []

    def bad(lineno, msg, text=""):
        problems.append((lineno, msg, text))

    raw = open(path, "rb").read()
    if raw[:3] == b"\xef\xbb\xbf":
        bad(0, "a BOM is present")
        raw = raw[3:]
    crlf = raw.count(b"\r\n")
    bare_lf = raw.count(b"\n") - crlf
    bare_cr = raw.count(b"\r") - crlf
    if bare_lf:
        bad(0, "%d bare LF" % bare_lf)
    if bare_cr:
        bad(0, "%d bare CR" % bare_cr)

    text = raw.decode("utf-8")
    lines = text.split("\r\n")
    if lines and lines[-1] == "":
        lines = lines[:-1]

    # --- Join line continuations (_) into "logical lines" ---
    # **Without this, an If ... Then spanning several lines is not picked up as an opener,
    #   and End If is falsely detected as "no matching If"**
    logical = []        # (first physical line number, joined code, original line)
    buf = ""
    buf_line = 0
    for i, line in enumerate(lines, 1):
        if line.rstrip() != line:
            bad(i, "trailing whitespace", line)
        code, quotes = strip_comment(line)
        if quotes % 2 != 0:
            bad(i, "odd number of quotation marks", line)
        s = code.rstrip()
        if buf == "":
            buf_line = i
        if s.endswith("_"):
            # **Writing a comment on the same line to the right of a line continuation is a syntax error**
            #   This check can only be done on the "physical line". After folding into a logical line
            #   the continuation line itself no longer reaches us, so it would slip through.
            pos = line.rfind("_")
            if pos >= 0 and "'" in line[pos:]:
                bad(i, "comment on the same line to the right of a line continuation (syntax error)", line)
            buf = buf + s[:-1] + " "
            continue
        logical.append((buf_line, buf + s, line))
        buf = ""
    if buf:
        logical.append((buf_line, buf, ""))

    # --- Check that member attribute lines are not separated from the declaration ---
    #   **Attribute m_foo.VB_VarHelpID = -1 must be "immediately after the declaration line"**
    #   If a line is inserted in between, the VBE cannot absorb it as an attribute and it is exposed as code.
    #   In K-2 a patch inserted a line between a declaration and its attribute and we nearly stepped on this.
    #   It is a classic accident when editing an exported file, so it is always checked.
    prev_code = ""
    for i, code, line in logical:
        s = code.strip()
        m = re.match(r'Attribute\s+(\w+)\.\w+\s*=', s)
        if m:
            name = m.group(1)
            if not re.search(r'\b' + re.escape(name) + r'\b', prev_code):
                bad(i, "member attribute %s is not immediately after the declaration line (the VBE treats it as code)"
                    % name, line)
        if s:
            prev_code = s

    stack = []          # (kind, lineno)
    for i, code, line in logical:

        # &H literal (design principle 29)
        # **Only 0x8000..0xFFFF is dangerous** This range is interpreted as an Integer and
        #   becomes negative. 0x10000 and above is a Long from the start, so there is no problem.
        #   If the & (Long) / ^ (LongLong) suffix is attached, it is already explicit.
        for mh in re.finditer(r'&H([0-9A-Fa-f]+)([&^]?)', code):
            v = int(mh.group(1), 16)
            if 0x8000 <= v <= 0xFFFF and mh.group(2) == "":
                bad(i, "&H%s turns into an Integer. The & suffix is needed (design principle 29)"
                    % mh.group(1), line)

        s = code.strip()
        if not s:
            continue
        if DECLARE_RE.match(code):
            continue

        # Open
        matched = False
        for kind, pat in OPENERS.items():
            if pat.match(code):
                # **One line completed with a colon** e.g. Private Sub EntryPoint(): End Sub
                if is_one_liner(code, kind, pat):
                    matched = True
                    break
                stack.append((kind, i))
                matched = True
                break
        if matched:
            continue

        # Close
        for kind, pat in CLOSERS.items():
            if pat.match(code):
                if not stack:
                    bad(i, "End %s has no matching opener" % kind, line)
                elif stack[-1][0] != kind:
                    bad(i, "End %s but the nearest opener is %s (line %d)" % (kind, stack[-1][0], stack[-1][1]), line)
                    stack.pop()
                else:
                    stack.pop()
                matched = True
                break
        if matched:
            continue

        if IF_OPEN_RE.match(code) and not re.match(r'^\s*ElseIf\b', code, re.I):
            stack.append(("if", i))
        elif IF_END_RE.match(code):
            if stack and stack[-1][0] == "if":
                stack.pop()
            else:
                bad(i, "End If has no matching If", line)
        elif FOR_OPEN_RE.match(code):
            stack.append(("for", i))
        elif FOR_END_RE.match(code):
            if stack and stack[-1][0] == "for":
                stack.pop()
            else:
                bad(i, "Next has no matching For", line)
        elif DO_OPEN_RE.match(code):
            stack.append(("do", i))
        elif DO_END_RE.match(code):
            if stack and stack[-1][0] == "do":
                stack.pop()
            else:
                bad(i, "Loop has no matching Do", line)
        else:
            me = EXIT_RE.match(code)
            if me:
                want = me.group(1).lower()
                kinds = [k for k, _ in stack]
                if want not in kinds:
                    bad(i, "Exit %s but not inside a %s block (what is open is %s)"
                        % (me.group(1), want, "/".join(kinds) or "none"), line)

    for kind, lineno in stack:
        bad(lineno, "%s is not closed" % kind)

    return problems


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    args = argv[1:]
    if args:
        files = []
        for a in args:
            if os.path.isdir(a):
                files += glob.glob(os.path.join(a, "*.bas")) + glob.glob(os.path.join(a, "*.cls"))
            else:
                files.append(a)
    else:
        src = os.path.join(repo, "src")
        files = glob.glob(os.path.join(src, "*.bas")) + glob.glob(os.path.join(src, "*.cls"))
    files = sorted(set(files))

    total = 0
    for f in files:
        problems = check_file(f)
        rel = os.path.relpath(f, repo).replace("\\", "/")
        if problems:
            print("=== %s : %d item(s) ===" % (rel, len(problems)))
            for lineno, msg, txt in problems:
                print("  %s:%d  %s" % (rel, lineno, msg))
                if txt:
                    print("      | %s" % txt.strip()[:110])
            total += len(problems)
        else:
            print("OK  %s" % rel)
    print()
    if total:
        print("**%d problem(s)**" % total)
    else:
        print("No problems (%d files)" % len(files))
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
