#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_idents.py -- checks whether new identifiers collide with VBA reserved words or existing identifiers

    python tools/check_idents.py LogE LogW LogI LogD LogPath
    python tools/check_idents.py --dump            # list the existing identifiers

Why it is needed:
    (1) VBA is case-insensitive for identifiers. eNum collides with the reserved word Enum (design principle 90).
    (2) **Spec fact 43** When a new module declares the same spelling as an existing one in a different case,
        VBA unifies the spelling across the whole project, and even modules you did not touch get rewritten.

    So for a new identifier, "lowercase it and compare with the reserved words" is not enough;
    it must be "lowercased and compared with all existing identifiers".

Verdict:
    NG   collides with a reserved word
    NG   the same spelling exists, with different case  -> the whole project gets rewritten
    OK*  the same spelling exists, with the same case  -> no rewrite occurs, but the names overlap
    OK   exists nowhere

Exit code: 0 = no problems / 1 = NG found
"""

import os
import re
import sys
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# VBA reserved words and built-ins (kept in lowercase)
RESERVED = set("""
abs addressof and any array as attribute base binary boolean byref byte byval call
case cbool cbyte ccur cdate cdbl cdec chdir chdrive cint clng clngptr clnglng close
collection compare const cos csng cstr currency cvar cvdate cverr date debug decimal
declare defbool defbyte defcur defdate defdbl defdec defint deflng defobj defsng
defstr defvar dim dir do doevents double each else elseif empty end endif enum eqv
erase err error event exit exp explicit false fix for friend function get global
gosub goto if imp implements in input instr int integer is kill lbound lcase left
len let lib like line lock long longlong longptr loop lset ltrim me mid mkdir mod
module name new next not nothing null object on open option optional or paramarray
preserve print private property ptrsafe public put raiseevent randomize redim rem
reset resume return rgb right rmdir rnd rset rtrim seek select set sgn shell single
sin space spc sqr static step stop str strcomp strconv string sub switch tab tan
text then time timer to trim true type typeof ubound ucase unlock until val variant
vartype wend while width with withevents write xor
""".split())

DECL_PATTERNS = [
    # Names of Sub / Function / Property
    re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
               r'(?:Sub|Function)\s+([A-Za-z_]\w*)', re.I),
    re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
               r'Property\s+(?:Get|Let|Set)\s+([A-Za-z_]\w*)', re.I),
    # Variables of Dim / Const / Public / Private
    re.compile(r'^\s*(?:Dim|Const|Public|Private|Global|Static)\s+'
               r'(?:WithEvents\s+)?([A-Za-z_]\w*)', re.I),
    # Names of Enum / Type
    re.compile(r'^\s*(?:Public\s+|Private\s+)?(?:Enum|Type)\s+([A-Za-z_]\w*)', re.I),
    # Declare
    re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?'
               r'(?:Sub|Function)\s+([A-Za-z_]\w*)', re.I),
]

# Formal parameters in argument lists with ByVal / ByRef / Optional
PARAM_RE = re.compile(r'(?:ByVal|ByRef|Optional|ParamArray)\s+([A-Za-z_]\w*)', re.I)


def collect(src_dir):
    """Return the existing identifiers as {lowercase: set(actual spellings)}"""
    table = {}

    def add(name, where):
        table.setdefault(name.lower(), set()).add(name)

    files = sorted(glob.glob(os.path.join(src_dir, "*.bas")) +
                   glob.glob(os.path.join(src_dir, "*.cls")))
    for f in files:
        mod = os.path.splitext(os.path.basename(f))[0]
        add(mod, f)
        with open(f, encoding="utf-8", newline="") as fh:
            text = fh.read()
        for line in text.split("\r\n"):
            s = line.strip()
            if s.startswith("'"):
                continue
            for pat in DECL_PATTERNS:
                m = pat.match(line)
                if m:
                    add(m.group(1), f)
            for m in PARAM_RE.finditer(line):
                add(m.group(1), f)
    return table, files


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src = os.path.join(repo, "src")

    args = [a for a in argv[1:] if not a.startswith("--")]
    dump = "--dump" in argv

    table, files = collect(src)
    print("Collected existing identifiers from %d file(s): %d kinds" % (len(files), len(table)))
    print()

    if dump:
        for k in sorted(table):
            sp = sorted(table[k])
            mark = "  <- multiple spellings" if len(sp) > 1 else ""
            print("  %s%s" % (" / ".join(sp), mark))
        return 0

    if not args:
        print("Pass the identifiers you want to check as arguments. e.g. python tools/check_idents.py LogE LogD")
        return 1

    ng = 0
    print("%-16s %-10s %s" % ("Candidate", "Verdict", "Reason"))
    print("-" * 72)
    for name in args:
        low = name.lower()
        if low in RESERVED:
            print("%-16s %-10s VBA reserved word" % (name, "NG"))
            ng += 1
            continue
        if low in table:
            spellings = table[low]
            if name in spellings and len(spellings) == 1:
                print("%-16s %-10s the same spelling exists (%s)" % (name, "OK*", " / ".join(sorted(spellings))))
            else:
                print("%-16s %-10s **a different case exists: %s -> the whole project gets rewritten**"
                      % (name, "NG", " / ".join(sorted(spellings))))
                ng += 1
            continue
        if len(name) <= 2:
            print("%-16s %-10s no collision. But it is too short (an easy target for spec fact 43)" % (name, "Note"))
            continue
        print("%-16s %-10s no collision" % (name, "OK"))
    print()
    if ng:
        print("**%d NG** Change the name." % ng)
    else:
        print("No NG.")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
