"""Verify that a translated VBA module is code-identical to the original.

Usage:
    python tools/check_translation.py <original_dir> <translated_dir> [file ...]

For every .bas/.cls/.frm file it removes comments, replaces each string literal
with a placeholder and compares what is left, line by line. Any difference in
real code is reported as an ERROR. String literals that changed are listed so a
human can check that only messages / UI text were translated (test inputs,
expected values, JavaScript, keys and selectors must stay as they were).

Exit code 0 = code identical, 1 = code differs.
"""
import re
import sys
from pathlib import Path

STR_RE = re.compile(r'"(?:[^"]|"")*"')


def split_code(line):
    """Return (code_without_comment, [string literals]) for one VBA line."""
    out, lits, i, n = [], [], 0, len(line)
    while i < n:
        c = line[i]
        if c == '"':
            m = STR_RE.match(line, i)
            if not m:  # unterminated - keep the rest as code
                out.append(line[i:])
                break
            lits.append(m.group(0))
            out.append('"S"')
            i = m.end()
            continue
        if c == "'":
            break
        out.append(c)
        i += 1
    code = "".join(out).rstrip()
    # "Rem" comments
    if re.match(r'\s*Rem(\s|$)', code, re.I):
        code = ""
    return code, lits


def body(path):
    text = Path(path).read_text(encoding="utf-8", errors="replace")
    lines = text.replace("\r\n", "\n").split("\n")
    code_lines, lits = [], []
    in_frm_header = path.suffix.lower() == ".frm"
    for ln in lines:
        if in_frm_header:
            # The designer block (Begin ... End) holds captions; compare it as
            # code but let literals differ (captions are UI text).
            if ln.startswith("Attribute VB_Name"):
                in_frm_header = False
        code, l = split_code(ln)
        lits.extend(l)
        code = re.sub(r'\s+', ' ', code).strip()
        if code:
            # VBA is case-insensitive; the IDE re-cases identifiers itself.
            code_lines.append(code.lower())
    return code_lines, lits


def main():
    a_dir, b_dir = Path(sys.argv[1]), Path(sys.argv[2])
    names = [x for x in sys.argv[3:] if not x.startswith("-")] or sorted(
        p.name for p in b_dir.iterdir() if p.suffix.lower() in (".bas", ".cls", ".frm"))
    bad = 0
    for name in names:
        a, b = a_dir / name, b_dir / name
        ca, la = body(a)
        cb, lb = body(b)
        if ca != cb:
            bad += 1
            print(f"ERROR {name}: code differs")
            for i, (x, y) in enumerate(zip(ca, cb)):
                if x != y:
                    print(f"   first difference at code line {i + 1}:\n     was: {x}\n     now: {y}")
                    break
            else:
                print(f"   code line count {len(ca)} -> {len(cb)}")
            continue
        changed = [(x, y) for x, y in zip(la, lb) if x != y]
        print(f"OK    {name}: code identical, {len(changed)} string literal(s) changed")
        if "-v" in sys.argv or "--literals" in sys.argv:
            for x, y in changed:
                print(f"      {x}\n   -> {y}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
