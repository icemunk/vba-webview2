# -*- coding: utf-8 -*-
"""Extract the vtable order from the WebView2 SDK header.

**Never write a COM vtable index from a guess** (lesson of N-1)
  In the preparation for N-1 add_WebResourceRequested was guessed to be 54, but the correct answer was 55.
  OpenDevToolsWindow had been missed in the count. **Had we tried it on a real machine, a different method would have been called and
  it would have crashed.** Always fix the index with this script.

Usage:
    python tools/vtable.py ICoreWebView2
    python tools/vtable.py ICoreWebView2WebResourceRequest
    python tools/vtable.py --enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT
    python tools/vtable.py --list            (list the IF names)
    python tools/vtable.py --find WebResource (search by name)

Header location:
    By default the candidates below are searched in order. It can also be specified explicitly with the environment variable WEBVIEW2_H.
    **Things in Downloads can disappear**, so when running this, use the environment variable or
    add to HEADER_CANDIDATES in this file.
"""
import io
import os
import re
import sys

HEADER_CANDIDATES = [
    os.environ.get("WEBVIEW2_H", ""),
    # English fork: the original author's personal paths were replaced by these.
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "WebView2.h"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "tmp", "WebView2.h"),
]


def find_header():
    for p in HEADER_CANDIDATES:
        if p and os.path.isfile(p):
            return p
    return None


def load(path):
    with io.open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def vtable(src, iface):
    """Return the vtable of an IF in declaration order. Serial numbers including the 3 IUnknown entries."""
    i = src.find("typedef struct " + iface + "Vtbl")
    if i < 0:
        return None
    j = src.find("} " + iface + "Vtbl", i)
    if j < 0:
        j = i + 200000
    # Pick up the name from a function pointer declaration "... ( STDMETHODCALLTYPE *Name )( ..."
    return re.findall(r"\*\s*(\w+)\s*\)\s*\(", src[i:j])


def enum_values(src, name):
    """Return enum values as (identifier, value). The SDK writes them in the form previous value + 1."""
    i = src.find(name)
    while i >= 0:
        j = src.find("}", i)
        blk = src[i:j]
        if "=" in blk:
            break
        i = src.find(name, i + 1)
    if i < 0:
        return []

    out = []
    n = 0
    for ln in blk.split("\n"):
        m = re.match(r"\s*(\w+)\s*=\s*(.+?),?\s*$", ln)
        if not m:
            continue
        ident, expr = m.group(1), m.group(2)
        m2 = re.match(r"^\s*(-?\d+)\s*$", expr)
        if m2:
            n = int(m2.group(1))
        else:
            n = n + 1 if out else 0
        out.append((ident, n))
    return out


def main():
    path = find_header()
    if not path:
        print("WebView2.h was not found. Specify the location with the environment variable WEBVIEW2_H.")
        print("Locations searched as candidates:")
        for p in HEADER_CANDIDATES:
            if p:
                print("   ", p)
        return 1

    src = load(path)
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        print("Header: %s" % path)
        return 0

    if args[0] == "--list":
        names = sorted(set(re.findall(r"typedef struct (\w+)Vtbl", src)))
        print("Header: %s  (%d items)" % (path, len(names)))
        for n in names:
            print("   ", n)
        return 0

    if args[0] == "--find":
        key = args[1].lower()
        names = sorted(set(re.findall(r"typedef struct (\w+)Vtbl", src)))
        hit = [n for n in names if key in n.lower()]
        print("Header: %s" % path)
        print("IFs containing \"%s\": %d" % (args[1], len(hit)))
        for n in hit:
            print("   ", n)
        return 0

    if args[0] == "--enum":
        vals = enum_values(src, args[1])
        if not vals:
            print("enum %s was not found" % args[1])
            return 1
        print("Header: %s" % path)
        print("=== %s" % args[1])
        for ident, n in vals:
            print("   %3d  %s" % (n, ident))
        return 0

    iface = args[0]
    v = vtable(src, iface)
    if v is None:
        print("IF %s was not found (it can be searched with --find)" % iface)
        return 1

    print("Header: %s" % path)
    print("=== %s  (vtable %d entries. 0-2 are IUnknown)" % (iface, len(v)))
    for k, n in enumerate(v):
        mark = "   " if k >= 3 else " * "
        print("%s%3d  %s" % (mark, k, n))
    return 0


if __name__ == "__main__":
    sys.exit(main())
