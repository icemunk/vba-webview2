# VBA-WebView2 Development Notes Part 4 — Axis E (Development Environment and Tools)

The axis that deals not with product code but with **the scaffolding of development**. The repository layout, the round trip between workbook and source,
the checking tools, consolidating instructions to Claude, and so on go here. It is a **current place to append**, in parallel with axis D.

As a rule, axis-E work **does not change a single line of the `Wv2*` product code.** If a change turns out to be needed,
it is split out as a stage of axis D or axis K.

---

## Stage E-1 — Moving development to Claude Code and building the round-trip foundation with the workbook

### Where it fits

2026-08-19, right after D-2 was completed. **Zero changes to product code.**

Until now we developed in ordinary chat, with Claude outputting modules that tarboh
downloaded and imported into the VBE, and the development notes were also swapped out by hand.
As the stages advanced, the labor of this hand-off became dominant, so we are moving to Claude Code,
which can read and write local files directly.

### State before the move (what the investigation revealed)

- **The GitHub repository had stopped at the time of 9.10c.** The current modules
  (`Wv2Pane` / `Wv2Element` / `Wv2Json` / `Wv2Tests` / `Wv2TabBar` and others) had
  **not a single one in it.** All that was in there was two files, `Wv2Browser.cls` and
  `Wv2Thunks.bas` as of 7/4, and those were uncommitted.
- The root-level development notes file was the old single-file version. The 4-volume split was not in the repository.
- **The current files were scattered across `Downloads`.** There were dozens of download leftovers with
  sequence numbers like `Wv2Tests_22.bas`, and from the outside there was no way to guarantee which of them matched the contents of the current workbook.
- The current workbook `WebView2_v0.5.2.xlsm` and the 4-volume development notes were also in `Downloads`.
  The line counts of the 4 volumes (509 / 6512 / 6629 / 365) matched what the handover prompt stated, so
  we could confirm that this was the latest.

### Issues and decisions (6 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | **Which side holds the truth** | **The `.cls` / `.bas` in `src/` are authoritative.** They are reflected into the workbook by import. Code's editing, grep, and `git diff` work only on this side |
| 2 | Means of round trip with the workbook | Scripted with PowerShell + Excel COM (VBIDE). `AccessVBOM=1` was **already set**, so no additional configuration was needed |
| 3 | Folder layout | Split by role into `src/` `forms/` `docs/` `book/` `tools/`. The 11 `.xlsm` files scattered at the root were moved to `Old/` |
| 4 | Handling of the old development notes | Moved to `Old/` (they remain in git history, so no information is lost) |
| 5 | Creating `CLAUDE.md` | Moves the role of the handover prompt. **It is loaded automatically every session, so the work of pasting it by hand disappears** |
| 6 | Whether to put the workbook under git | Since `src/` is authoritative, **commit only at milestones**. Adding it every time would pile up 900 KB binaries in the history |

**Forms and document modules are outside the scope of automation** was also agreed together.
`UserForm1` is a pair of `.frm` and `.frx`, and `.frx` is binary, and `ThisWorkbook` / `Sheet1`
cannot be replaced by Import in the first place.

### Confirmed repository layout

| Location | Contents | Nature |
|---|---|---|
| `src/` | 6 standard modules + 9 class modules | **★Authoritative★ edit here** |
| `src/document/` | `ThisWorkbook` / `Sheet1` | For reference (cannot be replaced by Import) |
| `forms/` | `UserForm1` / `TimerForm` | Manual operation. Outside automatic import |
| `book/` | `WebView2_v0.5.2.xlsm` | The current workbook. Commit only at milestones |
| `docs/` | The 4 volumes of development notes + `standard-thunk-method.md` + `handover-at-D2.md` | |
| `tools/` | `export.ps1` / `check_cp932.py` | |
| `Old/` | Past workbooks, modules, and old development notes | **Do not touch** |
| `CLAUDE.md` | Standing instructions to Claude | At the root |

### Round trip with the workbook (`tools/export.ps1`)

```
powershell -ExecutionPolicy Bypass -File tools/export.ps1 -Force
```

- The workbook is opened **ReadOnly**, after setting **`AutomationSecurity = 3`** (macros forcibly disabled).
  This is so that `Workbook_Open` does not fire; without it **merely exporting
  would launch WebView2.**
- It is closed with `SaveChanges:=$false`. The workbook is never modified.
- Sorted by type: standard/class -> `src/`, document -> `src/document/`,
  form -> `forms/` (**only when `-Forms` is given**).
- Right after writing out, it performs a **CP932 -> UTF-8 conversion** and **verifies that converting back to CP932
  matches the original bytes**. If it does not match, it forgoes the conversion, leaves the original as it is, and issues a warning.
- `.frx` is binary and is never touched.

**`import.ps1` for the reverse direction (`src/` -> workbook) has not been written. We will build it in E-2.**
For now the only option is a manual import in the VBE, so **the round trip is not yet closed.**

### ★Discovery: VBA source is held in the project's code page★

When we cross-checked the export results exhaustively against the distributed files, **`⚙` `⟳` `⇄` `↔` `〜` `—` had
all turned into `?`.**

**Isolation**: To determine whether the export dropped them or the workbook had already lost them, instead of `Export` we
**read `CodeModule.Lines` via COM (BSTR = UTF-16)**,
saved it as UTF-8, and compared. The result: **even reading via COM gave `?`.**
That is, **they had already been lost on the workbook side.** It is not an export problem.

On the other hand, 1,517 `★` and also `←` `→` `①②④` remained intact. These are characters that exist in CP932.

**Measurement (all 17 `.bas` + `.cls` files in `src/`)**:

| Item | Count |
|---|---|
| Total `?` | 238 |
| ★Of these, those that were corrupted into `?`★ | **76** |
| Legitimate `?` (URL query strings / JS ternary operators / the Immediate window's `?` shorthand / questions) | 162 |

Breakdown of the 76 corrupted ones:

| Character | Code | Count | Substitute candidate |
|---|---|---|---|
| `〜` | U+301C WAVE DASH | 33 | `～` U+FF5E |
| `⚙` | U+2699 | 28 | `[設定]` ("Settings") |
| `⇄` | U+21C4 | 6 | `⇔` |
| `↔` | U+2194 | 5 | `⇔` |
| `⟳` | U+27F3 | 3 | `[更新]` ("refresh") |
| `—` | U+2014 EM DASH | 1 | `―` U+2015 |

**Scope of the damage**: All of it is inside **comments and the guidance text of `Debug.Print`.**
**Execution logic and HTML/JS are intact.** The tab bar's gear button is drawn with the HTML character reference
`&#9881;`, so **there is no effect at all on how the screen looks.**

We confirmed that every corresponding line in the distributed files exists in `src/` in its `?`-replaced form
(zero lines could not be matched). **Therefore the 76 above are the full count of the damage.**

### The checker `tools/check_cp932.py`

```
python tools/check_cp932.py
```

It scans `src/`, reports **characters that cannot round-trip through CP932** with line and column, and offers substitute candidates.
Exit code 1 if there is a problem.

**★"Can it be encoded" is not enough★** because there are two failure modes:

| Mode | Examples | Result |
|---|---|---|
| Encoding itself is impossible | `⚙` `⟳` `⇄` `↔` `—`, emoji in general | Becomes `?` (noticeable by eye) |
| **Encoding succeeds but turns into a different character** | `〜`->`～`, `−`(U+2212)->`－`, `¢`->`￠` | **Silently replaced** |

Moreover, **Python's `cp932` codec is more lenient than Windows**, and `〜` (U+301C) passes encoding in
Python but **becomes `?` in VBA.** A round-trip check catches both modes.

**Verification of the checker itself**: Against `src/` there were zero problems (naturally, since it comes from the workbook).
So we ran it on a **distributed file containing known corrupting characters** (`Downloads/Wv2TabBar.cls`), and
confirmed that it detected the 4 gears at the correct lines and columns and returned exit code 1.
★A checker can be trusted only after confirming that it fires.★

### Post-implementation report (design principle 79)

1. **Made `export.ps1` not write out user forms by default** (only when `-Forms` is given).
   We learned that the **`.frx` changes in 5-byte units on every export even when its contents are unchanged**
   (apparently an OLE timestamp), and running it every time keeps `git status` dirty.
   This is also consistent with the agreement that forms are manual operation.
2. **When saving the handover prompt into `docs/`, we replaced 5 escape sequences with prose.**
   If `\"` or `<` remain literally in a saved document, they would cause false hits in future greps
   (a problem the handover itself warns about). The meaning was not changed.
3. **Added a note at the top of `docs/handover-at-D2.md` saying "this is a historical record, not current rules."**
   Claude Code has neither `present_files` nor `/home/claude/out/`, so reading it as is would be misleading.

### Checks performed and results

| Check | Result |
|---|---|
| Character encoding of all 17 files in `src/` | UTF-8 / no BOM / CRLF, zero bare LF |
| Round trip to CP932 | **Byte-for-byte match for all 17 files** (`export.ps1` verifies automatically) |
| Full-count cross-check with the distributed files | No version drift. The differences are only the 3 kinds below |
| CP932 round-trip check | Zero problems. The checker fired correctly on the known corrupting characters |

The differences from the distributed files are only the following 3 kinds, and we confirmed that all are **intended differences**:

1. The header lines the export adds (`VERSION 1.0 CLASS` / `Attribute VB_Name` and others) and 2 trailing blank lines
   — this is why the line counts in the handover looked uniformly off by +2. **It was not version drift**
2. **VBA's unification of identifier case** (`Value`->`value` / `Y`->`y` / `ptr`->`Ptr`)
3. Hidden attributes the VBE adds (`Attribute m_browser.VB_VarHelpID = -1`)

### Accidents we stepped on

| Accident | Cause | Fastest diagnosis |
|---|---|---|
| The `.ps1` became a pile of syntax errors | Saved without a BOM, so Windows PowerShell 5.1 read it as CP932 and the Japanese comments were garbled | If the errors are **concentrated on Japanese comment lines**, suspect the character encoding |
| `export.ps1` failed with "cannot convert Byte[] to SwitchParameter" | The local variable `$raw` (a byte array) collided with the switch parameter `$Raw`. **PowerShell too does not distinguish identifier case** | Same type as VBA's `eNum` / `Enum` (design principle 90). **Check new identifiers against parameter names too, lower-cased** |
| `git push` was rejected | The remote had an unfetched commit `a3de35b` (updated through stage 9.29) | **Never force push.** Fetch, read the diff, cross-check the contents, then merge |
| Character-level diff detection timed out | Ran `difflib.SequenceMatcher` at character level on a 170 KB string | Compare large files **line by line** |

### Corrections to the index

Measuring the actual files revealed errors in the index's "current state", so we corrected them.

- **Number of modules 18 -> 19.** The increase of `Wv2Element` in D-2 had not been reflected.
  The diagnosis method in the accident notes, "count the modules (v0_5_2 has 18)", is also corrected to **19**.
  Breakdown: standard 6 + class 9 + form 2 + document 2 = 19.
- **Workbook size 683 KB -> 888 KB** (909,245 bytes).

### Changed files

| File | Contents |
|---|---|
| `CLAUDE.md` | New. Standing instructions that took over the handover prompt's role |
| `tools/export.ps1` | New |
| `tools/check_cp932.py` | New |
| `.gitignore` | New (Excel lock files / backups / OneDrive conflict copies) |
| `.gitattributes` | Appended `binary` for `*.xlsm` `*.xlsx` `*.frx` |
| `src/` `forms/` `docs/` `book/` | Newly created and placed |

**Not a single line of the `Wv2*` product code was changed.**

### Design principles established (94 to 95)

## ★ Design Principle 94 — Judge a round trip that goes through a conversion by "whether what was brought back matches the original" ★

The success of one direction (whether it could be encoded) is not enough. **There are paths, like best-fit conversion, that "succeed yet
become something else."** In CP932, `〜`(U+301C) is silently replaced by `～`(U+FF5E), and
`−`(U+2212) by `－`(U+FF0D). Encoding succeeds, so
there is no error, and it cannot be noticed by eye either.

If you round-trip and check whether it matches the original, **both** the unencodable case and best-fit replacement are caught by the same
judgment. Both `export.ps1`'s conversion verification and `check_cp932.py` are built this way.

Note: as an implementation lesson, **the library used for verification does not necessarily use the same conversion table as production.**
Python's `cp932` is more lenient than Windows' CP932 and lets `〜` through.
A round-trip judgment withstands library differences.

## ★ Design Principle 95 — Turn a development-environment constraint into a checker and keep it permanently, rather than just fixing it on the spot ★

A generalization of design principle 90 (making the reserved-word check permanent). A constraint specific to the environment that we have stepped on once
is often one we **step on next months later**, by which time we have forgotten the circumstances.

The moment we fix it, we drop it into a checker and build it into the standard set of checks. **And we trust the checker
only after confirming that it fires on known bad data.**
That `src/` has no problems is a given (it comes from the workbook, so the characters are already lost), and
that alone is no evidence that the checker is working.

### Spec facts established (35 to 38)

## ★ Spec Fact 35 — VBA source is held in the project's code page, and characters it cannot represent are permanently lost at import ★

In a Japanese environment this is code page 932. Characters that CP932 cannot represent are replaced by `?` at the moment the module is
imported, and never come back. **It is unrelated to the file's character encoding**, and
passing it as UTF-8 gives the same result. There are two failure modes (see design principle 94).

The current workbook already contains this damage to the extent of 76 characters
(`〜`×33 / `⚙`×28 / `⇄`×6 / `↔`×5 / `⟳`×3 / `—`×1).
**All of it is inside comments and the guidance text of `Debug.Print`; execution logic and HTML/JS are intact.**

## ★ Spec Fact 36 — The VBE's Export writes CP932 / CRLF / no BOM ★

The file written by `VBComponent.Export` is in the system's ANSI code page (CP932),
with CRLF line breaks and no BOM. `.frx` is binary.

Conversely, **Import accepts UTF-8 without a BOM** — this is inferred from the track record that,
since 9.x, we have kept importing the UTF-8 files Claude output and the Japanese comments went in correctly, and
we will **confirm it directly in E-2 (import.ps1).**

## ★ Spec Fact 37 — `.frx` changes by a few bytes with every export ★

Even if the form's contents are not changed at all, the contents of `.frx` change every time
(5 bytes for `TimerForm.frx`, 28 bytes for `UserForm1.frx`.
It changes in 5-byte groups, apparently an OLE timestamp).

For this reason exporting forms every time makes the diff dirty. `export.ps1` does not
write them out by default.

## ★ Spec Fact 38 — VBA unifies the case of an identifier to match its first declaration, and rewrites existing occurrences too ★

`Wv2Element`'s `Public Property Get Value()` had become `value` in the workbook, and
`Wv2Thunks`'s `ByVal Y As Long` had become `y` and `ByVal ptr As LongPtr` had become `Ptr`.

**This is because VBA unifies the spelling of same-named identifiers across the whole project.** Behavior does not change, but
exporting changes the look of the source, so **when cross-checking with the distributed files,
do not misread "a diff appeared" as "version drift."**

### Notes for E-2 onward

- **E-2: `import.ps1` (top priority).** Without it, editing `src/` does not reach the workbook.
  Both K-1 and D-3 rewrite VBA, so **we need to close this first.**
  The agreed design:
  1. Before import, copy the workbook to `book/backup/` with a timestamp (already in `.gitignore`)
  2. The targets are only `src/*.bas` and `src/*.cls`. `src/document/` and `forms/` are excluded
  3. ★Do not use `VBComponents.Remove` + `Import`★ There is a known trap where deletion is delayed and an import of the same name
     turns into `Wv2Pane1`. **Delete all of the `CodeModule` and pour in with `AddFromString`**
     (the component is not removed, so the name is preserved)
  4. Import only the modules whose contents changed
  5. Pre-gate: always run `check_cp932.py`, and **abort** if even one item is reported
  6. Post-verification: after import, automatically export again and confirm it matches `src/`, closing the round trip
  7. Excel is opened invisibly and saved
- **Repairing the 76 corrupted characters** is a separate stage. Even if we restore them, importing turns them back into `?`, so
  **replacement with a substitute character that exists in CP932** is needed (candidates are in the breakdown table above).
  The `Debug.Print` guidance "press ? to open the settings tab" does not make sense,
  so we want to clear it up around K-1 (the logging foundation).
- **The procedure for updating `src/document/` and `forms/`** is not yet prepared. In rounds that touch forms,
  the operation will be to run with `-Forms` by hand.

---

## Stage E-2 — Closing the round trip with the workbook using `import.ps1`

### Where it fits

In E-1 we automated workbook -> `src/` but had no reverse direction, so **the round trip was not closed.**
Editing `src/` did not reach the workbook, so neither K-1 nor D-3 could be started. This is the round that closes that gap.

**Zero changes to product code.** All verification was done against a copy of the workbook, and
the md5 of `book/WebView2_v0.5.2.xlsm` never changed from start to finish.

### The design implemented (the 7 points agreed in E-1, as they were)

1. Before import, copy the workbook to `book/backup/` with a timestamp
2. The targets are only `src/*.bas` and `src/*.cls`. `src/document/` and `forms/` are excluded
3. **Do not use `VBComponents.Remove` + `Import`.** Delete all of the `CodeModule` and
   pour in with `AddFromString` (the component is not removed, so the name and attributes are preserved)
4. Import only the modules whose contents changed. If not a single one changed, do not save the workbook
5. Pre-gate: run `check_cp932.py`, and **abort** if even one item is reported
6. Post-verification: read back and cross-check immediately after writing
7. Excel is opened invisibly and saved

### ★Wall 1 we hit: asymmetry between `CodeModule.Lines` and member attribute lines★

In the first preview, 12 of the 15 matched perfectly, but **only 3 had more lines and became "needs update."**

| Module | Workbook | `src/` | Difference |
|---|---|---|---|
| `Wv2Browser` | 2583 | 2587 | +4 |
| `Wv2NavBar` | 1045 | 1046 | +1 |
| `Wv2TabBar` | 1263 | 1264 | +1 |

The cause was **member attribute lines such as `Attribute m_browser.VB_VarHelpID = -1`**.
Counting them in the corresponding files in `src/` gave 4 / 1 / 1, which **exactly matches** the line differences.
All of them appear right after the declaration of a `WithEvents` member.

What we learned by measurement:

- **`CodeModule.Lines` never returns member attribute lines** (the VBE hides them)
- **`Export` writes them out**
- **`AddFromString` accepts them and saves them correctly as member attributes**
  (when we exported again after writing back, they were restored at the same position)

Therefore the right approach is **to keep them for writing and drop them from both sides only for comparison.**
If we dropped them when writing, the attributes themselves would vanish, and conversely if we did not drop them when comparing,
these 3 would stay "needs update" forever.

### ★Wall 2 we hit: `AddFromString` adds one newline at the end★

In the first write test, `Wv2NavBar` went from 1045 to 1047 lines.
After saving and reopening it was 1046 lines, and when exported again the difference from `src/` was **just 1 trailing blank line**.

That is, `AddFromString` adds one newline to the end of the string passed.
**If we pass it without dropping the trailing blank lines, it keeps growing by one line with every import.**
We solved it by having `Get-CodeBody` drop all the trailing blank lines.

Note: the `CountOfLines` immediately after import (1047) differed from the value after saving and reopening (1046).
   **Do not trust `CountOfLines` immediately after writing.**

### Real-machine test results (all passed)

**All were performed against a copy of the workbook.** The real workbook's md5 was unchanged before and after the verification
(`FCF5E0D94B6D242371DFDA00EBF45C56`).

| # | Check | Result |
|---|---|---|
| 1 | Preview against an untouched copy | **All 15 "no change".** Handling of header removal, member attributes, and trailing blank lines was correct, and we also confirmed at the same time that `src/` is faithful to the workbook |
| 2 | Write path (put markers in 2 places of `Wv2Url` and run `-Apply`) | **Only `Wv2Url` was written back, 153 -> 155, and the other 14 were unchanged.** The differential import works |
| 3 | Export again after writing back and compare bytes | **14 were byte-for-byte identical.** The difference in `Wv2Url` was only 1 trailing blank line the Export adds. The markers were at the **intended positions**, line 2 and line 155 |
| 4 | ★Idempotence★ Preview again with the same `src` | **0 / 15.** Does not accumulate |
| 5 | ★Idempotence★ Preview with the re-exported result | **0 / 15.** The import side absorbs the trailing blank line the Export adds |
| 6 | Preview against the real workbook | **0 / 15.** `src/` and the workbook are in sync |

Checks 4 and 5 were **the acceptance condition for this round.** Even if the 1st run succeeds, if a difference appears on the 2nd,
the round trip is not closed.

### Post-implementation report (design principle 79)

1. **Made the default a "preview (dry run)."** An addition not among the 7 agreed points.
   Since this is a tool that rewrites the workbook, unless `-Apply` is given explicitly it opens read-only and
   merely reports "what would change."
2. **Added `-SkipGate` and `-NoBackup`.** Because there were occasions during verification when we wanted to
   skip the pre-gate or the backup. Not used in normal operation.
3. **Did not implement adding new modules.** `AddFromString` only replaces the contents of an existing component,
   so a module that is in `src/` but not in the workbook is reported as
   "not present in the workbook" and skipped. The procedure is to create an empty module in the VBE and then
   run again. There was an idea to create it with `VBComponents.Add`, but
   **as we saw in wall 1, behavior around component creation is hard to read**, so
   we decided not to go in until there is real demand.

### What we learned about character encoding

`AddFromString` takes a BSTR (UTF-16), so **the file-side character encoding problem is
already resolved at the point PowerShell reads it as UTF-8.** Since we do not use `Import`,
spec fact 36's "does Import accept UTF-8 without BOM" is **not tested in this round trip.**
(It is still relevant when importing manually in the VBE.)

However, the fact that VBA holds the string it receives in CP932 does not change, so
**the constraint of spec fact 35 applies as is.** The pre-gate is mandatory.
In check 3, the byte-for-byte match of 14 files containing a large amount of Japanese comments let us confirm that
the whole path `UTF-8 file -> PowerShell -> AddFromString -> workbook -> Export -> UTF-8 file`
is lossless.

### Changed files

| File | Contents |
|---|---|
| `tools/import.ps1` | New |
| `CLAUDE.md` | Added import to the round-trip procedure. Updated the index's line counts and the upper numbers of principles/facts |

### Design principles established (96 to 97)

## ★ Design Principle 96 — Make a destructive tool default to "preview" and require an explicit flag to execute ★

`import.ps1` rewrites the workbook. If the default behavior when arguments are wrong is "rewrite," it
cannot be undone. If the default is preview, even running it with wrong arguments
only prints "what was supposed to change."

We use backups alongside, but backups are about **after an accident happens** and are
not a mechanism for **preventing accidents.** Both are needed.

## ★ Design Principle 97 — Make the acceptance condition of a round-trip tool "no difference on the 2nd run" ★

Even if the first write succeeds, that is no guarantee that the round trip is closed.
In E-2, because `AddFromString` adds a newline at the end, it was broken in the way that **the 1st run succeeds yet
blank lines increase with each repetition.**

If we make idempotence the acceptance condition, such discrepancies "invisible in one direction" are sure to surface.
Verification is applied from both the outbound and return directions (the same input once more / feeding the output back as input).

### Spec facts established (39 to 40)

## ★ Spec Fact 39 — `CodeModule.Lines` does not return member attribute lines, but `Export` writes them out ★

Member attribute lines such as `Attribute m_browser.VB_VarHelpID = -1` are placed
right after the declaration of a `WithEvents` member.

| Path | Attribute lines |
|---|---|
| `CodeModule.Lines` | **Not included** (the VBE hides them) |
| `VBComponent.Export` | **Written out** |
| `CodeModule.AddFromString` | ~~Accepts them and saves them as member attributes~~ |

★**The last row was wrong. We corrected it in K-2.**★
`AddFromString` merely **inserts the attribute line as plain text**, and the VBE separately
generates `VB_VarHelpID` itself. As a result **the same attribute line appears twice, and one of them is
exposed in the VBE's code pane and causes a compile error.**

We also know why we could not catch this in E-2. **`CodeModule.Lines` never returns
Attribute lines**, so the duplicate is invisible in a cross-check via COM.
`Export` also collapses the duplicate and writes out only one, so it is invisible in a round-trip cross-check too.
★It cannot be found without directly reading the module stream inside the workbook★
(we newly created `tools/check_book.py`).

Therefore the line counts of the exported file and the `CodeModule` contents **do not match.**
In the current workbook: `Wv2Browser` 4 / `Wv2NavBar` 1 / `Wv2TabBar` 1.

## ★ Spec Fact 40 — `AddFromString` adds one newline at the end ★

One newline is added to the end of the string passed. If we pass it without dropping the trailing blank lines,
blank lines keep increasing by one with every import.

Also **`CountOfLines` immediately after writing differs from the value after saving** (1047 and 1046 measured).
If verifying by line count, read it after saving and reopening.

### Notes for E-3 onward

- **Adding new modules is not yet supported.** In K-1, which newly creates `Wv2Log.bas`,
  we need to first create an empty standard module in the VBE. If that becomes a bottleneck,
  we will cut an implementation of `VBComponents.Add` as E-3.
- **The procedure for updating `forms/` and `src/document/` is still manual.**
- Replacing the 76 corrupted characters (spec fact 35) can now be
  **executed mechanically**, since `import.ps1` works. We just need to replace them with substitutes that pass the pre-gate and import.

---

## Stage E-3 — Adding new modules (`-AllowNew` of `import.ps1`)

### Where it fits

E-2's `import.ps1` **only replaced the contents of existing components**, and skipped modules that are in `src/` but not in
the workbook. Because K-1 newly creates `Wv2Log.bas`,
the procedure "create an empty module by hand in the VBE, then import" remained. This is the round that eliminates it.

**Zero changes to product code.** All verification was done against a copy of the workbook, and
the md5 of `book/WebView2_v0.5.2.xlsm` never changed from start to finish
(`FCF5E0D94B6D242371DFDA00EBF45C56`).

### ★The first switch★ We chose `Import` rather than `VBComponents.Add`

tarboh's instruction was to implement `VBComponents.Add`, but based on the preliminary investigation we
**presented the option of using `Import`, and switched with agreement** (the lesson of 9.30 / D-1).

| Option | Description | Assessment |
|---|---|---|
| A | `VBComponents.Add` + rename + `AddFromString` | **There is no way to set the header attributes.** Right now all classes are at defaults so there is zero actual harm, but if we someday create a class with `VB_PredeclaredId = True`, **something different would be silently created** |
| **B (adopted)** | Write a temporary file in **CP932 + CRLF** and `VBComponents.Import` | The proper API for creating new components. **The header attributes and the name are both set correctly and automatically** |

What E-1 / E-2 avoided with `Import` was the trap of the "**`Remove` + `Import`**" combination
(deletion is delayed and an import of the same name turns into `Wv2Pane1`), and **it does not apply to adding a new
one without a `Remove`.** Replacement of existing ones stays with `AddFromString` as before, and
we made it a hybrid in which **only for new ones do we use `Import`.**

In the preliminary investigation we had confirmed that the current 9 classes **all have default header attributes**
(`PredeclaredId=False` / `GlobalNameSpace=False` / `Creatable=False` / `Exposed=False` /
`MultiUse=-1`). The `Attribute VB_Name` of the `.bas` files also matched
the file name in every file.

### Issues and decisions (7 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Means of adding new | **A temporary file in CP932 + CRLF + `Import`** (above) |
| 2 | Explicit flag | **Make `-AllowNew` mandatory.** Prevents things from being added on their own when junk is left in `src/` (an extension of design principle 96) |
| 3 | Target types | Only standard `.bas` and class `.cls` |
| 4 | Verification of header attributes | **Cross-check regardless of new or existing, and warn if they differ.** `import.ps1` had compared only the body, so there was a hole that missed attribute drift |
| 5 | `VB_Name` and file name | **Abort if they do not match.** `Import` determines the name from `VB_Name` |
| 6 | On failure | Guide to recovery from the backup and exit with code 1 |
| 7 | Experiment on the side | Check whether `Import` accepts UTF-8 without BOM, and close the unconfirmed part of spec fact 36 |

### ★Wall 1 we hit: writing the temporary file with LF turns a class into a standard module★

In the first write test, the `.bas` succeeded but the `.cls` was **created as a standard module.**
The body should have been 12 lines but was 16 lines, and the header had fallen into the code:

```
  1| Attribute VB_Name = "E3TestCls"
  2| VERSION 1.0 CLASS
  3| BEGIN
  4|   MultiUse = -1  'True
  5| End                      <- traces of VBA normalizing case. Evidence it was interpreted as code
  6| ' --- E3TestCls.cls ...
```

The header's byte sequence **exactly matched** a real `.cls`, so the cause was not the contents.
The true cause was that **`Split-ModuleFile` normalized line breaks to LF, and that string was written out
as CP932 as it was.** `Import` **does not recognize the header block
(`VERSION 1.0 CLASS` / `BEGIN..END`) unless it is CRLF**, and treats it as a standard module.

We solved it by writing the temporary file with CRLF. We also made it **cross-check the type after adding**
(looking only at the name lets this kind of breakage pass).

### ★Wall 2 we hit: adding a new module rewrites "modules we did not touch"★

When we wrote `Public Property Get N() As Long` with an **uppercase `N`** in the test class,
**VBA rewrote the `n` of the whole project to `N`.**
A total of 70 lines in `Wv2Browser` / `Wv2Pane` / `Wv2TabBar` / `Wv2Thunks`
changed **even though import did not touch them at all**.

```
-    Dim n As Long
-    n = m_tabs.count
+    Dim N As Long
+    N = m_tabs.count
```

Spec fact 38 (unification of identifier case) had been recorded in D-2 / E-1, but
we had not written about the fang that **"adding a new module rewrites all the other modules."**
Behavior does not change (VBA does not distinguish case), but **`git diff` becomes dirty and hides the real changes.**

### ★Wall 3 we hit: our own drift check quietly lied★

To detect wall 2, we wrote a check that reopens the workbook after saving and cross-checks.
But **it reported "no drift" even though drift was actually happening.**

- Close and reopen in the same Excel instance -> not visible
- Quit Excel and create a new COM instance **from the same PowerShell process** -> not visible
- **`export.ps1` in a separate process** -> visible

★This normalization cannot be observed from the process that performed the import.★
We dropped our own comparison and switched to **calling `export.ps1` in a separate process and cross-checking bytes.**
It lets us reuse already-verified code, and not only case unification but also drift in trailing blank lines and header attributes
is caught by the same mechanism.

**A checker that does not fire is worse than no checker** (design principle 95).
Had we let the "no drift" lie pass without ever doubting it, we would have stepped on this wall in production.

### Real-machine test results (all passed)

**All performed against a copy of the workbook.** The real workbook's md5 was unchanged before and after.

| # | Check | Result |
|---|---|---|
| 1 | Preview without `-AllowNew` | **Refused.** Reports "not in the workbook (-AllowNew is needed)" and does not add |
| 2 | `-AllowNew -Apply` | **Added 2 files, a `.bas` and a `.cls`.** Cross-check OK |
| 3 | Export again and confirm type and contents | **The `.cls` is a class module, the `.bas` is a standard module.** Byte counts also match the originals (459 / 318). Japanese and `★ ← → ⇔ ① ～ ―` are also intact |
| 4 | Drift detection | **Correctly enumerated the 4 files `Wv2Browser` / `Wv2Pane` / `Wv2TabBar` / `Wv2Thunks`** |
| 5 | ★Convergence★ Take the drift into `src/` and preview | **0 / 17.** The round trip converges |
| 6 | `VB_Name` mismatch | **Aborted with exit code 1.** The workbook is unchanged |
| 7 | Header attribute mismatch (`PredeclaredId=True`) | **Detected even when the body is identical.** Reported "src=True / workbook=False" |
| 8 | Pass UTF-8 without BOM to `Import` | **Garbled** (below) |
| 9 | Preview against the real workbook (regression) | **0 / 15.** Same as E-2 |

### Experiment for issue 7 — `Import` does not accept UTF-8 without BOM

When we directly `Import`ed a `.bas` in UTF-8 without BOM and read it back, we got typical mojibake.

```
Passed    : ' UTF-8 BOM なしで Import した場合に日本語が無傷か: 日本語テスト ★①～
Read back : ' UTF-8 BOM 縺ｪ縺励〒 Import 縺励◆蝣ｴ蜷医↓譌･譛ｬ隱槭′辟｡蛯ｷ縺・ ...
```

**`VBComponents.Import` reads the file as the ANSI code page (CP932).**
This closes the unconfirmed part of spec fact 36. At the same time it settled that **the design of writing the temporary file in CP932
was not a "just in case" but a necessity.**

Note: how it was that, in the chat era, tarboh imported the UTF-8 files Claude output and
the Japanese survived remains **unexplained**. Possibly the VBE's GUI (File -> Import File) makes a different determination
from the programmatic `Import`, or the operation may have been pasting the code.
It **does not affect the current round trip**, so we do not pursue it.

### Post-implementation report (design principle 79)

1. **Added `-SkipDriftCheck`.** The drift check runs one extra export,
   so we made it possible to drop it when in a hurry. Not used in normal operation.
2. **Made it cross-check the "type" after adding.** Wall 1 passes through if we look only at the name.
3. **Made it issue a warning when the drift check could not be run.**
   The first implementation was built to **skip silently** with `if ($null -ne $wb2)`,
   and this delayed the discovery of wall 3. **Do not build branches that silently skip a failure.**

### Changed files

| File | Contents |
|---|---|
| `tools/import.ps1` | Added `-AllowNew` / `-SkipHeaderCheck` / `-SkipDriftCheck`. Adding new modules, header attribute cross-check, drift check |
| `CLAUDE.md` | Updated the explanation of the round trip |

**Not a single line of the `Wv2*` product code was changed.**

### Design principle established (98)

## ★ Design Principle 98 — Verify by the same path as production (observation within the same process lies because of caches) ★

We stepped on this in wall 3. VBA's identifier normalization could not be observed from the PowerShell process that performed the import,
**even if we quit Excel and created a new COM instance.**
The reality becomes visible only when opened in a separate process.

When we write our own "reread and compare" processing, it cannot get past layers of such caching and
returns a **plausible-looking lie.** Calling a tool that already works on the production path (here `export.ps1`)
as a child process is shorter to write and also correct.

### Spec facts established (41 to 43)

## ★ Spec Fact 41 — `VBComponents.Import` reads the file as the ANSI code page ★

If you pass UTF-8 without BOM, the Japanese becomes mojibake (measured).
`import.ps1` writes the temporary file in **CP932** and then passes it to `Import`.
Whether it can be represented in CP932 is guaranteed by the pre-gate (`check_cp932.py`).

## ★ Spec Fact 42 — The file passed to `Import` must be CRLF ★

If written with LF only, the header block `VERSION 1.0 CLASS` / `BEGIN..END` is not recognized, and
**a class module is created as a standard module and the header falls into the code.**
It happens even when the header's byte sequence is correct, so suspect the line-ending code.

## ★ Spec Fact 43 — Adding a new module rewrites the spelling of identifiers across the whole project ★

A continuation of spec fact 38. If a new module declares an identifier with the same spelling as an existing one but in different case,
**at compile time even the contents of "modules we did not touch" are rewritten.**
In E-3's measurement, the test class's `Property Get N()` changed a total of 70 lines in
`Wv2Browser` / `Wv2Pane` / `Wv2TabBar` / `Wv2Thunks` from `n` -> `N`.

★**This normalization cannot be observed from the process that performed the import**★
Even if you quit Excel and create a new COM instance, the pre-normalization spelling comes back.
You need to open it in a separate process (design principle 98).

Behavior does not change but `git diff` becomes dirty, so **`import.ps1` cross-checks and reports using `export.ps1` in a separate
process.** When a report appears, take it into `src/` with `export.ps1 -Force`.

### Notes for E-4 onward

- **Updating `forms/` and `src/document/` is still manual.**
- When adding a new module, being careful **not to declare an identifier with the same spelling as an existing one in a different case**
  avoids wall 2. Even when it cannot be avoided, it is detected.
- Replacing the 76 corrupted characters (spec fact 35) can now be fully mechanized.

---

## Stage E-4 — Escaping from OneDrive, and duplicated attribute lines

### Where it fits

During the K-2 work we caused an **accident that put the workbook into a compile error**. There were two causes,
and both were problems of axis E (development environment and tools), so we record them separately from the bug fix (K-2).

- **(1) `import.ps1` was duplicating member attribute lines** — a defect in the tool
- **(2) The workbook was on OneDrive, and Excel's AutoSave crushed the contents we had fixed** — an environment problem

**Zero changes to product code.**

### ★Accident 1: duplicated member attribute lines★

#### What happened

We received a report from tarboh: "`Attribute` lines are exposed in the VBE. There are 4 places.
It's a compile error and nothing can run."

As a result of directly reading the module streams inside the workbook:

```
572| Private WithEvents m_newTabTimer As SafeTimer
573| Attribute m_newTabTimer.VB_VarHelpID = -1    <- the genuine one the VBE holds (hidden)
574| Attribute m_newTabTimer.VB_VarHelpID = -1    <- the duplicate AddFromString inserted (exposed)
575| Private m_pendingUrls As Collection
```

`Wv2Browser` had 5 duplicates (`m_newTabTimer` / `m_settingsTimer` / `m_syncTimer` /
`m_prewarmTimer` / `m_adoptTimer`). The other modules were intact.

#### True cause — the second half of spec fact 39 was wrong

In E-2 we recorded **"`AddFromString` accepts attribute lines and saves them as member attributes"**,
but **this was wrong.**

`AddFromString` merely **inserts the attribute line as plain text**, and
the VBE **separately generates `VB_VarHelpID` itself** for a `WithEvents` variable.
As a result the same attribute line appears twice, and **one of them is exposed as code.**

#### ★Why we could not catch it in E-2 (this is the main point)★

In E-2 we built up even the round-trip cross-check, yet only this duplicate slipped through. There are two reasons.

| Path | Is the duplicate visible? |
|---|---|
| `CodeModule.Lines` | ★**Not visible**★ Never returns Attribute lines |
| `VBComponent.Export` | ★**Not visible**★ Collapses the duplicate and writes out only one |
| Directly reading the streams inside the workbook | **Visible** |

**Both verification means I had were structurally hiding this anomaly.**
Both the post-verification of `import.ps1` and the comparison against `export.ps1` kept answering "no anomaly",
and we went back and forth many times while remaining out of line with the reality tarboh was seeing in the VBE.

★**Information that an API hides cannot be verified by that API**★ — an obvious but
hard-to-notice-for-oneself kind of blind spot.

#### What we fixed

- **`import.ps1`**: **Drop attribute lines** from the body passed to `AddFromString`.
  The grounds: `m_syncTimer`, added in K-2, was **auto-generated by the VBE** even when imported with no attribute line in `src/`.
  `WithEvents` attributes are in the VBE's jurisdiction, and we have no need to send them.
- **`tools/check_book.py` (new)**: Uses `olevba` to read the module streams inside the workbook
  directly, and detects duplicated attribute lines / attributes separated from their declarations / attributes other than `VB_VarHelpID`.
  ★A tool for seeing what is invisible via `CodeModule`★
- **`tools/repair_book.ps1` (new)**: An exposed attribute line cannot be seen from `CodeModule`, so
  it cannot be removed with `DeleteLines`. We rebuild in two stages: **`Remove` the component -> fully quit Excel ->
  reopen and `Import`**. Excel is interposed to avoid the trap recorded in E-1 that
  "the deletion by `Remove` is delayed and an `Import` of the same name turns into `Wv2Pane1`".
- **`tools/check_vba.py`**: Added a rule that a member attribute line must be **immediately after its declaration line**.
  The trigger of the accident was that K-2's patch inserted between a declaration and its attribute, orphaning one of them.

#### Verification

- Rebuilt `Wv2Browser` with `repair_book.ps1`, **duplicates 5 -> 0**
- ★In a disposable copy, confirmed that "writing back with the fixed `import.ps1` produces no duplicates"★
  (even changing 2757 -> 2759 lines, the attribute lines stayed at 5). Before the fix, 5 more would have appeared here
- Also confirmed that `check_book.py` fires correctly on the known broken state

### ★Accident 2: OneDrive's AutoSave crushed the contents we had fixed★

#### What happened

In the middle of accident 1, I normalized the workbook with `import.ps1` (22:40), but
the attribute lines were still visible in tarboh's VBE. The measurement (disk) and
the display (Excel) disagreed, and the isolation dragged on.

The reason became clear when we looked at the path Excel opened the workbook with.

```
Opened workbook: https://d.docs.live.net/555c0e1ccf55908a/ドキュメント/GitHub/.../WebView2_v0.5.2.xlsm
```

★**It was opened with a cloud URL, not a local path**★
This was because the repository was under `OneDrive\ドキュメント\` (Known Folder Move put the whole Documents folder
into OneDrive).

And **for files on OneDrive, Excel's AutoSave is ON by default.**
The edit in which tarboh commented out the problem spot was **autosaved without being noticed**,
and at 23:01 it overwrote the normal version I had fixed at 22:40.

#### What we fixed

**We moved the whole repository out of OneDrive.**

```
C:\Users\gugug\GitHub\WebView2-For-Excel-VBA
```

- With `robocopy`, 835 / 835 files, 0 failures (48.4 MB)
- `git status` clean, `HEAD` and `origin/main` match
- `export.ps1` / `import.ps1` find the repository root from their own location, so **no change was needed**
- **Confirmed Excel opens it with the local path.** With this AutoSave does not operate

★**"Unchecking it in OneDrive's selective sync" is dangerous**★ Unchecking it
**removes the files from local** (they remain only in the cloud).
It is not behavior of "leave it in place and just stop syncing." **Moving it out first is the right order.**

#### Why it was double management

What was under version control on GitHub was also being synced by OneDrive.
GitHub was enough for backup, and the OneDrive side added only the harms of
"Excel opens it with a cloud URL" and "AutoSave runs."

### Post-implementation report (design principle 79)

1. **Made it a copy, not a move.** If we leave the original we can go back. The procedure is to delete
   the old one after operation is confirmed.
2. **Made the default of `repair_book.ps1` a preview** (design principle 96). It is a tool that deletes components, so
   a default of "delete" when arguments are wrong could not be undone.
3. **Newly introduced `oletools`.** There was no other way to read the inside of the workbook.
   ★Product code remains at zero external libraries; this is a dependency only on the checking-tool side★

### Changed files

| File | Contents |
|---|---|
| `tools/import.ps1` | Drop attribute lines from what is written |
| `tools/check_book.py` | New. Directly inspects the inside of the workbook |
| `tools/repair_book.ps1` | New. Rebuilds components |
| `tools/check_vba.py` | Added a rule for attribute line position |
| `CLAUDE.md` | Created an OneDrive section. Made the standard checks a 5-point set |
| (repository) | Moved out of OneDrive |

### Design principle established (102)

## ★ Design Principle 102 — Information that an API hides cannot be verified by that API ★

`CodeModule.Lines` deliberately hides `Attribute` lines. So **as long as we use `CodeModule`,
an anomaly in attribute lines can never be detected.** `Export` also collapses duplicates, so it is likewise invisible.

In E-4, because of this blind spot we kept reporting "no anomaly" and went back and forth many times while out of line with what
tarboh was seeing in the VBE.

**When choosing a means of verification, first ask "does this means hide the target?"**
If there is a layer that hides it, **prepare a tool that looks from outside that layer** (here `check_book.py`,
which directly reads the inside of the workbook with `olevba`).

It is a continuation of design principles 98 (verify by the same path as production) and 100 (doubt the way of measuring), but
the difference is that this is **not a problem of path; that path structurally drops information.**

### Spec facts established (47 to 48)

## ★ Spec Fact 47 — An attribute line passed to `AddFromString` is duplicated ★

`CodeModule.AddFromString` inserts `Attribute m_foo.VB_VarHelpID = -1` as
**plain text**. The VBE also generates the same attribute itself for a `WithEvents` variable, so **the same line appears twice, and one of them is exposed as code
and causes a compile error.**

★**This duplicate is invisible from both `CodeModule.Lines` and `Export`**★
The former never returns Attribute lines, and the latter collapses the duplicate and writes out only one.
To detect it, you need to directly read the module streams inside the workbook
(`tools/check_book.py`).

**Countermeasure**: Drop attribute lines on the import side. `VB_VarHelpID` is auto-generated by the VBE, so
there is no need to send it. Ones that have already become exposed are invisible from `CodeModule`, so
they cannot be removed with `DeleteLines`, and the whole component is rebuilt (`tools/repair_book.ps1`).

Note: this corrects the second half of spec fact 39.

## ★ Spec Fact 48 — A workbook on OneDrive is opened by Excel with a cloud URL, and AutoSave is ON ★

When you open a workbook under `OneDrive\ドキュメント\`, `Workbook.FullName` returns
`https://d.docs.live.net/...`. It is not a local path.

As a result:

- **AutoSave is ON by default**, and **edits are written back without being noticed**
- Excel keeps holding the version at the time it was opened, so **updating the local copy is not reflected**
- Contents fixed by a tool can be crushed by the AutoSave of an Excel holding an old copy

**Countermeasure**: Place the workbook (more precisely, the whole repository) outside OneDrive.
If it is opened with a local path, AutoSave does not operate.

### Notes for E-5 onward

- **Updating `forms/` and `src/document/` is still manual.**
- **Deleting modules became possible with `repair_book.ps1`**, but as a regular
  operation it is not yet in place. Whether to integrate it into `import.ps1` is for after real demand appears.
- A dependency on `oletools` was added. It is only on the checking-tool side so it does not affect the product,
  but a new environment needs `pip install oletools`.
