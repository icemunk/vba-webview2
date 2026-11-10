<#
    import.ps1 -- writes the VBA modules in src/ back into the workbook

    Usage:
        powershell -ExecutionPolicy Bypass -File tools\import.ps1                      # preview (default)
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply               # write back
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply -AllowNew     # also allow adding new modules
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply -Modules Wv2Json,Wv2Url

    **The default is a preview (dry run)** It only reports what would change and does not touch the workbook.
    To actually write back, -Apply must be given explicitly (design principle 96).

    Design (the 7 points agreed in E-1):
        1. Before the import, copy the workbook to book\backup\ with a timestamp
        2. Only src\*.bas and src\*.cls are targeted. src\document\ and forms\ are excluded
        3. VBComponents.Remove + Import is not used. Deletion is deferred, and importing the same name
           turns it into Wv2Pane1, which is a trap. Delete the whole CodeModule and pour the text in with AddFromString
           (the component is not deleted, so names and attributes are preserved)
        4. Only modules whose content changed are put in
        5. Pre-gate: run check_cp932.py and abort if even one item appears
        6. Post-verification: re-read the written-back modules and confirm they match src\
        7. Excel is opened invisibly and saved

    "Adding new modules" added in E-3:
        - **Replacing existing ones uses AddFromString, adding new ones uses Import (a hybrid)**
          Since nothing is Removed, the trap of 3. does not apply. With Import, the header attributes
          (VB_PredeclaredId / VB_Exposed etc.) and the name are set correctly and automatically
        - The temporary file passed to Import is **written in CP932**. The pre-gate
          guarantees representability, so this is safe (spec facts 35 / 36)
        - -AllowNew is required for adding new ones. This prevents modules from increasing on their own when junk is left in src\
        - Abort if Attribute VB_Name differs from the file name. Import decides the name by VB_Name,
          so a mismatch creates a module under a different name

    Character encoding:
        AddFromString passes a BSTR (UTF-16), so the file-side encoding is settled
        at the point PowerShell reads it as UTF-8. However, VBA holds the received string
        in the project's code page (CP932), so characters not in
        CP932 are lost here (spec fact 35).
        That is why the pre-gate of 5. is mandatory.
#>

[CmdletBinding()]
param(
    [string]   $Book,
    [string]   $SrcDir,
    [string[]] $Modules,
    [switch]   $Apply,
    [switch]   $AllowNew,
    [switch]   $NoBackup,
    [switch]   $SkipGate,
    [switch]   $SkipHeaderCheck,
    [switch]   $SkipDriftCheck
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Book) {
    $cand = @(Get-ChildItem (Join-Path $repoRoot 'book') -Filter '*.xlsm' -ErrorAction SilentlyContinue)
    if ($cand.Count -eq 0) { throw "No .xlsm was found in the book folder. Specify it with -Book." }
    if ($cand.Count -gt 1) { throw ("There are multiple .xlsm files in the book folder: " + ($cand.Name -join ', ')) }
    $Book = $cand[0].FullName
}
$Book = (Resolve-Path $Book).Path
if (-not $SrcDir) { $SrcDir = Join-Path $repoRoot 'src' }
$SrcDir = (Resolve-Path $SrcDir).Path

$modeLabel = 'preview only (the workbook is not touched)'
if ($Apply) { $modeLabel = '**write back**' }
if ($Apply -and $AllowNew) { $modeLabel = '**write back + allow adding new**' }

Write-Host ("Book    : " + $Book)
Write-Host ("Source  : " + $SrcDir)
Write-Host ("Mode    : " + $modeLabel)
Write-Host ""

$cp932 = [System.Text.Encoding]::GetEncoding(932)

# ---------------------------------------------------------------
# Split a module file into header and body
#   .cls : VERSION line / BEGIN..END block / a run of Attribute VB_xxx
#   .bas : the single Attribute VB_Name line
# Member attributes (Attribute m_browser.VB_VarHelpID) belong to the body side, so they are not dropped
# ---------------------------------------------------------------
function Split-ModuleFile([string]$path) {
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $all = @($text -split "`n")
    if ($all.Count -gt 0 -and $all[$all.Count - 1] -eq '') { $all = @($all[0..($all.Count - 2)]) }

    $i = 0
    while ($i -lt $all.Count) {
        $l = $all[$i]
        if ($l -match '^VERSION\s') { $i++; continue }
        if ($l -eq 'BEGIN') {
            $i++
            while ($i -lt $all.Count -and $all[$i] -ne 'END') { $i++ }
            if ($i -lt $all.Count) { $i++ }
            continue
        }
        if ($l -match '^Attribute VB_[A-Za-z]+ = ') { $i++; continue }
        break
    }

    $header = @()
    if ($i -gt 0) { $header = @($all[0..($i - 1)]) }
    $body = @()
    if ($i -lt $all.Count) { $body = @($all[$i..($all.Count - 1)]) }
    # Drop all trailing blank lines. AddFromString adds one newline,
    # so unless they are dropped, one blank line is added at every import (spec fact 40)
    while ($body.Count -gt 0 -and $body[$body.Count - 1] -eq '') {
        $body = @($body[0..($body.Count - 2)])
    }

    $attrs = @{}
    foreach ($l in $header) {
        if ($l -match '^Attribute (VB_[A-Za-z]+) = (.*)$') {
            $attrs[$Matches[1]] = $Matches[2].Trim()
        }
    }

    return [pscustomobject]@{
        Header = ($header -join "`n")
        Body   = ($body -join "`n")
        Attrs  = $attrs
        Full   = $text
    }
}

# **Member attribute lines (Attribute m_browser.VB_VarHelpID = -1) are also dropped from what is written**
#
#   At first they were kept on the assumption that "AddFromString saves them as attributes",
#   but **that was wrong** AddFromString inserts attribute lines as "plain text".
#   Meanwhile the VBE generates VB_VarHelpID by itself for WithEvents variables,
#   so as a result **the same attribute line appears twice and one of them is exposed as code**.
#   An Attribute line appears in the VBE code pane and causes a compile error.
#
#   **Moreover CodeModule.Lines never returns Attribute lines,
#     so a comparison via COM cannot detect this duplication** (actually stepped on in K-2).
#     To detect it, the real content inside the workbook must be read with olevba (tools/check_book.py).
#
#   Grounds for saying dropping is harmless: m_syncTimer added in K-2 was imported with no attribute line
#   in src/, yet the VBE generated VB_VarHelpID automatically.
#   In other words the WithEvents attribute is under the VBE's jurisdiction and we need not send it.
function Strip-MemberAttrs([string]$s) {
    $keep = @()
    foreach ($l in ($s -split "`n")) {
        if ($l -match '^Attribute\s+\w+\.\w+\s*=') { continue }
        $keep += $l
    }
    return ($keep -join "`n")
}

function Normalize([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace "`r`n", "`n"
    $s = $s -replace "`r", "`n"
    return ($s -replace "`n+$", '')
}

# ---- 5. Pre-gate -------------------------------------------------------
if (-not $SkipGate) {
    $checker = Join-Path $PSScriptRoot 'check_cp932.py'
    if (Test-Path $checker) {
        Write-Host "Pre-gate: check_cp932.py"
        & python $checker $SrcDir
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "**ABORTED** There are characters that cannot round-trip through CP932. Putting them into the workbook loses them permanently (spec fact 35)." -ForegroundColor Red
            Write-Host "Replace them with the replacement candidates above and re-run." -ForegroundColor Red
            exit 1
        }
        Write-Host ""
    } else {
        Write-Host "Warning: check_cp932.py was not found, so the pre-gate was skipped." -ForegroundColor Yellow
    }
    $asciiChecker = Join-Path $PSScriptRoot 'check_ascii.py'
    if (Test-Path $asciiChecker) {
        Write-Host "Pre-gate: check_ascii.py (English fork: source must be ASCII on any code page)"
        & python $asciiChecker $SrcDir
        if ($LASTEXITCODE -ne 0) {
            Write-Host "**ABORTED** Non-ASCII characters would be lost on import. Use ChrW() for non-ASCII data." -ForegroundColor Red
            exit 1
        }
    }
}

# ---- Collect target files -----------------------------------------------
$files = @(Get-ChildItem $SrcDir -File | Where-Object { @('.bas', '.cls') -contains $_.Extension } | Sort-Object Name)
if ($Modules -and $Modules.Count -gt 0) {
    $files = @($files | Where-Object { $Modules -contains $_.BaseName })
    if ($files.Count -eq 0) { throw ("No file matches -Modules: " + ($Modules -join ', ')) }
}

# ---- Issue 5: first check that VB_Name matches the file name -----------
$parsed = @{}
$nameErrors = @()
foreach ($f in $files) {
    $m = Split-ModuleFile $f.FullName
    $parsed[$f.BaseName] = $m
    $vbName = $m.Attrs['VB_Name']
    if ($null -eq $vbName) {
        $nameErrors += ($f.Name + " : there is no Attribute VB_Name")
    } else {
        $vbName = $vbName.Trim('"')
        if ($vbName -ne $f.BaseName) {
        $nameErrors += ($f.Name + " : Attribute VB_Name = " + $vbName + " differs from the file name")
        }
    }
}
if ($nameErrors.Count -gt 0) {
    Write-Host "**ABORTED** Attribute VB_Name and the file name disagree:" -ForegroundColor Red
    foreach ($e in $nameErrors) { Write-Host ("  " + $e) -ForegroundColor Red }
    Write-Host "Import decides the name by VB_Name, so importing with a mismatch creates a module under a different name." -ForegroundColor Red
    exit 1
}

Write-Host ("Targets: " + $files.Count + " file(s)")
Write-Host ""

# ---- 3. Backup (only when writing back) ---------------------------------
if ($Apply -and (-not $NoBackup)) {
    $backupDir = Join-Path (Split-Path -Parent $Book) 'backup'
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($Book) + '_' + $stamp + '.xlsm')
    Copy-Item $Book $dest
    Write-Host ("Backup: " + $dest)
    Write-Host ""
}

# ---- Start Excel --------------------------------------------------------
$excel = $null
$wb = $null
$rows = @()
$changedNames = @()
$addedNames = @()
$blocked = @()
$headerWarn = @()
$drifted = @()
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('wv2import_' + [Guid]::NewGuid().ToString('N'))

try {
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3   # forcibly disable macros. Workbook_Open does not fire

    $readOnly = -not $Apply
    for ($i = 1; $i -le 10; $i++) {
        try { $wb = $excel.Workbooks.Open($Book, 0, $readOnly); break }
        catch { Start-Sleep -Milliseconds 800 }
    }
    if ($null -eq $wb) { throw "Workbooks.Open failed 10 times. Quit Excel once and re-run." }

    # **If it was opened read-only even though we intend to write back, abort**
    #   If the workbook is open in Excel / a stray ~$ lock file remains,
    #   Open succeeds but becomes read-only. Proceeding as is makes both the replacement and the verification
    #   succeed against an "in-memory copy", and **OK is shown although nothing reached the disk**.
    #   (actually stepped on in the D-4b write-back)
    if ($Apply -and $wb.ReadOnly) {
        $wb.Close($false)
        $excel.Quit()
        throw ("The workbook was opened read-only. Aborting because it cannot be written back to disk." +
               "`n  - Check that this workbook is not open in Excel" +
               "`n  - If book\~`$" + [System.IO.Path]::GetFileName($Book) +
               " remains even after closing Excel, delete it")
    }

    $vbp = $null
    try { $vbp = $wb.VBProject } catch { }
    if ($null -eq $vbp) { throw "Cannot access the VBProject. Check the Trust Center settings." }

    $byName = @{}
    foreach ($c in $vbp.VBComponents) { $byName[$c.Name] = $c }

    foreach ($f in $files) {
        $name = $f.BaseName
        $m = $parsed[$name]
        $body = $m.Body
        $bodyCmp = Normalize (Strip-MemberAttrs $body)

        # ================= Adding new (E-3) =================
        if (-not $byName.ContainsKey($name)) {
            if (-not $AllowNew) {
                $blocked += $name
                $rows += [pscustomobject]@{ Module = $name; State = '**not in the workbook (-AllowNew is required)**'; Lines = ($body -split "`n").Count }
                continue
            }
            if (-not $Apply) {
                $rows += [pscustomobject]@{ Module = $name; State = 'scheduled to be added'; Lines = ($body -split "`n").Count }
                continue
            }

            # Write the temporary file in CP932 + CRLF and Import it
            # **CRLF is mandatory** If written with LF only, Import cannot recognize the header block
            #   (VERSION 1.0 CLASS / BEGIN..END), the class module is created as a
            #   standard module and the header falls into the code (stepped on in E-3)
            $tmp = Join-Path $tmpDir ($name + $f.Extension)
            $full = $m.Full -replace "`n", "`r`n"
            if (-not $full.EndsWith("`r`n")) { $full = $full + "`r`n" }
            $bytes = $cp932.GetBytes($full)
            # Round-trip verification (always passes if the pre-gate passed. Double insurance)
            if ($cp932.GetString($bytes) -ne $full) {
                throw ($name + " : the conversion to CP932 is lossy. Check the pre-gate.")
            }
            [System.IO.File]::WriteAllBytes($tmp, $bytes)

            $newComp = $vbp.VBComponents.Import($tmp)
            $actualName = $newComp.Name
            $actualType = $newComp.Type
            $byName[$actualName] = $newComp

            $wantType = 1
            if ($f.Extension -eq '.cls') { $wantType = 2 }

            $state = '**wrong name after adding: ' + $actualName + '**'
            if ($actualName -ne $name) {
                # as is
            } elseif ($actualType -ne $wantType) {
                $state = '**wrong type (expected ' + $wantType + ' / actual ' + $actualType + ')**'
            } else {
                $cmN = $newComp.CodeModule
                $backN = ''
                if ($cmN.CountOfLines -ge 1) { $backN = Normalize $cmN.Lines(1, $cmN.CountOfLines) }
                if ((Normalize (Strip-MemberAttrs $backN)) -eq $bodyCmp) {
                    $state = 'added (verified OK)'
                    $addedNames += $name
                } else {
                    $state = '**added but body verification failed**'
                }
            }
            $rows += [pscustomobject]@{ Module = $name; State = $state; Lines = ('new -> ' + $newComp.CodeModule.CountOfLines) }
            continue
        }

        # ================= Replacing existing =================
        $comp = $byName[$name]
        if (@(1, 2) -notcontains $comp.Type) {
            $rows += [pscustomobject]@{ Module = $name; State = ('**skipped because the type is wrong (Type=' + $comp.Type + ')**'); Lines = '-' }
            continue
        }

        $cm = $comp.CodeModule
        $cur = ''
        if ($cm.CountOfLines -ge 1) { $cur = Normalize $cm.Lines(1, $cm.CountOfLines) }
        $curCmp = Normalize (Strip-MemberAttrs $cur)

        if ($curCmp -eq $bodyCmp) {
            $rows += [pscustomobject]@{ Module = $name; State = 'unchanged'; Lines = $cm.CountOfLines }
            continue
        }

        $changedNames += $name
        $before = $cm.CountOfLines
        $after = ($body -split "`n").Count

        if ($Apply) {
            if ($cm.CountOfLines -ge 1) { $cm.DeleteLines(1, $cm.CountOfLines) }
            # **Pass it with the attribute lines dropped** If left, they double up and get exposed
            $cm.AddFromString((Strip-MemberAttrs $body))
            $back = ''
            if ($cm.CountOfLines -ge 1) { $back = Normalize $cm.Lines(1, $cm.CountOfLines) }
            $state = '**verification after write-back failed**'
            if ((Normalize (Strip-MemberAttrs $back)) -eq $bodyCmp) { $state = 'written back (verified OK)' }
            $rows += [pscustomobject]@{ Module = $name; State = $state; Lines = ($before.ToString() + ' -> ' + $cm.CountOfLines) }
        } else {
            $rows += [pscustomobject]@{ Module = $name; State = 'needs update'; Lines = ($before.ToString() + ' -> ' + $after) }
        }
    }

    # ---- Issue 4: verify the header attributes --------------------------
    # VBIDE has no API to read VB_PredeclaredId etc., so Export and read.
    # There is also no way to set them from VBIDE, so on a mismatch we only warn (fix by hand).
    if (-not $SkipHeaderCheck) {
        foreach ($f in $files) {
            $name = $f.BaseName
            if (-not $byName.ContainsKey($name)) { continue }
            $comp = $byName[$name]
            if (@(1, 2) -notcontains $comp.Type) { continue }
            $hp = Join-Path $tmpDir ('hdr_' + $name + $f.Extension)
            $comp.Export($hp)
            $hb = [System.IO.File]::ReadAllBytes($hp)
            $ht = $cp932.GetString($hb)
            $actual = @{}
            foreach ($l in ($ht -replace "`r`n", "`n" -split "`n")) {
                if ($l -match '^Attribute (VB_[A-Za-z]+) = (.*)$') { $actual[$Matches[1]] = $Matches[2].Trim() }
                elseif ($l -notmatch '^(VERSION|BEGIN|END|\s+\w+ =|Attribute VB_)') { break }
            }
            foreach ($k in $parsed[$name].Attrs.Keys) {
                $want = $parsed[$name].Attrs[$k]
                $got = $actual[$k]
                if ($got -ne $want) {
                    $headerWarn += ($name + " : " + $k + " src=" + $want + " / workbook=" + $got)
                }
            }
        }
    }

    if ($Apply -and ($changedNames.Count + $addedNames.Count) -gt 0) {
        $wb.Save()
        Write-Host ("Saved the workbook (replaced " + $changedNames.Count + " / new " + $addedNames.Count + ")")
        Write-Host ""

    }
}
finally {
    if ($null -ne $wb)    { try { $wb.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($o in @($wb, $excel)) {
        if ($null -ne $o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}

# ---- Post-save drift check (E-3) ----------------------------------------
# **VBA unifies the case of identifiers across the whole project (spec fact 38)**
#   If a new module declares an identifier with the same spelling as an existing one but in a different case,
#   at compile time even the contents of "modules that were not touched" get rewritten.
#   Behavior does not change, but git diff gets dirty and hides the real changes, so it is always surfaced.
#
# **This normalization cannot be observed from the PowerShell process that did the import**
#   Even if Excel is quit and a new COM instance is created, the pre-normalization spelling is returned
#   (measured in E-3). So comparing on our own would lie "no drift".
#   It is arranged to call export.ps1 in a separate process and compare byte by byte.
if ($Apply -and ($changedNames.Count + $addedNames.Count) -gt 0 -and (-not $SkipDriftCheck)) {
    Write-Host "Post-save drift check (re-export in a separate process and compare) ..."
    $dTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('wv2drift_' + [Guid]::NewGuid().ToString('N'))
    try {
        New-Item -ItemType Directory -Path $dTmp -Force | Out-Null
        $expPath = Join-Path $PSScriptRoot 'export.ps1'
        & powershell -ExecutionPolicy Bypass -NoProfile -File $expPath -Book $Book -OutDir $dTmp -Force | Out-Null

        $dSrc = Join-Path $dTmp 'src'
        if (-not (Test-Path $dSrc)) {
            Write-Host "  **The drift check could not be done (export failed)**" -ForegroundColor Yellow
        } else {
            foreach ($f in $files) {
                $nm = $f.BaseName
                if ($changedNames -contains $nm) { continue }
                if ($addedNames -contains $nm) { continue }
                $g = Join-Path $dSrc $f.Name
                if (-not (Test-Path $g)) { continue }
                $h1 = (Get-FileHash $f.FullName -Algorithm MD5).Hash
                $h2 = (Get-FileHash $g -Algorithm MD5).Hash
                if ($h1 -ne $h2) { $drifted += $nm }
            }
            if ($drifted.Count -eq 0) { Write-Host "  No drift" }
        }
    }
    finally {
        if (Test-Path $dTmp) { Remove-Item $dTmp -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-Host ""
}

$rows | Format-Table -AutoSize

$failed = @($rows | Where-Object { $_.State -like '*failed*' -or $_.State -like '*wrong name*' })
Write-Host ("Replaced: " + $changedNames.Count + " / New: " + $addedNames.Count + " / Targets: " + $files.Count)

if ($blocked.Count -gt 0) {
    Write-Host ""
    Write-Host ("Modules not in the workbook: " + ($blocked -join ', ')) -ForegroundColor Yellow
    Write-Host "  If adding them is fine, re-run with -AllowNew." -ForegroundColor Yellow
}

if ($drifted.Count -gt 0) {
    Write-Host ""
    Write-Host "**Modules whose content changed although nothing was written back**" -ForegroundColor Yellow
    Write-Host ("  " + ($drifted -join ', ')) -ForegroundColor Yellow
    Write-Host "  Because VBA unified the case of identifiers across the whole project (spec fact 38)." -ForegroundColor Yellow
    Write-Host "  Behavior does not change, but run export.ps1 -Force to bring them into src\." -ForegroundColor Yellow
    Write-Host "  If you do not, they keep showing up as diffs from now on." -ForegroundColor Yellow
}

if ($headerWarn.Count -gt 0) {
    Write-Host ""
    Write-Host "Header attributes disagree:" -ForegroundColor Yellow
    foreach ($w in $headerWarn) { Write-Host ("  " + $w) -ForegroundColor Yellow }
    Write-Host "  **These attributes cannot be set from VBIDE, so handle them in the VBE.**" -ForegroundColor Yellow
}

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "**Some modules failed verification**" -ForegroundColor Red
    Write-Host "  Restore from the copy in book\backup\." -ForegroundColor Red
    exit 1
}

if (-not $Apply) {
    Write-Host ""
    Write-Host "Note: preview only. The workbook was not touched. Add -Apply to write back."
} else {
    Write-Host ""
    Write-Host "Note: after writing back, it is a good idea to run tools\export.ps1 -Force and confirm it matches src\."
}
