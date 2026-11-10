<#
    repair_book.ps1 -- removes exposed member attribute lines (rebuilds the components)

    Usage:
        powershell -ExecutionPolicy Bypass -File tools\repair_book.ps1 -Modules Wv2Browser
        powershell -ExecutionPolicy Bypass -File tools\repair_book.ps1 -Modules Wv2Browser -Apply

    **The default is a preview** -Apply is required to run (design principle 96).

    What it does:
        1. Copy the workbook to book\backup\
        2. [Phase 1] Open the workbook, Remove the target components, save and close,
           and **quit Excel completely**
        3. [Phase 2] Reopen in a new Excel, Import the files in src\ and save
        4. Confirm with check_book.py that the duplicates are gone

    **Why Remove + Import**
        When a member attribute line is doubled, one of them is exposed in the VBE code pane
        and causes a compile error. But CodeModule.Lines does not
        return Attribute lines, so CodeModule.DeleteLines cannot delete it.
        The only way is to rebuild the whole component.

    **Why it is split into 2 phases**
        VBComponents.Remove defers the deletion. Importing the same name in the same session
        creates it under a different name such as Wv2Browser1 (a trap recorded in E-1).
        If Excel is quit before the Import, you can start with the deletion finalized.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]] $Modules,
    [string]   $Book,
    [string]   $SrcDir,
    [switch]   $Apply,
    [switch]   $NoBackup
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Book) {
    $cand = @(Get-ChildItem (Join-Path $repoRoot 'book') -Filter '*.xlsm' -ErrorAction SilentlyContinue)
    if ($cand.Count -ne 1) { throw "Cannot narrow the .xlsm in the book folder down to one. Specify it with -Book." }
    $Book = $cand[0].FullName
}
$Book = (Resolve-Path $Book).Path
if (-not $SrcDir) { $SrcDir = Join-Path $repoRoot 'src' }
$SrcDir = (Resolve-Path $SrcDir).Path

$cp932 = [System.Text.Encoding]::GetEncoding(932)

Write-Host ("Book    : " + $Book)
Write-Host ("Targets : " + ($Modules -join ', '))
$modeLabel = 'preview only (the workbook is not touched)'
if ($Apply) { $modeLabel = '**rebuild**' }
Write-Host ("Mode    : " + $modeLabel)
Write-Host ""

# ---- Check that the target files exist ----------------------------------
$srcFiles = @{}
foreach ($m in $Modules) {
    $f = $null
    foreach ($ext in @('.cls', '.bas')) {
        $p = Join-Path $SrcDir ($m + $ext)
        if (Test-Path $p) { $f = $p; break }
    }
    if ($null -eq $f) { throw ("Neither " + $m + ".cls nor .bas was found in src") }
    $srcFiles[$m] = $f
    Write-Host ("  " + $m + " <- " + $f)
}
Write-Host ""

if (-not $Apply) {
    Write-Host "Note: preview only. Add -Apply to run."
    Write-Host "   When run, it deletes the target components once and re-imports them from src."
    return
}

# ---- Backup -------------------------------------------------------------
if (-not $NoBackup) {
    $backupDir = Join-Path (Split-Path -Parent $Book) 'backup'
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($Book) + '_repair_' + $stamp + '.xlsm')
    Copy-Item $Book $dest
    Write-Host ("Backup: " + $dest)
    Write-Host ""
}

function Open-Book([bool]$readOnly) {
    $ex = New-Object -ComObject Excel.Application
    $ex.Visible = $false
    $ex.DisplayAlerts = $false
    $ex.EnableEvents = $false
    $ex.AutomationSecurity = 3
    $w = $null
    for ($i = 1; $i -le 10; $i++) {
        try { $w = $ex.Workbooks.Open($Book, 0, $readOnly); break }
        catch { Start-Sleep -Milliseconds 800 }
    }
    if ($null -eq $w) { throw "Workbooks.Open failed. Quit Excel and re-run." }
    return @($ex, $w)
}

function Close-Book($ex, $w, [bool]$save) {
    if ($null -ne $w) {
        try { if ($save) { $w.Save() } } catch { }
        try { $w.Close($false) } catch { }
    }
    if ($null -ne $ex) { try { $ex.Quit() } catch { } }
    foreach ($o in @($w, $ex)) {
        if ($null -ne $o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ---- [Phase 1] Delete ---------------------------------------------------
Write-Host "[Phase 1] Delete the components"
$r = Open-Book $false
$ex1 = $r[0]; $wb1 = $r[1]
try {
    $vbp = $wb1.VBProject
    foreach ($m in $Modules) {
        $comp = $null
        try { $comp = $vbp.VBComponents.Item($m) } catch { }
        if ($null -eq $comp) {
            Write-Host ("  " + $m + " : not in the workbook (skipped)") -ForegroundColor Yellow
            continue
        }
        $vbp.VBComponents.Remove($comp)
        Write-Host ("  " + $m + " : Removed")
    }
}
finally {
    Close-Book $ex1 $wb1 $true
}

Write-Host "  Quitting Excel to finalize the deletion ..."
Start-Sleep -Seconds 3
$still = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($still.Count -gt 0) {
    Write-Host ("  Warning: " + $still.Count + " Excel process(es) still remain") -ForegroundColor Yellow
    Start-Sleep -Seconds 3
}
Write-Host ""

# ---- [Phase 2] Import ---------------------------------------------------
Write-Host "[Phase 2] Re-import from src"
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('wv2repair_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

$r = Open-Book $false
$ex2 = $r[0]; $wb2 = $r[1]
$failed = @()
try {
    $vbp = $wb2.VBProject
    foreach ($m in $Modules) {
        $srcPath = $srcFiles[$m]
        $text = [System.IO.File]::ReadAllText($srcPath, [System.Text.Encoding]::UTF8)
        $text = $text -replace "`r`n", "`n"
        $text = $text -replace "`r", "`n"
        $text = $text -replace "`n", "`r`n"
        if (-not $text.EndsWith("`r`n")) { $text = $text + "`r`n" }

        $bytes = $cp932.GetBytes($text)
        if ($cp932.GetString($bytes) -ne $text) {
            throw ($m + " : the conversion to CP932 is lossy. Check check_cp932.py.")
        }
        $tmp = Join-Path $tmpDir ([System.IO.Path]::GetFileName($srcPath))
        [System.IO.File]::WriteAllBytes($tmp, $bytes)

        $newComp = $vbp.VBComponents.Import($tmp)
        if ($newComp.Name -ne $m) {
            $failed += ($m + " : the name after Import became " + $newComp.Name)
            Write-Host ("  " + $m + " : **the name became " + $newComp.Name + "**") -ForegroundColor Red
        } else {
            Write-Host ("  " + $m + " : Imported (" + $newComp.CodeModule.CountOfLines + " lines)")
        }
    }
}
finally {
    Close-Book $ex2 $wb2 $true
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
if ($failed.Count -gt 0) {
    Write-Host "**Failed**" -ForegroundColor Red
    foreach ($f in $failed) { Write-Host ("  " + $f) -ForegroundColor Red }
    Write-Host "  Restore from the copy in book\backup\." -ForegroundColor Red
    exit 1
}

Write-Host "The rebuild is finished. Confirm with check_book.py:"
Write-Host "  python tools\check_book.py"
