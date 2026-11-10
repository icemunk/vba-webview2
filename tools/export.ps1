<#
    export.ps1 -- writes the VBA modules of the workbook out to files

    Usage:
        powershell -ExecutionPolicy Bypass -File tools\export.ps1
        powershell -ExecutionPolicy Bypass -File tools\export.ps1 -OutDir C:\temp\dump -Raw

    Prerequisites:
        - Excel's "Trust access to the VBA project object model" is ON
          (HKCU:\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1)
        - Excel is not running

    Output:
        src\           standard modules (.bas) / class modules (.cls)
        src\document\  ThisWorkbook / Sheet* (for reference, since import cannot replace them)
        forms\         UserForms (.frm + .frx) -- only when -Forms is given
                       .frx changes by a few bytes on every export, so it is not touched by default

    Character encoding:
        The VBE's Export writes in CP932 (the system's ANSI code page).
        This script converts it to UTF-8 / no BOM / CRLF and saves it.
        After the conversion it converts back to CP932 and verifies that it matches the original bytes,
        and reports it as an anomaly if even one byte differs.
        .frx is binary, so it is not touched at all.

    Safety:
        - The workbook is opened ReadOnly and closed with SaveChanges:=$false. The workbook is not changed
        - Macros are forcibly disabled with AutomationSecurity (Workbook_Open does not fire)
        - There is no operation that deletes files (same names are overwritten; -Force is required if any exist)
#>

[CmdletBinding()]
param(
    [string] $Book,
    [string] $OutDir,
    [switch] $Raw,
    [switch] $Forms,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# ---- Determine paths ---------------------------------------------------
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Book) {
    $cand = @(Get-ChildItem (Join-Path $repoRoot 'book') -Filter '*.xlsm' -ErrorAction SilentlyContinue)
    if ($cand.Count -eq 0) { throw "No .xlsm was found in the book folder. Specify it with -Book." }
    if ($cand.Count -gt 1) { throw ("There are multiple .xlsm files in the book folder: " + ($cand.Name -join ', ') + " / specify one with -Book.") }
    $Book = $cand[0].FullName
}
$Book = (Resolve-Path $Book).Path
if (-not $OutDir) { $OutDir = $repoRoot }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

$modeLabel = 'sorted into src / src\document / forms'
if ($Raw) { $modeLabel = 'Raw (flat)' }

Write-Host ("Book    : " + $Book)
Write-Host ("Output  : " + $OutDir)
Write-Host ("Mode    : " + $modeLabel)
Write-Host ""

# ---- Prepare the output location ----------------------------------------
if ($Raw) {
    $dirStd = $OutDir; $dirCls = $OutDir; $dirFrm = $OutDir; $dirDoc = $OutDir
} else {
    $dirStd = Join-Path $OutDir 'src'
    $dirCls = Join-Path $OutDir 'src'
    $dirFrm = Join-Path $OutDir 'forms'
    $dirDoc = Join-Path (Join-Path $OutDir 'src') 'document'
}
foreach ($d in @($dirStd, $dirCls, $dirFrm, $dirDoc)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

if (-not $Force) {
    $existing = @()
    foreach ($d in (@($dirStd, $dirCls, $dirFrm, $dirDoc) | Select-Object -Unique)) {
        $existing += @(Get-ChildItem $d -File -ErrorAction SilentlyContinue |
                       Where-Object { @('.bas', '.cls', '.frm', '.frx') -contains $_.Extension })
    }
    if ($existing.Count -gt 0) {
        Write-Host ("The output location already contains " + $existing.Count + " module(s). To overwrite them, re-run with -Force.") -ForegroundColor Yellow
        Write-Host ("  e.g.: " + (($existing | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ', '))
        return
    }
}

$cp932 = [System.Text.Encoding]::GetEncoding(932)
$utf8  = New-Object System.Text.UTF8Encoding($false)

# ---- Start Excel --------------------------------------------------------
$excel = $null
$wb = $null
$rows = @()
$warn = @()
$skipped = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    for ($i = 1; $i -le 10; $i++) {
        try { $wb = $excel.Workbooks.Open($Book, 0, $true); break }
        catch { Start-Sleep -Milliseconds 800 }
    }
    if ($null -eq $wb) { throw "Workbooks.Open failed 10 times. Quit Excel once and re-run." }

    $vbp = $null
    try { $vbp = $wb.VBProject } catch { }
    if ($null -eq $vbp) {
        throw "Cannot access the VBProject. In Excel, under [File]-[Options]-[Trust Center]-[Trust Center Settings]-[Macro Settings], turn ON 'Trust access to the VBA project object model'."
    }

    foreach ($comp in $vbp.VBComponents) {
        $name = $comp.Name
        $type = $comp.Type          # 1=standard 2=class 3=form 100=document
        $lines = 0
        try { $lines = $comp.CodeModule.CountOfLines } catch { }

        switch ($type) {
            1   { $ext = '.bas'; $dest = $dirStd; $kind = 'Standard module' }
            2   { $ext = '.cls'; $dest = $dirCls; $kind = 'Class module' }
            3   { $ext = '.frm'; $dest = $dirFrm; $kind = 'UserForm' }
            100 { $ext = '.cls'; $dest = $dirDoc; $kind = 'Document' }
            default { $ext = '.txt'; $dest = $dirDoc; $kind = ('Other(Type=' + $type + ')') }
        }

        if (($type -eq 3) -and (-not $Forms)) {
            $skipped += ($name + ' (UserForm. Written out only when -Forms is given)')
            continue
        }
        $path = Join-Path $dest ($name + $ext)
        $comp.Export($path)

        # --- CP932 -> UTF-8 conversion (with round-trip verification) ---
        $srcBytes  = [System.IO.File]::ReadAllBytes($path)
        $text = $cp932.GetString($srcBytes)
        $back = $cp932.GetBytes($text)
        $lossless = ($back.Length -eq $srcBytes.Length)
        if ($lossless) {
            for ($k = 0; $k -lt $srcBytes.Length; $k++) {
                if ($back[$k] -ne $srcBytes[$k]) { $lossless = $false; break }
            }
        }
        if (-not $lossless) {
            $warn += ($name + $ext + " : the round trip back to CP932 is lossy. Skipped the UTF-8 conversion and left the original as it is.")
        } else {
            [System.IO.File]::WriteAllText($path, $text, $utf8)
        }

        $size = 0
        if (Test-Path $path) { $size = (Get-Item $path).Length }

        $rows += [pscustomobject]@{
            Name   = $name
            Kind   = $kind
            Lines  = $lines
            Bytes  = $size
            Output = (Resolve-Path $path).Path.Substring($OutDir.Length).TrimStart('\')
        }
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
}

$rows | Sort-Object Kind, Name | Format-Table -AutoSize
Write-Host ("Total " + $rows.Count + " components, " + (($rows | Measure-Object Lines -Sum).Sum) + " lines")

if ($warn.Count -gt 0) {
    Write-Host ""
    Write-Host "Warnings:" -ForegroundColor Yellow
    foreach ($w in $warn) { Write-Host ("  " + $w) -ForegroundColor Yellow }
}

if ($skipped.Count -gt 0) {
    Write-Host ""
    Write-Host "Not written out:"
    foreach ($s in $skipped) { Write-Host ("  " + $s) }
}

Write-Host ""
Write-Host "Note: the workbook was opened read-only and closed without saving. The workbook is unchanged."
Write-Host "Note: .frx is binary, so it was not converted."
