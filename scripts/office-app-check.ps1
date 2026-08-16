# Open every generated demo file with its real Office application (Word / Excel /
# PowerPoint) via COM and report files the app refuses to open.
#
# XSD validation (scripts/validate.ts) proves schema conformance; this script
# catches the gap XSDs cannot see: elements that are schema-valid but rejected by
# Office's own reader (e.g. w:objectEmbed — CT_Object declares it, Word only
# accepts o:OLEObject).
#
# Prerequisites: the demo outputs must already exist under packages/<pkg>/.temp/
# (run `pnpm tsx scripts/validate.ts` first — it regenerates them), and Office
# must be installed. Run from the repo root:
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/office-app-check.ps1            # all formats
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/office-app-check.ps1 -Format docx
#
# Exit code 1 when any file fails to open.

param(
  [ValidateSet("docx", "xlsx", "pptx", "all")]
  [string]$Format = "all"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

$formats = @()
if ($Format -in @("docx", "all")) { $formats += @{ ProgId = "Word.Application";     Ext = "*.docx"; Dir = Join-Path $root "packages\docx\.temp" } }
if ($Format -in @("xlsx", "all")) { $formats += @{ ProgId = "Excel.Application";    Ext = "*.xlsx"; Dir = Join-Path $root "packages\xlsx\.temp" } }
if ($Format -in @("pptx", "all")) { $formats += @{ ProgId = "PowerPoint.Application"; Ext = "*.pptx"; Dir = Join-Path $root "packages\pptx\.temp" } }

function Open-WordFile($app, [string]$path) {
  # Documents.Open positional args — the 14th (OpenAndRepair:=$false) makes a
  # corrupt file throw instead of popping a repair dialog.
  return $app.Documents.Open($path, $false, $true, $false, "", "", $false, "", "", 0, 1252, $false, $true, $false)
}

function Open-ExcelFile($app, [string]$path) {
  # Workbooks.Open(FileName, UpdateLinks, ReadOnly) — trailing optionals are
  # omitted (their defaults amount to normal load; typing them as $null or
  # [Type]::Missing fails PowerShell 5.1's COM binding). A corrupt file throws
  # or returns null with DisplayAlerts disabled.
  return $app.Workbooks.Open($path, 0, $true)
}

function Open-PowerPointFile($app, [string]$path) {
  # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow) throws on
  # presentations it cannot load.
  return $app.Presentations.Open($path, $true, $false, $false)
}

$total = 0
$fail = 0

foreach ($fmt in $formats) {
  if (-not (Test-Path $fmt.Dir)) { Write-Output ("SKIP  " + $fmt.Ext + " :: no .temp dir (run scripts/validate.ts first)"); continue }
  $files = @(Get-ChildItem -Path $fmt.Dir -Filter $fmt.Ext | Sort-Object Name)
  if ($files.Count -eq 0) { Write-Output ("SKIP  " + $fmt.Ext + " :: no generated files"); continue }

  $app = New-Object -ComObject $fmt.ProgId
  # COM optionality quirks per app: PowerPoint's Visible is MsoTriState (bool
  # coercion throws — it starts hidden anyway) and its DisplayAlerts enum starts
  # at 1 (ppAlertsNone), unlike Word/Excel whose "none" is 0.
  if ($fmt.ProgId -eq "PowerPoint.Application") {
    $app.DisplayAlerts = 1
  } else {
    $app.Visible = $false
    $app.DisplayAlerts = 0
  }

  foreach ($file in $files) {
    $total++
    try {
      $doc = $null
      switch ($fmt.Ext) {
        "*.docx" { $doc = Open-WordFile $app $file.FullName }
        "*.xlsx" { $doc = Open-ExcelFile $app $file.FullName }
        "*.pptx" { $doc = Open-PowerPointFile $app $file.FullName }
      }
      if ($null -eq $doc) { Write-Output ("FAIL  " + $file.Name + " :: null document"); $fail++; continue }
      Write-Output ("OK    " + $file.Name)
      switch ($fmt.Ext) {
        "*.docx" { $doc.Close(0) }
        "*.xlsx" { $doc.Close($false) }
        "*.pptx" { $doc.Close() }
      }
    } catch {
      $msg = $_.Exception.Message.Split("`n")[0]
      Write-Output ("FAIL  " + $file.Name + " :: " + $msg)
      $fail++
    }
  }

  $app.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}

Write-Output ("SUMMARY $total files, $fail fail")
if ($fail -gt 0) { exit 1 }
