# Open every cross-conversion output in Microsoft Office via COM, then verify
# ONLYOFFICE's converter can load the same file into its native editor model.
param(
  [string]$Path = ".temp/cross-convert",
  [string]$OnlyOfficeConverter = "C:\Program Files\ONLYOFFICE\DesktopEditors\converter\x2t.exe",
  [string]$OnlyOfficeOutput = ".temp/onlyoffice-roundtrip",
  [switch]$SkipMicrosoft,
  [switch]$SkipOnlyOffice
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$inputDir = Join-Path $root $Path
$onlyOfficeOutputDir = Join-Path $root $OnlyOfficeOutput

if (-not (Test-Path $inputDir)) { throw "Conversion directory not found: $inputDir" }

$applications = @(
  @{ ProgId = "Word.Application"; Extensions = @(".docx", ".docm", ".dotx", ".dotm", ".rtf", ".odt") }
  @{ ProgId = "Excel.Application"; Extensions = @(".xlsx", ".xlsm", ".xltx", ".xltm", ".ods") }
  @{ ProgId = "PowerPoint.Application"; Extensions = @(".pptx", ".pptm", ".potx", ".potm", ".odp") }
)

$onlyOfficeTargets = @{
  ".doc" = ".docx"; ".docx" = ".docx"; ".docm" = ".docx"; ".dotx" = ".docx"; ".dotm" = ".docx";
  ".rtf" = ".docx"; ".odt" = ".docx";
  ".xls" = ".xlsx"; ".xlsx" = ".xlsx"; ".xlsm" = ".xlsx"; ".xltx" = ".xlsx"; ".xltm" = ".xlsx";
  ".ods" = ".xlsx";
  ".ppt" = ".pptx"; ".pptx" = ".pptx"; ".pptm" = ".pptx"; ".potx" = ".pptx"; ".potm" = ".pptx";
  ".odp" = ".pptx";
}

$results = [System.Collections.Generic.List[object]]::new()

function Add-Result([string]$file, [string]$application, [bool]$passed, [string]$detail) {
  $script:results.Add([pscustomobject]@{
    File = $file
    Application = $application
    Passed = $passed
    Detail = $detail
  })
  $state = if ($passed) { "OK" } else { "FAIL" }
  Write-Output ("{0,-4} {1,-12} {2} :: {3}" -f $state, $application, (Split-Path -Leaf $file), $detail)
}

function Open-MicrosoftFile($application, [string]$progId, [string]$path) {
  switch ($progId) {
    "Word.Application" { return $application.Documents.Open($path, $false, $true, $false) }
    "Excel.Application" { return $application.Workbooks.Open($path, 0, $true) }
    "PowerPoint.Application" { return $application.Presentations.Open($path, $true, $false, $false) }
    default { throw "Unsupported ProgId $progId" }
  }
}

function Close-MicrosoftFile($document, [string]$progId) {
  switch ($progId) {
    "Word.Application" { $document.Close(0) }
    "Excel.Application" { $document.Close($false) }
    "PowerPoint.Application" { $document.Close() }
  }
}

if (-not $SkipMicrosoft) {
  foreach ($application in $applications) {
    $files = @(Get-ChildItem -LiteralPath $inputDir -File | Where-Object {
      $application.Extensions -contains $_.Extension.ToLowerInvariant()
    } | Sort-Object Name)
    if ($files.Count -eq 0) { continue }
    $app = New-Object -ComObject $application.ProgId
    if ($application.ProgId -eq "PowerPoint.Application") {
      $app.DisplayAlerts = 1
    } else {
      $app.Visible = $false
      $app.DisplayAlerts = 0
    }
    foreach ($file in $files) {
      try {
        $document = Open-MicrosoftFile $app $application.ProgId $file.FullName
        if ($null -eq $document) { throw "returned null" }
        Close-MicrosoftFile $document $application.ProgId
        Add-Result $file.FullName $application.ProgId $true "opened and closed"
      } catch {
        Add-Result $file.FullName $application.ProgId $false $_.Exception.Message.Split("`n")[0]
      }
    }
    $app.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
  }
}

if (-not $SkipOnlyOffice) {
  if (-not (Test-Path $OnlyOfficeConverter)) { throw "ONLYOFFICE converter not found: $OnlyOfficeConverter" }
  New-Item -ItemType Directory -Force -Path $onlyOfficeOutputDir | Out-Null
  foreach ($file in (Get-ChildItem -LiteralPath $inputDir -File | Sort-Object Name)) {
    $target = $onlyOfficeTargets[$file.Extension.ToLowerInvariant()]
    if (-not $target) { continue }
    $output = Join-Path $onlyOfficeOutputDir (
      "{0}-onlyoffice{1}" -f [System.IO.Path]::GetFileNameWithoutExtension($file.Name), $target
    )
    & $OnlyOfficeConverter $file.FullName $output | Out-Null
    if ($LASTEXITCODE -eq 0 -and (Test-Path $output) -and (Get-Item $output).Length -gt 0) {
      Add-Result $file.FullName "ONLYOFFICE" $true "native load succeeded"
    } else {
      Add-Result $file.FullName "ONLYOFFICE" $false "exit $LASTEXITCODE"
    }
  }
}

$failed = @($results | Where-Object { -not $_.Passed })
Write-Output ("SUMMARY {0} checks, {1} fail" -f $results.Count, $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
