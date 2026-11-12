<#
.SYNOPSIS
  Injects (or repairs) an Office 2010+ ribbon (customUI14.xml) into an
  .xlsm or .xlam file, using only built-in .NET zip classes - no extra
  download needed. Safe to run more than once on the same file.

.PARAMETER FilePath
  Path to the .xlsm or .xlam file (must be closed in Excel).

.PARAMETER XmlPath
  Path to customUI14.xml (defaults to customUI14.xml next to this script).

.EXAMPLE
  .\Add-RibbonTab.ps1 -FilePath "C:\MAAO\SmartControlCenter.xlam"
#>
param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [string]$XmlPath = (Join-Path $PSScriptRoot 'customUI14.xml')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Fail($msg) {
    Write-Host ""
    Write-Host "FAILED: $msg" -ForegroundColor Red
    Write-Host ""
    exit 1
}

if (-not (Test-Path -LiteralPath $FilePath)) { Fail "File not found: $FilePath" }
if (-not (Test-Path -LiteralPath $XmlPath))  { Fail "Ribbon XML not found: $XmlPath (expected next to this script)" }

$full = (Resolve-Path -LiteralPath $FilePath).Path
Write-Host "Target file: $full"

# Make sure the file isn't open/locked (e.g. still open in Excel) before touching it.
try {
    $stream = [System.IO.File]::Open($full, 'Open', 'ReadWrite', 'None')
    $stream.Close()
} catch {
    Fail "The file appears to be open (in Excel or elsewhere). Close it completely and try again."
}

$backup = "$full.bak"
Copy-Item -LiteralPath $full -Destination $backup -Force
Write-Host "Backup saved: $backup"

$zip = $null
try {
    $zip = [System.IO.Compression.ZipFile]::Open($full, 'Update')

    # 1) customUI/customUI14.xml - always replace with the current XML
    $e = $zip.GetEntry('customUI/customUI14.xml')
    if ($e) { $e.Delete() }
    $e = $zip.CreateEntry('customUI/customUI14.xml')
    $sw = New-Object System.IO.StreamWriter($e.Open())
    $sw.Write([System.IO.File]::ReadAllText($XmlPath))
    $sw.Close()
    Write-Host "customUI/customUI14.xml written."

    # 2) [Content_Types].xml - strip any previous entry, add a fresh one (self-healing)
    $ctEntry = $zip.GetEntry('[Content_Types].xml')
    if (-not $ctEntry) { Fail "'[Content_Types].xml' missing - is this really a valid .xlsm/.xlam?" }
    $sr = New-Object System.IO.StreamReader($ctEntry.Open())
    $ct = $sr.ReadToEnd()
    $sr.Close()
    $ct = [regex]::Replace($ct, '<Override PartName="/customUI/customUI14\.xml"[^>]*/>', '')
    $ct = $ct -replace '</Types>', '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/></Types>'
    $ctEntry.Delete()
    $ne = $zip.CreateEntry('[Content_Types].xml')
    $sw2 = New-Object System.IO.StreamWriter($ne.Open())
    $sw2.Write($ct)
    $sw2.Close()
    Write-Host "[Content_Types].xml updated."

    # 3) _rels/.rels - strip any previous ribbon relationship, add a fresh one
    $relsEntry = $zip.GetEntry('_rels/.rels')
    if (-not $relsEntry) { Fail "'_rels/.rels' missing - is this really a valid .xlsm/.xlam?" }
    $sr2 = New-Object System.IO.StreamReader($relsEntry.Open())
    $rels = $sr2.ReadToEnd()
    $sr2.Close()
    $rels = [regex]::Replace($rels, '<Relationship[^>]*ui/extensibility[^>]*/>', '')
    $relXml = '<Relationship Id="rIdSCCRibbon" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>'
    $rels = $rels -replace '</Relationships>', $relXml
    $relsEntry.Delete()
    $ne2 = $zip.CreateEntry('_rels/.rels')
    $sw3 = New-Object System.IO.StreamWriter($ne2.Open())
    $sw3.Write($rels)
    $sw3.Close()
    Write-Host "_rels/.rels updated."
}
catch {
    if ($zip) { $zip.Dispose() }
    Copy-Item -LiteralPath $backup -Destination $full -Force
    Fail "Something went wrong, original file restored from backup. Details: $($_.Exception.Message)"
}
finally {
    if ($zip) { $zip.Dispose() }
}

# Verify: reopen read-only and confirm the three parts are really there.
$ok = $true
$verify = [System.IO.Compression.ZipFile]::OpenRead($full)
try {
    if (-not $verify.GetEntry('customUI/customUI14.xml')) { $ok = $false }
    $ctCheck = (New-Object System.IO.StreamReader($verify.GetEntry('[Content_Types].xml').Open())).ReadToEnd()
    if ($ctCheck -notmatch 'customUI14\.xml') { $ok = $false }
    $relCheck = (New-Object System.IO.StreamReader($verify.GetEntry('_rels/.rels').Open())).ReadToEnd()
    if ($relCheck -notmatch 'ui/extensibility') { $ok = $false }
} finally {
    $verify.Dispose()
}

Write-Host ""
if ($ok) {
    Write-Host "SUCCESS - ribbon tab installed into $full" -ForegroundColor Green
} else {
    Fail "The write finished but verification did not find all three parts. Restore $backup and try again."
}
Write-Host ""
Write-Host "Next: open the file in Excel." -ForegroundColor Yellow
Write-Host "If Excel shows 'we found a problem with some content' - click Yes/Repair once;" -ForegroundColor Yellow
Write-Host "Excel will drop only the broken part and keep your workbook and VBA project intact." -ForegroundColor Yellow
Write-Host "If the file was downloaded from the internet, right-click it > Properties > Unblock first." -ForegroundColor Yellow
