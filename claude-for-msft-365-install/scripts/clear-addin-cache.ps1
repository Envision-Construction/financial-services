<#
.SYNOPSIS
  Clear a single Office add-in's cached / sideloaded manifest on Windows.

.DESCRIPTION
  The Wef cache holds every add-in side by side, each file named
  <addin-id>.manifest-*.xml. This removes ONLY the files matching one
  add-in ID -- it never wipes the Wef folder.

.EXAMPLE
  clear-addin-cache.ps1                          # list every add-in found, do nothing
  clear-addin-cache.ps1 -Manifest C:\m.xml       # dry-run: show what would be removed
  clear-addin-cache.ps1 -Id <GUID>               # dry-run by ID (no manifest needed)
  clear-addin-cache.ps1 -Manifest C:\m.xml -Apply  # actually delete
#>
[CmdletBinding()]
param(
  [string]$Manifest,
  [string]$Id,
  [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'

if (-not (Test-Path $wef)) { Write-Host "No Wef folder at $wef -- nothing cached."; return }

# No args -> just list what's cached, then exit.
if (-not $Manifest -and -not $Id) {
  Write-Host "Add-ins currently cached in wef (id  <-  filename):"
  Get-ChildItem $wef -Recurse -Filter *.xml -ErrorAction SilentlyContinue | ForEach-Object {
    $guid = ($_.Name -split '\.')[0]
    "  {0}  <-  {1}" -f $guid, $_.Name
  }
  Write-Host "`nRe-run with -Manifest <path> or -Id <GUID> to clear one (add -Apply to delete)."
  return
}

if (-not $Id) {
  if (-not (Test-Path $Manifest)) { throw "manifest not found: $Manifest" }
  $Id = ([xml](Get-Content $Manifest)).OfficeApp.Id
}
if (-not $Id) { throw "could not determine add-in ID" }

$matches = Get-ChildItem $wef -Recurse -Filter "$Id*" -ErrorAction SilentlyContinue

if ($Apply) { Write-Host "Removing cached/sideloaded manifests for add-in $Id" }
else        { Write-Host "DRY RUN -- would remove these (re-run with -Apply to delete):" }

if (-not $matches) {
  Write-Host "  (nothing found for $Id -- already clear)"
} else {
  foreach ($f in $matches) {
    if ($Apply) { Remove-Item $f.FullName -Force; Write-Host "  removed $($f.FullName)" }
    else        { Write-Host "  would remove $($f.FullName)" }
  }
}
Write-Host "Quit and reopen the Office apps so they re-fetch the manifest."
