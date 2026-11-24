$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "AddinRegistry.ps1")
. (Join-Path $PSScriptRoot "AddinConstants.ps1")

Unregister-XtendedRefsAddin `
    -ProgId $ReleaseAddin.ProgId `
    -Clsid $ReleaseAddin.Clsid

Write-Host "Release add-in registry entries removed."
