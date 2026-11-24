param(
    [Parameter(Mandatory = $true)][string]$InstallDir,
    [int]$LoadBehavior = 3
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "AddinRegistry.ps1")
. (Join-Path $PSScriptRoot "AddinConstants.ps1")

$dllPath = Join-Path $InstallDir "XtendedRefs.ComAddIn.dll"

Register-XtendedRefsAddin `
    -DllPath $dllPath `
    -ProgId $ReleaseAddin.ProgId `
    -Clsid $ReleaseAddin.Clsid `
    -FriendlyName $ReleaseAddin.FriendlyName `
    -Description $ReleaseAddin.Description `
    -LoadBehavior $LoadBehavior

Write-Host "Release add-in registered from: $dllPath"
