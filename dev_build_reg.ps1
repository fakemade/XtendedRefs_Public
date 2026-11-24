param(
    [string]$Configuration = "Debug",
    [string]$Version = "1.0.0.0",
    [switch]$Activate
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
if (-not $root) { $root = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $root) { $root = (Get-Location).Path }
Set-Location $root

. (Join-Path $root "Scripts\AddinRegistry.ps1")
. (Join-Path $root "Scripts\AddinConstants.ps1")

$projectPath = Join-Path $root "XtendedRefs.ComAddIn.csproj"
$outputRoot = Join-Path $root "bin\$Configuration"

Write-Host "Project root: $root"
Write-Host "Configuration: $Configuration"
Write-Host "Version: $Version"

Invoke-XtendedRefsBuild `
    -ProjectPath $projectPath `
    -Configuration $Configuration `
    -Version $Version `
    -OutputRoot $outputRoot

$wordBitness = Get-XtendedRefsWordBitness
if (-not $wordBitness) {
    Write-Warning "Microsoft Word not found. Skipping local COM registration (CI/build-only mode)."
    exit 0
}

$dllPath = Join-Path $outputRoot "$wordBitness\XtendedRefs.ComAddIn.dll"

Register-XtendedRefsAddin `
    -DllPath $dllPath `
    -ProgId $DevAddin.ProgId `
    -Clsid $DevAddin.Clsid `
    -FriendlyName $DevAddin.FriendlyName `
    -Description $DevAddin.Description `
    -LoadBehavior ($(if ($Activate) { 3 } else { 2 }))

if ($Activate) {
    Set-XtendedRefsAddinLoadBehavior -ProgId $ReleaseAddin.ProgId -LoadBehavior 2
    Write-Host "Dev add-in activated. Release add-in disabled for next Word start."
}
else {
    Write-Host "Dev add-in registered in passive mode (LoadBehavior = 2)."
}

Write-Host "Word bitness: $wordBitness"
Write-Host "Registered DLL: $dllPath"
Write-Host "Dev LoadBehavior: $(Get-XtendedRefsAddinLoadBehavior -ProgId $DevAddin.ProgId)"
Write-Host "Release LoadBehavior: $(Get-XtendedRefsAddinLoadBehavior -ProgId $ReleaseAddin.ProgId)"
