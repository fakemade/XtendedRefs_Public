Set-StrictMode -Version Latest

function Get-XtendedRefsMsBuildPath {
    $msbuild = Get-ChildItem "C:\Program Files\Microsoft Visual Studio" -Recurse -Filter "MSBuild.exe" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notlike "*arm64*" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $msbuild) {
        throw "MSBuild.exe not found under Visual Studio installation."
    }

    return $msbuild
}

function Get-XtendedRefsWordBitness {
    if (-not [Environment]::Is64BitOperatingSystem) {
        return "x86"
    }

    $ctrPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $ctrPath) {
        try {
            $platform = (Get-ItemProperty $ctrPath).Platform
            if ($platform -eq "x64") { return "x64" }
            if ($platform -eq "x86") { return "x86" }
        }
        catch {
        }
    }

    $versions = @("16.0", "15.0", "14.0")

    foreach ($ver in $versions) {
        $wowPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$ver\Word\InstallRoot"
        if (Test-Path $wowPath) {
            return "x86"
        }
    }

    foreach ($ver in $versions) {
        $nativePath = "HKLM:\SOFTWARE\Microsoft\Office\$ver\Word\InstallRoot"
        if (Test-Path $nativePath) {
            return "x64"
        }
    }

    return $null
}

function Invoke-XtendedRefsBuild {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$Configuration,
        [Parameter(Mandatory = $true)][string]$Version,
        [Parameter(Mandatory = $true)][string]$OutputRoot
    )

    $msbuild = Get-XtendedRefsMsBuildPath
    Write-Host "Using MSBuild: $msbuild"

    $targets = @(
        @{ Platform = "x86"; OutputPath = (Join-Path $OutputRoot "x86") }
        @{ Platform = "x64"; OutputPath = (Join-Path $OutputRoot "x64") }
    )

    foreach ($target in $targets) {
        Write-Host "Building $($target.Platform) ($Configuration) -> $($target.OutputPath)"

        & $msbuild $ProjectPath `
            /p:Configuration=$Configuration `
            /p:PlatformTarget=$($target.Platform) `
            /p:OutputPath=$($target.OutputPath) `
            /p:AssemblyVersion=$Version `
            /p:FileVersion=$Version `
            /t:Rebuild

        if ($LASTEXITCODE -ne 0) {
            throw "$($target.Platform) build failed."
        }
    }
}

function Register-XtendedRefsAddin {
    param(
        [Parameter(Mandatory = $true)][string]$DllPath,
        [Parameter(Mandatory = $true)][string]$ProgId,
        [Parameter(Mandatory = $true)][string]$Clsid,
        [Parameter(Mandatory = $true)][string]$FriendlyName,
        [Parameter(Mandatory = $true)][string]$Description,
        [int]$LoadBehavior = 3,
        [string]$ClassName = "XtendedRefs.ComAddIn.Connect",
        [string]$RuntimeVersion = "v4.0.30319"
    )

    if (-not (Test-Path $DllPath)) {
        throw "DLL not found: $DllPath"
    }

    $resolvedDllPath = (Resolve-Path $DllPath).Path
    $assemblyVersion = [System.Reflection.AssemblyName]::GetAssemblyName($resolvedDllPath).Version
    $assemblyString = "XtendedRefs.ComAddIn, Version=$assemblyVersion, Culture=neutral, PublicKeyToken=null"
    $codeBase = "file:///" + $resolvedDllPath.Replace("\", "/")

    function Register-InView {
        param([Microsoft.Win32.RegistryView]$View)

        $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $View)

        try {
            $progIdKey = $hive.CreateSubKey("SOFTWARE\Classes\$ProgId")
            $progIdKey.SetValue("", $ClassName)

            $progIdClsidKey = $hive.CreateSubKey("SOFTWARE\Classes\$ProgId\CLSID")
            $progIdClsidKey.SetValue("", $Clsid)

            $clsidKey = $hive.CreateSubKey("SOFTWARE\Classes\CLSID\$Clsid")
            $clsidKey.SetValue("", $ClassName)

            $clsidProgIdKey = $hive.CreateSubKey("SOFTWARE\Classes\CLSID\$Clsid\ProgId")
            $clsidProgIdKey.SetValue("", $ProgId)

            $inprocKey = $hive.CreateSubKey("SOFTWARE\Classes\CLSID\$Clsid\InprocServer32")
            $inprocKey.SetValue("", "mscoree.dll")
            $inprocKey.SetValue("Class", $ClassName)
            $inprocKey.SetValue("Assembly", $assemblyString)
            $inprocKey.SetValue("RuntimeVersion", $RuntimeVersion)
            $inprocKey.SetValue("ThreadingModel", "Both")
            $inprocKey.SetValue("CodeBase", $codeBase)

            $versionKey = $hive.CreateSubKey("SOFTWARE\Classes\CLSID\$Clsid\InprocServer32\$assemblyVersion")
            $versionKey.SetValue("Class", $ClassName)
            $versionKey.SetValue("Assembly", $assemblyString)
            $versionKey.SetValue("RuntimeVersion", $RuntimeVersion)
            $versionKey.SetValue("CodeBase", $codeBase)

            $hive.CreateSubKey("SOFTWARE\Classes\CLSID\$Clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}") | Out-Null

            $wordAddinKey = $hive.CreateSubKey("Software\Microsoft\Office\Word\Addins\$ProgId")
            $wordAddinKey.SetValue("FriendlyName", $FriendlyName)
            $wordAddinKey.SetValue("Description", $Description)
            $wordAddinKey.SetValue("LoadBehavior", $LoadBehavior, [Microsoft.Win32.RegistryValueKind]::DWord)
            $wordAddinKey.SetValue("CommandLineSafe", 0, [Microsoft.Win32.RegistryValueKind]::DWord)
        }
        finally {
            $hive.Close()
        }
    }

    Register-InView ([Microsoft.Win32.RegistryView]::Registry32)
    Register-InView ([Microsoft.Win32.RegistryView]::Registry64)
}

function Unregister-XtendedRefsAddin {
    param(
        [Parameter(Mandatory = $true)][string]$ProgId,
        [Parameter(Mandatory = $true)][string]$Clsid
    )

    function Remove-KeyIfExists {
        param(
            [Microsoft.Win32.RegistryKey]$Hive,
            [string]$Path
        )

        try {
            if ($Hive.OpenSubKey($Path) -ne $null) {
                $Hive.DeleteSubKeyTree($Path, $false)
                Write-Host "Removed: $Path"
            }
        }
        catch {
            Write-Warning "Failed to remove '$Path': $_"
        }
    }

    function Unregister-InView {
        param([Microsoft.Win32.RegistryView]$View)

        $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $View)

        try {
            Remove-KeyIfExists $hive "SOFTWARE\Classes\$ProgId"
            Remove-KeyIfExists $hive "SOFTWARE\Classes\CLSID\$Clsid"
            Remove-KeyIfExists $hive "Software\Microsoft\Office\Word\Addins\$ProgId"
        }
        finally {
            $hive.Close()
        }
    }

    Unregister-InView ([Microsoft.Win32.RegistryView]::Registry32)
    Unregister-InView ([Microsoft.Win32.RegistryView]::Registry64)
}

function Set-XtendedRefsAddinLoadBehavior {
    param(
        [Parameter(Mandatory = $true)][string]$ProgId,
        [Parameter(Mandatory = $true)][int]$LoadBehavior
    )

    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $view)

        try {
            $path = "Software\Microsoft\Office\Word\Addins\$ProgId"
            $key = $hive.OpenSubKey($path, $true)
            if ($key -ne $null) {
                $key.SetValue("LoadBehavior", $LoadBehavior, [Microsoft.Win32.RegistryValueKind]::DWord)
            }
        }
        finally {
            $hive.Close()
        }
    }
}

function Get-XtendedRefsAddinLoadBehavior {
    param([Parameter(Mandatory = $true)][string]$ProgId)

    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        [Microsoft.Win32.RegistryView]::Registry32)

    try {
        $key = $hive.OpenSubKey("Software\Microsoft\Office\Word\Addins\$ProgId")
        if ($key -eq $null) {
            return $null
        }

        return $key.GetValue("LoadBehavior")
    }
    finally {
        $hive.Close()
    }
}
