; XtendedRefs COM Add-in — installer for 64-bit Microsoft Office
;
; Compile:  ISCC /DBuildVersion=1.2.3.4 setup_x64.iss

#ifndef BuildVersion
  #define BuildVersion "1.0.0.0"
#endif

#define FriendlyName "XtendedRefs COM Add-in"

[Setup]
AppName={#FriendlyName} (64-bit)
AppVersion={#BuildVersion}
AppPublisher=XtendedRefs
AppId={{B2C3D4E5-F6A7-8901-8064-BCDEF1234567}
UninstallDisplayIcon={app}\XtendedRefs.ComAddIn.dll
UninstallDisplayName={#FriendlyName} (64-bit)
DefaultDirName={localappdata}\XtendedRefs.ComAddIn
OutputBaseFilename=XtendedRefsInstaller_x64
OutputDir=bin
Compression=lzma
SolidCompression=yes
VersionInfoVersion={#BuildVersion}
VersionInfoProductVersion={#BuildVersion}
VersionInfoProductName=XtendedRefs COM Add-in
VersionInfoCompany=XtendedRefs
MinVersion=6.1
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64 arm64

[Code]
function IsDotNet472Installed(): Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release', Release) and (Release >= 461808);
end;

function InitializeSetup(): Boolean;
begin
  if not IsWin64() then begin
    MsgBox(
      'This installer is for 64-bit Office on 64-bit Windows.' + #13#10 +
      'For 32-bit Office use XtendedRefsInstaller_x86.exe.',
      mbError, MB_OK);
    Result := False;
    Exit;
  end;

  if not IsDotNet472Installed() then begin
    MsgBox(
      '.NET Framework 4.7.2 or higher is required.' + #13#10 +
      'Please install it and run the installer again.',
      mbError, MB_OK);
    Result := False;
  end else
    Result := True;
end;

[Files]
Source: "bin\x64\XtendedRefs.ComAddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\x64\*.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "RegisterAddin.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "UnregisterAddin.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\RegisterAddin.ps1"" -InstallDir ""{app}"""; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\UnregisterAddin.ps1"""; Flags: runhidden waituntilterminated
