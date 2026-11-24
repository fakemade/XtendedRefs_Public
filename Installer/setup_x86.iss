; XtendedRefs COM Add-in — installer for 32-bit Microsoft Office
;
; Compile:  ISCC /DBuildVersion=1.2.3.4 setup_x86.iss

#ifndef BuildVersion
  #define BuildVersion "1.0.0.0"
#endif

#define FriendlyName "XtendedRefs COM Add-in"

[Setup]
AppName={#FriendlyName} (32-bit)
AppVersion={#BuildVersion}
AppPublisher=XtendedRefs
AppId={{F1E2D3C4-B5A6-9708-8032-567890ABCDEF}
UninstallDisplayIcon={app}\XtendedRefs.ComAddIn.dll
UninstallDisplayName={#FriendlyName} (32-bit)
DefaultDirName={localappdata}\XtendedRefs.ComAddIn
OutputBaseFilename=XtendedRefsInstaller_x86
OutputDir=bin
Compression=lzma
SolidCompression=yes
VersionInfoVersion={#BuildVersion}
VersionInfoProductVersion={#BuildVersion}
VersionInfoProductName=XtendedRefs COM Add-in
VersionInfoCompany=XtendedRefs
MinVersion=6.1
PrivilegesRequired=lowest

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
Source: "bin\x86\XtendedRefs.ComAddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\x86\*.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "RegisterAddin.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "UnregisterAddin.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\RegisterAddin.ps1"" -InstallDir ""{app}"""; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\UnregisterAddin.ps1"""; Flags: runhidden waituntilterminated
