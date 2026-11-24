; XtendedRefs COM Add-in — unified installer
; Automatically detects whether Microsoft Word is installed and its bitness,
; then installs the matching DLL (x86 for 32-bit Word, x64 for 64-bit Word).
;
; Compile:  ISCC /DBuildVersion=1.2.3.4 setup.iss

#ifndef BuildVersion
  #define BuildVersion "1.0.0.0"
#endif

#define FriendlyName "XtendedRefs COM Add-in"

[Setup]
AppName={#FriendlyName}
AppVersion={#BuildVersion}
AppPublisher=XtendedRefs
AppId={{A8C3F156-2E47-4B9D-B362-DC4F87213A5E}
UninstallDisplayIcon={app}\XtendedRefs.ComAddIn.dll
UninstallDisplayName={#FriendlyName}
DefaultDirName={localappdata}\XtendedRefs.ComAddIn
OutputBaseFilename=XtendedRefsInstaller
OutputDir=bin
Compression=lzma
SolidCompression=yes
VersionInfoVersion={#BuildVersion}
VersionInfoProductVersion={#BuildVersion}
VersionInfoProductName=XtendedRefs COM Add-in
VersionInfoCompany=XtendedRefs
MinVersion=6.1
PrivilegesRequired=lowest
; Run as a 64-bit process on 64-bit Windows so we can probe both
; the native and WOW6432Node registry hives for Word.
ArchitecturesInstallIn64BitMode=x64 arm64

[Code]

{ ------------------------------------------------------------------ }
{ Helpers                                                             }
{ ------------------------------------------------------------------ }

function IsDotNet472Installed(): Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release', Release) and (Release >= 461808);
end;

{ Returns True when 64-bit Word is present (64-bit OS only).
  On 64-bit Windows all 32-bit apps redirect their HKLM\SOFTWARE writes to
  WOW6432Node. So 32-bit Word always has its InstallRoot there; 64-bit Word
  has it only in the native hive. This is more reliable than path parsing. }
function Is64BitWord(): Boolean;
var
  Dummy: String;
begin
  Result := False;
  if not Is64BitInstallMode() then Exit;

  { 32-bit Word found in WOW6432Node — definitely not 64-bit }
  if RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot', 'Path', Dummy) or
     RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Word\InstallRoot', 'Path', Dummy) or
     RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Word\InstallRoot', 'Path', Dummy) then
    Exit;

  { 64-bit Word — present only in native hive }
  Result :=
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot', 'Path', Dummy) or
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot', 'Path', Dummy) or
    RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot', 'Path', Dummy);
end;

{ Returns True when Word (any bitness) is found. }
function IsWordInstalled(): Boolean;
var
  Path: String;
begin
  { Native registry — finds 64-bit Word on 64-bit OS, any Word on 32-bit OS }
  if RegQueryStringValue(HKLM,
      'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
      '', Path) then
  begin
    Result := True;
    Exit;
  end;
  { WOW6432Node — finds 32-bit Word when installer runs in 64-bit mode }
  if Is64BitInstallMode() then
    Result := RegQueryStringValue(HKLM,
      'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
      '', Path)
  else
    Result := False;
end;

{ ------------------------------------------------------------------ }
{ Startup checks                                                      }
{ ------------------------------------------------------------------ }

function InitializeSetup(): Boolean;
begin
  if not IsWordInstalled() then
  begin
    MsgBox('Microsoft Word was not found on this system.' + #13#10 +
           'Please install Microsoft Word and run the installer again.',
           mbError, MB_OK);
    Result := False;
    Exit;
  end;

  if not IsDotNet472Installed() then
  begin
    MsgBox('.NET Framework 4.7.2 or higher is required.' + #13#10 +
           'Please install it and run the installer again.',
           mbError, MB_OK);
    Result := False;
    Exit;
  end;

  Result := True;
end;

[Files]
; 32-bit DLL — installed when Word is 32-bit (or OS is 32-bit)
Source: "bin\x86\XtendedRefs.ComAddIn.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: not Is64BitWord
Source: "bin\x86\*.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist; Check: not Is64BitWord

; 64-bit DLL — installed when Word is 64-bit
Source: "bin\x64\XtendedRefs.ComAddIn.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: Is64BitWord
Source: "bin\x64\*.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist; Check: Is64BitWord

Source: "RegisterAddin.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "UnregisterAddin.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Scripts\AddinRegistry.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Scripts\AddinConstants.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\RegisterAddin.ps1"" -InstallDir ""{app}"""; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\UnregisterAddin.ps1"""; Flags: runhidden waituntilterminated
