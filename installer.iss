[Setup]
AppName=Clear File Association
AppVersion=1.0.3
CloseApplications=force
AppPublisher=
DefaultDirName={autopf}\ClearFileAssoc
DefaultGroupName=Clear File Association
OutputBaseFilename=ClearFileAssoc-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
OutputDir=installer-output
UninstallDisplayIcon={app}\ClearFileAssoc.dll
DisableProgramGroupPage=yes
DisableDirPage=yes

[Files]
Source: "ClearFileAssoc\bin\x64\Release\ClearFileAssoc.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "ClearFileAssoc\bin\x64\Release\SharpShell.dll"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/codebase ""{app}\ClearFileAssoc.dll"""; Flags: runhidden; StatusMsg: "Registering shell extension..."

[UninstallRun]
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/unregister ""{app}\ClearFileAssoc.dll"""; Flags: runhidden; RunOnceId: "UnregisterShellExt"

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Restart Explorer after uninstall to ensure extension is unloaded
    Exec(ExpandConstant('{win}\explorer.exe'), '', '', SW_SHOW, ewNoWait, ResultCode);
  end;
end;
