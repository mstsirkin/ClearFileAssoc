[Setup]
AppName=Clear File Association
AppVersion=1.0.2
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
procedure RestartExplorer(Kill: Boolean);
var
  ResultCode: Integer;
begin
  if Kill then
  begin
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/f /im explorer.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Sleep(1000);
  end
  else
  begin
    Exec(ExpandConstant('{win}\explorer.exe'), '', '', SW_SHOW, ewNoWait, ResultCode);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Restart Explorer to load the new extension
    RestartExplorer(True);
    RestartExplorer(False);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    // Kill Explorer BEFORE unregistering and deleting files
    RestartExplorer(True);
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    // Restart Explorer after uninstall is complete
    RestartExplorer(False);
  end;
end;
