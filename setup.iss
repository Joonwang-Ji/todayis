[Setup]
AppName=오늘은
AppVersion=1.2
DefaultDirName={pf}\오늘은
DefaultGroupName=오늘은
OutputDir=Output
OutputBaseFilename=TodayIsSetup_v1.2
SetupIconFile=C:\Users\VC2F PC1\PycharmProjects\todayis_bell\icon.ico
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Files]
Source: "dist\class_bell_app\class_bell_app.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\class_bell_app\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\오늘은"; Filename: "{app}\class_bell_app.exe"
Name: "{userdesktop}\오늘은"; Filename: "{app}\class_bell_app.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Run]
Filename: "powershell.exe"; Parameters: "-Command ""Add-MpPreference -ExclusionPath '{app}'"""; Flags: runhidden waituntilterminated; Description: "Add Defender exclusion"
Filename: "{app}\class_bell_app.exe"; Description: "{cm:LaunchProgram,오늘은}"; Flags: nowait postinstall skipifsilent runasoriginaluser

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    Exec('powershell.exe', '-Command "Add-MpPreference -ExclusionPath \"' + ExpandConstant('{app}') + '\""',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode <> 0 then
      MsgBox('Defender 제외 항목 추가 실패: ' + IntToStr(ResultCode), mbError, MB_OK);
  end;
end;