; MD Doc Link - installer
; Author: MD (Mayank Dholariya)

#define AppName "MD Doc Link"
#define AppVer "1.0.0"
#define Author "MD (Mayank Dholariya)"
#define PluginName "MDDocLink"
#define MinYear 2017
#define MaxYear 2027
#define Y

[Setup]
AppId={{981AB0F0-4CDD-41AC-9DD6-ADFE25628E20}
AppName={#AppName}
AppVersion={#AppVer}
AppVerName={#AppName} v{#AppVer}
AppPublisher={#Author}
AppCopyright=Copyright (c) 2026 {#Author}
AppComments=Link documents (PDF, folders, URLs) to Navisworks model items directly from an Excel list. Navisworks Manage / Simulate {#MinYear}-{#MaxYear}. Author: {#Author}
AppContact={#Author}
VersionInfoVersion=1.0.0.0
VersionInfoCompany={#Author}
VersionInfoCopyright=Copyright (c) 2026 {#Author}
VersionInfoDescription={#AppName} Setup - {#Author}
VersionInfoProductName={#AppName}
VersionInfoProductVersion=1.0.0.0
DefaultDirName={commonpf64}\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableWelcomePage=no
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=6.1sp1
OutputDir=..\release
OutputBaseFilename=MD_Navisworks_DocLink_Setup_v{#AppVer}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
InfoBeforeFile=before.txt
UninstallDisplayName={#AppName} v{#AppVer} - {#Author}
UninstallDisplayIcon={app}\MDDocLink.ico
SetupIconFile=..\icon\MDDocLink.ico
WizardImageFile=..\icon\wizard_large.bmp
WizardSmallImageFile=..\icon\wizard_small.bmp
SetupMutex=MDDocLinkSetupMutex
CloseApplications=no
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%nAuthor: {#Author}%n%nPlease close Navisworks before continuing.

[Files]
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "..\icon\MDDocLink.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\icon\MDDocLink_16.png"; DestDir: "{app}\Images"; Flags: ignoreversion
Source: "..\icon\MDDocLink_32.png"; DestDir: "{app}\Images"; Flags: ignoreversion
#sub AddYear
Source: "..\build\{#Y}\{#PluginName}.dll"; DestDir: "{app}\bin\{#Y}"; Flags: ignoreversion
#endsub
#for {Y = MinYear; Y <= MaxYear; Y++} AddYear

[Code]
const
  UninstKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';
  ListFile = 'installed-locations.txt';

var
  NwPath: array of String;
  NwEdition: array of String;
  NwYear: array of Integer;
  Unsupported: String;
  TargetPage: TInputOptionWizardPage;

function IsNavisworksRunning(): Boolean;
var
  Locator, Service, Procs: Variant;
begin
  Result := False;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('.', 'root\CIMV2');
    Procs := Service.ExecQuery('SELECT ProcessId FROM Win32_Process WHERE Name = ''Roamer.exe''');
    Result := Procs.Count > 0;
  except
    Result := False;
  end;
end;

function RunningSilently(): Boolean;
begin
  if IsUninstaller() then
    Result := UninstallSilent()
  else
    Result := WizardSilent();
end;

function WaitForNavisworksClosed(): Boolean;
begin
  Result := True;
  while IsNavisworksRunning() do
  begin
    if RunningSilently() then
    begin
      Result := False;
      Exit;
    end;
    Log('Navisworks (Roamer.exe) is running.');
    if MsgBox('Navisworks is running.' #13#10#13#10 +
              'Please close every Navisworks window, then click Retry.',
              mbError, MB_RETRYCANCEL) <> IDRETRY then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

function YearFromApi(const Dir: String): Integer;
var
  MS, LS: Cardinal;
  S: String;
begin
  Result := 0;
  if GetVersionNumbers(AddBackslash(Dir) + 'Autodesk.Navisworks.Api.dll', MS, LS) then
    Result := Integer(MS shr 16) + 2003;
  if (Result < 2010) or (Result > 2100) then
  begin
    S := RemoveBackslashUnlessRoot(Dir);
    Result := StrToIntDef(Copy(S, Length(S) - 3, 4), 0);
  end;
end;

procedure AddInstall(Dir, Hint: String);
var
  I, N, Year: Integer;
  Ed: String;
begin
  Dir := RemoveBackslashUnlessRoot(Trim(RemoveQuotes(Dir)));
  if Dir = '' then Exit;
  if not FileExists(AddBackslash(Dir) + 'Roamer.exe') then Exit;
  if not FileExists(AddBackslash(Dir) + 'Autodesk.Navisworks.Api.dll') then Exit;
  if Pos('freedom', Lowercase(Dir + ' ' + Hint)) > 0 then Exit;
  for I := 0 to GetArrayLength(NwPath) - 1 do
    if CompareText(NwPath[I], Dir) = 0 then Exit;

  if Pos('simulate', Lowercase(Dir + ' ' + Hint)) > 0 then
    Ed := 'Simulate'
  else
    Ed := 'Manage';
  Year := YearFromApi(Dir);

  if (Year < {#MinYear}) or (Year > {#MaxYear}) then
  begin
    Unsupported := Unsupported + '  Navisworks ' + Ed + ' ' + IntToStr(Year) + '  (' + Dir + ')' + #13#10;
    Exit;
  end;

  N := GetArrayLength(NwPath);
  SetArrayLength(NwPath, N + 1);
  SetArrayLength(NwEdition, N + 1);
  SetArrayLength(NwYear, N + 1);
  NwPath[N] := Dir;
  NwEdition[N] := Ed;
  NwYear[N] := Year;
end;

procedure ScanUninstallKeys(RootKey: Integer);
var
  Names: TArrayOfString;
  I: Integer;
  Key, DisplayName, Location: String;
begin
  if not RegGetSubkeyNames(RootKey, UninstKey, Names) then Exit;
  for I := 0 to GetArrayLength(Names) - 1 do
  begin
    Key := UninstKey + '\' + Names[I];
    if RegQueryStringValue(RootKey, Key, 'DisplayName', DisplayName) then
      if (Pos('navisworks manage', Lowercase(DisplayName)) > 0) or
         (Pos('navisworks simulate', Lowercase(DisplayName)) > 0) then
        if RegQueryStringValue(RootKey, Key, 'InstallLocation', Location) then
          AddInstall(Location, DisplayName);
  end;
end;

procedure DetectNavisworks();
var
  Year: Integer;
  Base: String;
begin
  SetArrayLength(NwPath, 0);
  SetArrayLength(NwEdition, 0);
  SetArrayLength(NwYear, 0);
  Unsupported := '';
  Base := ExpandConstant('{commonpf64}') + '\Autodesk\';
  for Year := 2014 to 2040 do
  begin
    AddInstall(Base + 'Navisworks Manage ' + IntToStr(Year), 'Manage');
    AddInstall(Base + 'Navisworks Simulate ' + IntToStr(Year), 'Simulate');
  end;
  ScanUninstallKeys(HKLM64);
  ScanUninstallKeys(HKLM32);
end;

function InitializeSetup(): Boolean;
var
  Msg: String;
begin
  Result := False;
  DetectNavisworks();
  if GetArrayLength(NwPath) = 0 then
  begin
    Msg := 'No supported Navisworks was found on this PC.' #13#10#13#10 +
           'This plugin needs Autodesk Navisworks Manage or Navisworks Simulate, ' +
           '{#MinYear} to {#MaxYear} (64-bit).';
    if Unsupported <> '' then
      Msg := Msg + #13#10#13#10 + 'Found, but not supported by this version:' + #13#10 + Unsupported;
    Msg := Msg + #13#10#13#10 + 'Setup will now close. Nothing was installed.';
    SuppressibleMsgBox(Msg, mbError, MB_OK, IDOK);
    Exit;
  end;
  if not WaitForNavisworksClosed() then Exit;
  Result := True;
end;

procedure InitializeWizard();
var
  I: Integer;
begin
  TargetPage := CreateInputOptionPage(wpInfoBefore,
    'Choose Navisworks versions',
    'Setup found these Navisworks installations.',
    'Tick each Navisworks you want MD Doc Link added to, then click Next.',
    False, False);
  for I := 0 to GetArrayLength(NwPath) - 1 do
  begin
    TargetPage.Add('Navisworks ' + NwEdition[I] + ' ' + IntToStr(NwYear[I]) + '   -   ' + NwPath[I]);
    TargetPage.Values[I] := True;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  I: Integer;
  Any: Boolean;
begin
  Result := True;
  if CurPageID = TargetPage.ID then
  begin
    Any := False;
    for I := 0 to GetArrayLength(NwPath) - 1 do
      if TargetPage.Values[I] then Any := True;
    if not Any then
    begin
      MsgBox('Tick at least one Navisworks version.', mbError, MB_OK);
      Result := False;
    end;
  end;
  if CurPageID = wpReady then
    Result := WaitForNavisworksClosed();
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  I: Integer;
begin
  Result := 'MD Doc Link will be added to:' + NewLine;
  for I := 0 to GetArrayLength(NwPath) - 1 do
    if TargetPage.Values[I] then
      Result := Result + Space + 'Navisworks ' + NwEdition[I] + ' ' + IntToStr(NwYear[I]) + NewLine +
                Space + Space + NwPath[I] + '\Plugins\{#PluginName}\' + NewLine;
  if Unsupported <> '' then
    Result := Result + NewLine + 'Skipped (not supported by this version):' + NewLine + Unsupported;
  Result := Result + NewLine + 'Author: {#Author}';
end;

procedure AddUnique(var List: TArrayOfString; const S: String);
var
  I, N: Integer;
begin
  for I := 0 to GetArrayLength(List) - 1 do
    if CompareText(List[I], S) = 0 then Exit;
  N := GetArrayLength(List);
  SetArrayLength(List, N + 1);
  List[N] := S;
end;

procedure InstallPlugins();
var
  I: Integer;
  Src, DestDir, Failed, Done: String;
  Recorded: TArrayOfString;
begin
  Failed := '';
  Done := '';
  if not LoadStringsFromFile(ExpandConstant('{app}\' + ListFile), Recorded) then
    SetArrayLength(Recorded, 0);

  for I := 0 to GetArrayLength(NwPath) - 1 do
  begin
    if WizardSilent or TargetPage.Values[I] then
    begin
      Src := ExpandConstant('{app}\bin\') + IntToStr(NwYear[I]) + '\{#PluginName}.dll';
      DestDir := NwPath[I] + '\Plugins\{#PluginName}';
      if ForceDirectories(DestDir + '\Images') and CopyFile(Src, DestDir + '\{#PluginName}.dll', False) and
         CopyFile(ExpandConstant('{app}\Images\MDDocLink_16.png'), DestDir + '\Images\MDDocLink_16.png', False) and
         CopyFile(ExpandConstant('{app}\Images\MDDocLink_32.png'), DestDir + '\Images\MDDocLink_32.png', False) then
      begin
        AddUnique(Recorded, DestDir);
        Done := Done + '  Navisworks ' + NwEdition[I] + ' ' + IntToStr(NwYear[I]) + #13#10;
        Log('Installed: ' + DestDir);
      end
      else
      begin
        Failed := Failed + '  ' + DestDir + #13#10;
        Log('FAILED: ' + DestDir);
      end;

    end;
  end;

  SaveStringsToFile(ExpandConstant('{app}\' + ListFile), Recorded, False);

  if Failed <> '' then
    SuppressibleMsgBox('Could not install into:' #13#10 + Failed + #13#10 +
      'Make sure Navisworks is closed and run Setup again.', mbError, MB_OK, IDOK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    InstallPlugins();
end;

function InitializeUninstall(): Boolean;
begin
  Result := WaitForNavisworksClosed();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Recorded: TArrayOfString;
  I: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    if LoadStringsFromFile(ExpandConstant('{app}\' + ListFile), Recorded) then
      for I := 0 to GetArrayLength(Recorded) - 1 do
      begin
        if Trim(Recorded[I]) <> '' then
        begin
          DeleteFile(AddBackslash(Trim(Recorded[I])) + '{#PluginName}.dll');
          DeleteFile(AddBackslash(Trim(Recorded[I])) + 'Images\MDDocLink_16.png');
          DeleteFile(AddBackslash(Trim(Recorded[I])) + 'Images\MDDocLink_32.png');
          RemoveDir(AddBackslash(Trim(Recorded[I])) + 'Images');
          RemoveDir(Trim(Recorded[I]));
          Log('Removed: ' + Recorded[I]);
        end;
      end;
    DeleteFile(ExpandConstant('{app}\' + ListFile));
  end;
end;
