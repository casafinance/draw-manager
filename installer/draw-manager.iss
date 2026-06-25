; ============================================================================
; Draw Manager - Inno Setup installer
;
; Per-user install (no admin). Bundles both exes + icon. Prompts the installing
; user to select their google-credentials.json (REQUIRED to finish), copies it
; into the install folder, and writes a minimal settings.json pointing at it.
; The app fills in all other settings defaults on first launch.
;
; MyAppVersion is passed in from the build command:
;   iscc /DMyAppVersion=0.1.0 installer\draw-manager.iss
; ============================================================================

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "Draw Manager"
#define MyMainExe "Draw Manager.exe"
#define MyWorkerExe "draw-request.exe"

[Setup]
AppId={{B7E6B0A2-3C4D-4E5F-9A1B-DR4WM4N4G3R01}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Casa Finance
DefaultDirName={localappdata}\Programs\Draw Manager
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=yes
; Per-user install: no admin elevation required.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=installer_output
OutputBaseFilename=Draw-Manager-Setup
SetupIconFile=..\draw_manager.ico
UninstallDisplayIcon={app}\{#MyMainExe}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"

[Files]
; The two exes and icon are copied from the build output (dist) at build time.
Source: "..\dist\{#MyMainExe}";   DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\{#MyWorkerExe}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\casa-updater.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\draw_manager.ico";    DestDir: "{app}"; Flags: ignoreversion
; The user-selected credentials file is copied at runtime in [Code] (see
; CurStepChanged), because its source path is chosen on the wizard page.

[Icons]
Name: "{group}\{#MyAppName}";        Filename: "{app}\{#MyMainExe}"; IconFilename: "{app}\draw_manager.ico"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}";  Filename: "{app}\{#MyMainExe}"; IconFilename: "{app}\draw_manager.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyMainExe}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
var
  CredPage: TInputFileWizardPage;

procedure InitializeWizard;
begin
  { Custom page asking for the Google credentials JSON. Required. }
  CredPage := CreateInputFilePage(wpSelectDir,
    'Google Credentials',
    'Select your Google service-account key file.',
    'Draw Manager needs your google-credentials.json file to connect to Google Sheets. ' +
    'This file was provided to you separately by your administrator. ' +
    'Click Browse and select it, then click Next.');
  CredPage.Add('Credentials file (google-credentials.json):',
    'JSON files (*.json)|*.json|All files (*.*)|*.*', '.json');
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  P: String;
begin
  Result := True;
  if CurPageID = CredPage.ID then
  begin
    P := Trim(CredPage.Values[0]);
    if P = '' then
    begin
      MsgBox('You must select your google-credentials.json file to continue.' + #13#10 +
             'If you do not have it, contact your administrator.', mbError, MB_OK);
      Result := False;
      Exit;
    end;
    if not FileExists(P) then
    begin
      MsgBox('That file does not exist. Please select a valid google-credentials.json file.',
             mbError, MB_OK);
      Result := False;
      Exit;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  SrcCred, DestCred, SettingsPath, Json, EscPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    { Copy the chosen credentials file into the install folder. }
    SrcCred  := Trim(CredPage.Values[0]);
    DestCred := ExpandConstant('{app}\google-credentials.json');
    if SrcCred <> '' then
      FileCopy(SrcCred, DestCred, False);

    { Write a minimal settings.json pointing at the copied key. The app
      deep-merges defaults for everything else on first launch. JSON needs
      backslashes escaped, so we double them. }
    EscPath := DestCred;
    StringChangeEx(EscPath, '\', '\\', True);
    SettingsPath := ExpandConstant('{app}\settings.json');
    Json :=
      '{' + #13#10 +
      '  "connections": {' + #13#10 +
      '    "google_sheets": {' + #13#10 +
      '      "credentials_path": "' + EscPath + '"' + #13#10 +
      '    }' + #13#10 +
      '  }' + #13#10 +
      '}' + #13#10;
    { Only write if settings.json doesn't already exist, so we never clobber
      an existing user's configured settings on reinstall/update. }
    if not FileExists(SettingsPath) then
      SaveStringToFile(SettingsPath, Json, False);
  end;
end;
