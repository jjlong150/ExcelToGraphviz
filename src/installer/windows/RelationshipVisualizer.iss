; Relationship Visualizer installer for Windows.
;
; Compile with the Inno Setup 6 command-line compiler:
;     ISCC.exe RelationshipVisualizer.iss
;
; The compiled installer is written to dist\RelationshipVisualizerSetup.exe.
;
; What this does that the plain zip download does not:
;   1. Shows the MIT license and requires the user to accept it (native
;      Inno Setup wizard page).
;   2. Detects Graphviz's dot.exe (PATH, then Program Files / Program
;      Files (x86)) and offers to open the Graphviz download page if it
;      isn't found, rather than leaving the user to run `dot -V` blind.
;   3. Runs `dot -c` to register Graphviz's plugins, the step the manual
;      docs warn is easy to forget.
;   4. Registers the install folder as an Excel "Trusted Location" for the
;      current user, so opening Relationship Visualizer.xlsm doesn't show
;      a macro-security prompt.
;   5. Lets the user opt out of installing the sample workbooks to save
;      disk space, via a normal Components checkbox instead of a manual
;      "skip this folder" step.
;   6. Adds Start Menu shortcuts to the documentation website and to the
;      full PDF export of it (both live at exceltographviz.com, not
;      bundled), instead of shipping the old standalone user-manual PDF.

#define AppVersion "11.0.0"
#define DistDir "..\..\..\dist\Relationship Visualizer"
#define OutputDir "..\..\..\dist"

[Setup]
AppId={{EF9C08A4-90A0-4CC3-BB04-B75D9F5ECE6B}
AppName=Relationship Visualizer
AppVersion={#AppVersion}
AppVerName=Relationship Visualizer {#AppVersion}
AppPublisher=Jeffrey Long
AppPublisherURL=https://exceltographviz.com/
AppSupportURL=https://exceltographviz.com/
AppUpdatesURL=https://exceltographviz.com/
DefaultDirName={userdocs}\Relationship Visualizer
DefaultGroupName=Relationship Visualizer
LicenseFile={#DistDir}\license.txt
OutputDir={#OutputDir}
OutputBaseFilename=RelationshipVisualizerSetup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayIcon={app}\Relationship Visualizer.xlsm
DisableWelcomePage=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "compact"; Description: "Compact installation (no sample workbooks)"
Name: "full"; Description: "Full installation (includes sample workbooks)"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main"; Description: "Relationship Visualizer (required)"; Types: full compact custom; Flags: fixed
Name: "samples"; Description: "Sample workbooks (uses extra disk space)"; Types: full

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "{#DistDir}\Relationship Visualizer.xlsm"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#DistDir}\README.txt"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#DistDir}\changelog.md"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#DistDir}\security.md"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#DistDir}\license.txt"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#DistDir}\licenses\*"; DestDir: "{app}\licenses"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: main
Source: "{#DistDir}\samples\*"; DestDir: "{app}\samples"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: samples

[Icons]
Name: "{group}\Relationship Visualizer"; Filename: "{app}\Relationship Visualizer.xlsm"
Name: "{group}\Relationship Visualizer Online Documentation"; Filename: "https://exceltographviz.com/"
Name: "{group}\Relationship Visualizer PDF Documentation"; Filename: "https://exceltographviz.com/relationship_visualizer.pdf"
Name: "{group}\Uninstall Relationship Visualizer"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Relationship Visualizer"; Filename: "{app}\Relationship Visualizer.xlsm"; Tasks: desktopicon

[Registry]
; Trust the install folder in Excel so opening the workbook doesn't show a
; macro-security prompt. "16.0" is the Office registry version used by
; Office 2016, 2019, 2021, 2024, LTSC, and Microsoft 365 alike.
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationRelationshipVisualizer"; ValueType: string; ValueName: "Path"; ValueData: "{app}\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationRelationshipVisualizer"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: "1"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationRelationshipVisualizer"; ValueType: string; ValueName: "Description"; ValueData: "Relationship Visualizer"

[Code]
var
  DotPath: String;
  DotFoundGlobal: Boolean;

function FindDotOnPath(): String;
var
  PathVar, Dir: String;
  P: Integer;
  Candidate: String;
begin
  Result := '';
  PathVar := ExpandConstant('{%PATH}');
  while Length(PathVar) > 0 do
  begin
    P := Pos(';', PathVar);
    if P = 0 then
    begin
      Dir := PathVar;
      PathVar := '';
    end
    else
    begin
      Dir := Copy(PathVar, 1, P - 1);
      PathVar := Copy(PathVar, P + 1, Length(PathVar));
    end;
    Dir := Trim(Dir);
    if Length(Dir) > 0 then
    begin
      Candidate := AddBackslash(Dir) + 'dot.exe';
      if FileExists(Candidate) then
      begin
        Result := Candidate;
        Exit;
      end;
    end;
  end;
end;

function FindDotUnder(BaseDir: String): String;
var
  FindRec: TFindRec;
  Candidate: String;
begin
  Result := '';
  if (BaseDir = '') or (not DirExists(BaseDir)) then
    Exit;

  Candidate := AddBackslash(BaseDir) + 'Graphviz\bin\dot.exe';
  if FileExists(Candidate) then
  begin
    Result := Candidate;
    Exit;
  end;

  if FindFirst(AddBackslash(BaseDir) + 'Graphviz*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        begin
          Candidate := AddBackslash(BaseDir) + FindRec.Name + '\bin\dot.exe';
          if FileExists(Candidate) then
          begin
            Result := Candidate;
            Exit;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function DetectGraphviz(): String;
begin
  Result := FindDotOnPath();
  if Result = '' then
    Result := FindDotUnder(ExpandConstant('{pf}'));
  if Result = '' then
    Result := FindDotUnder(ExpandConstant('{pf32}'));
end;

function InitializeSetup(): Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;
  DotPath := DetectGraphviz();
  DotFoundGlobal := (DotPath <> '');
  if not DotFoundGlobal then
  begin
    if MsgBox('Graphviz''s dot.exe was not found on this computer.' + #13#10 + #13#10 +
       'Relationship Visualizer requires Graphviz to generate diagrams. You can install Graphviz first and then run this setup again, or continue installing Relationship Visualizer now and finish the Graphviz setup later by re-running this installer.' + #13#10 + #13#10 +
       'Open the Graphviz download page and stop this setup now?', mbConfirmation, MB_YESNO) = IDYES then
    begin
      ShellExec('open', 'https://graphviz.org/download/', '', '', SW_SHOW, ewNoWait, ErrorCode);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if (CurStep = ssPostInstall) and DotFoundGlobal then
  begin
    // Registering Graphviz's plugins writes into Graphviz's own install
    // folder, which usually requires administrator rights, independent of
    // whether this setup itself is running elevated. Request elevation
    // just for this one command via the 'runas' verb, rather than forcing
    // the whole install to run as admin.
    if not ShellExec('runas', DotPath, '-c', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      MsgBox('Relationship Visualizer could not automatically register Graphviz''s plugins (this requires administrator rights).' + #13#10 + #13#10 +
        'Please open a Command Prompt as Administrator and run:' + #13#10 + #13#10 +
        '"' + DotPath + '" -c', mbInformation, MB_OK);
    end;
  end;
end;
