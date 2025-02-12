[Setup]
AppName=Account Payables Automation
AppVersion=1.0
AppPublisher=Your Company Name
DefaultDirName={commonpf}\AccountPayablesAutomation
UninstallDisplayIcon={app}\AccountPayablesAPP.bat
OutputBaseFilename=AccountPayablesAutomationInstaller
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "python-3.12.8-amd64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "app.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "Payable_Account_Automation.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "first_run.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "AccountPayablesAPP.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

[Code]
var
  PythonInstalled: Boolean;
  ResultCode: Integer;

function NeedsPython: Boolean;
var
  PythonVersion: String;
begin
  Result := True;
  
  // Check registry installation
  if RegKeyExists(HKLM, 'SOFTWARE\Python\PythonCore\3.12\InstallPath') then
  begin
    Result := False;
    Exit;
  end;

  // Check Python in PATH
  if Exec(ExpandConstant('{cmd}'), '/c where python && exit 0 || exit 1', '', SW_HIDE, 
    ewWaitUntilTerminated, ResultCode) then
  begin
    if ResultCode = 0 then
    begin
      if Exec(ExpandConstant('{cmd}'), '/c python --version', '', SW_HIDE, 
        ewWaitUntilTerminated, ResultCode) then
      begin
        if GetVersionNumbersString('python.exe', PythonVersion) then
        begin
          Result := (CompareStr(Copy(PythonVersion, 1, 4), '3.12') < 0);
        end;
      end;
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  PythonInstalled := not NeedsPython;
  Result := True;
end;

[Run]
Filename: "{tmp}\python-3.12.8-amd64.exe"; Parameters: "/quiet InstallAllUsers=1 PrependPath=1 Include_pip=1"; Check: NeedsPython
Filename: "{cmd}"; Parameters: "/c python -m venv ""{app}\venv"""; Flags: runhidden
Filename: "{app}\venv\Scripts\python.exe"; Parameters: "-m pip install -r ""{app}\requirements.txt"""; Flags: runhidden