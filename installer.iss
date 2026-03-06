; installer.iss -- Inno Setup script for DocDivide
;
; Requirements:
;   - Inno Setup 6: https://jrsoftware.org/isdl.php
;   - PyInstaller output in dist\DocDivide\  (run PyInstaller first)
;
; To build the installer:
;   1. Install Inno Setup 6
;   2. Open this file in Inno Setup Compiler
;   3. Click Build > Compile  (or press F9)
;   Output: Output\DocDivideSetup.exe

#define AppName      "DocDivide"
#define AppVersion   "1.0.0"
#define AppPublisher "Your Company Name"
#define AppExeName   "DocDivide.exe"
#define SourceDir    "dist\DocDivide"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://yourcompany.com
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename=DocDivideSetup
SetupIconFile=icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";    Description: "Create a &desktop shortcut";     GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startmenuicon"; Description: "Create a &Start Menu shortcut";   GroupDescription: "Additional icons:"; Flags: checkedonce

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}";             Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}";   Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";       Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
procedure InitializeWizard();
begin
  MsgBox(
    'Welcome to DocDivide' + #13#10 + #13#10 +
    'This application uses the Anthropic Claude API to read ' +
    'engineering drawing title blocks.' + #13#10 + #13#10 +
    'Each user will need to provide their own Anthropic API key ' +
    'when they first launch the application.' + #13#10 + #13#10 +
    'API keys can be created at: https://console.anthropic.com',
    mbInformation, MB_OK
  );
end;
