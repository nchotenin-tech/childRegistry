#define AppVersion "1.0.0"
[Setup]
AppId={{BFE4EA2D-3797-42BC-A923-5AE32B031E57}
AppName=ChildRegistry
AppVersion={#AppVersion}
AppPublisher=ChildRegistry
AppPublisherURL=https://github.com/nchotenin-tech/childRegistry
DefaultDirName={localappdata}\Programs\ChildRegistry
DefaultGroupName=ChildRegistry
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\installer-output
OutputBaseFilename=ChildRegistry-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
UninstallDisplayIcon={app}\ChildRegistry.exe
MinVersion=10.0.19041
[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked
[Files]
Source: "..\dist\ChildRegistry\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{group}\ChildRegistry"; Filename: "{app}\ChildRegistry.exe"
Name: "{autodesktop}\ChildRegistry"; Filename: "{app}\ChildRegistry.exe"; Tasks: desktopicon
Name: "{group}\User Manual"; Filename: "{app}\_internal\USER_MANUAL_TH.html"
[Run]
Filename: "{app}\ChildRegistry.exe"; Description: "Launch ChildRegistry"; Flags: nowait postinstall skipifsilent
