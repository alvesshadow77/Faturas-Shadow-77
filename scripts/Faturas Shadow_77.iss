[Setup]
AppName=Faturas Shadow_77
AppVersion=9.7
AppPublisher=Faturas Shadow_77
DefaultDirName={autopf}\Faturas Shadow_77
DefaultGroupName=Faturas Shadow_77
AllowNoIcons=yes
OutputDir=..\Instalador_Output
OutputBaseFilename=Faturas_Shadow_77_Setup_v9.7
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64
LanguageDetectionMethod=uilanguage

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Files]
Source: "..\dist\Faturas Shadow_77.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Faturas Shadow_77"; Filename: "{app}\Faturas Shadow_77.exe"
Name: "{autodesktop}\Faturas Shadow_77"; Filename: "{app}\Faturas Shadow_77.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Run]
Filename: "{app}\Faturas Shadow_77.exe"; Description: "{cm:LaunchProgram,Faturas Shadow_77}"; Flags: nowait postinstall skipifsilent