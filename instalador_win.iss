[Setup]
AppName=EcoRenamer Pro
AppVersion=1.4.3
DefaultDirName={autopf}\EcoRenamerPro
DefaultGroupName=EcoRenamer Pro
OutputDir=Output
OutputBaseFilename=Setup_EcoRenamer_v1.4.3
Compression=lzma2
SolidCompression=yes
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\RenomeadorApp.exe

[Files]
Source: "dist\RenomeadorApp\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\EcoRenamer Pro"; Filename: "{app}\RenomeadorApp.exe"
Name: "{autodesktop}\EcoRenamer Pro"; Filename: "{app}\RenomeadorApp.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar um atalho na Área de Trabalho"; GroupDescription: "Atalhos Adicionais:"
