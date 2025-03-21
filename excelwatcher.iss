[Setup]
AppName=ExcelWatcher
AppVersion=1.0
DefaultDirName={pf}\ExcelWatcher
DefaultGroupName=ExcelWatcher
OutputDir=.
OutputBaseFilename=ExcelWatcherInstaller
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Files]
Source: "target\release\excel-watcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "setup.bat"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{app}\setup.bat"; Description: "Install dependencies and run ExcelWatcher"; Flags: nowait runhidden postinstall
