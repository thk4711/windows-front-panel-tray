; Inno Setup Script for Tray Serial Monitor
; This creates a Windows installer using standalone executables (no Python required)

[Setup]
AppName=Tray Serial Monitor Service
AppVersion=2.0
AppPublisher=Your Company
AppPublisherURL=https://your-website.com
AppSupportURL=https://your-website.com/support
AppUpdatesURL=https://your-website.com/updates
DefaultDirName={autopf}\TraySerialMonitor
DefaultGroupName=Tray Serial Monitor
AllowNoIcons=yes
LicenseFile=
InfoBeforeFile=
InfoAfterFile=
OutputDir=installer_output
OutputBaseFilename=TraySerialMonitor_Standalone_Setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "C:\TrayTemp"; Permissions: everyone-full

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startupicon"; Description: "Start client application automatically when Windows starts"; GroupDescription: "Startup Options"

[Files]
; Standalone executable files (no Python required) - FIXED VERSION
Source: "exe\TrayHardwareMonitorService.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "exe\TraySerialMonitorClient.exe"; DestDir: "{app}"; Flags: ignoreversion

; Service management executables (standalone - no Python required)
Source: "exe\InstallService.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "exe\UninstallService.exe"; DestDir: "{app}"; Flags: ignoreversion

; Configuration and resources
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

; LibreHardwareMonitor dependencies (required by service executable)
Source: "LibreHardwareMonitorLib.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "LibreHardwareMonitorLib.sys"; DestDir: "{app}"; Flags: ignoreversion


[Icons]
Name: "{group}\Tray Serial Monitor"; Filename: "{app}\TraySerialMonitorClient.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"
Name: "{group}\Service Manager"; Filename: "{app}\InstallService.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"
Name: "{group}\Uninstall Service"; Filename: "{app}\UninstallService.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"
Name: "{group}\{cm:UninstallProgram,Tray Serial Monitor Service}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Tray Serial Monitor"; Filename: "{app}\TraySerialMonitorClient.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
; Install and start the Windows service using standalone executable
Filename: "{app}\InstallService.exe"; WorkingDir: "{app}"; Flags: runhidden waituntilterminated; StatusMsg: "Installing Windows service..."
; Start the client application after installation
Filename: "{app}\TraySerialMonitorClient.exe"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent; Description: "{cm:LaunchProgram,Tray Serial Monitor}"

[UninstallRun]
; Stop and remove the Windows service during uninstallation using standalone executable
Filename: "{app}\UninstallService.exe"; WorkingDir: "{app}"; Flags: runhidden waituntilterminated; RunOnceId: "UninstallTrayService"

[Registry]
; Create startup entry for client application if requested (system-wide since we have admin privileges)
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "TraySerialMonitor"; ValueData: """{app}\TraySerialMonitorClient.exe"""; Tasks: startupicon

[UninstallDelete]
; Clean up any additional files created during runtime
Type: files; Name: "{app}\*.log"
Type: files; Name: "{app}\*.tmp"
; Clean up PyInstaller temp directory and its contents
Type: filesandordirs; Name: "C:\TrayTemp"

[Code]
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  // Check if service is already running and stop it
  Exec('net', 'stop TrayHardwareMonitor', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := '';
end;
