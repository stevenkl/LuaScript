; Script generated by the Inno Script Studio Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define _WIN32 1

#ifdef _WIN64
  #undef _WIN32
#endif

#define MyAppName "LuaScript"
#define MyAppVersion "0.2.0.0"
#define MyAppPublisher "Natural Style Co. Ltd."
#define MyAppURL "http://na-s.jp/"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{850B67F5-7449-4E93-B07B-1D679E2D0590}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=.\Build

#ifdef _WIN32
  OutputBaseFilename=LuaScript-{#MyAppVersion}_x86-Setup
#endif

#ifdef _WIN64
  OutputBaseFilename=LuaScript-{#MyAppVersion}_x64-Setup
#endif

Compression=lzma
SolidCompression=yes
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardSmallImageFile=compiler:WizModernSmallImage-IS.bmp
ChangesAssociations=True
AllowUNCPath=False
EnableDirDoesntExistWarning=True
UninstallDisplayName={#MyAppName} {#MyAppVersion}
UninstallDisplayIcon={uninstallexe}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoTextVersion={#MyAppVersion}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoProductTextVersion={#MyAppVersion}
SetupLogging=True
SourceDir=.

#ifdef _WIN64
ArchitecturesInstallIn64BitMode=x64
#endif

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Icons]
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Files]
Source: "Sources\LuaScript.0.2.0.0\LuaScript_x86.dll"; DestDir: "{app}"; Flags: 32bit regserver

#ifdef _WIN64
Source: "Sources\LuaScript.0.2.0.0\LuaScript_x64.dll"; DestDir: "{app}"; Flags: 64bit regserver; Check: IsWin64
#endif
Source: "Sources\LuaScript.0.2.0.0\Readme.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "Sources\LuaScript.0.2.0.0\test.asp"; DestDir: "{app}"; Flags: ignoreversion
Source: "Sources\LuaScript.0.2.0.0\test.lua"; DestDir: "{app}"; Flags: ignoreversion
Source: "Sources\LuaScript.0.2.0.0\test.wsf"; DestDir: "{app}"; Flags: ignoreversion
Source: "Sources\LuaScript.0.2.0.0.source.zip"; DestDir: "{app}"; Flags: ignoreversion
