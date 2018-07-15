; -- AisDecoder.iss --

;SourcePath is where the .iss file is located
#pragma message SourcePath
#define MyAppName "RacingSignals.exe" 
#pragma message "MyAppName info: " + MyAppName
#define MyAppFile SourcePath + MyAppName
#pragma message "MyAppFile info: " + MyAppFile
#define MyAppVersion GetFileVersion(MyAppFile)
#pragma message "Detailed version info: " + MyAppVersion
#define MyAppVersion StringChange(MyAppVersion, ".0.", "." )
#pragma message "Stripped version info: " + MyAppVersion

#define public MyFileDateTimeString GetFileDateTimeString(MyAppFile, 'dd/mm/yyyy hh:nn:ss', '-', ':');
#pragma message "File Date info: " + MyFileDateTimeString
#define MyDateTimeString GetDateTimeString('dd/mm/yyyy hh:nn:ss', '-', ':');

#define MyProgramData "C:\ProgramData"
#pragma message "MyProgramData: " + MyProgramData
#define MySys32 "C:\Windows\SysWOW64"
#pragma message "MySys32: " + MySys32

;#define ch FileOpen("c:\website\backup.bat")
;#define batcommand FileRead(ch)
;#define batcommand StringChange(batcommand, "Version", MyAppVersion )

; this works #define result Exec('cmd /c xcopy/s/y/q', '"e:\My Documents\Ais\NmeaRouterSource" "e:\My Documents\Ais\NmeaRouter_Backup\NmeaRouter_1.1.7\"')
#define result Exec('cmd /c xcopy/s/y/q', '"c:\Users\Admin\Documents\Ais\YachtSource" "c:\Users\Admin\Documents\Ais\YachtSourceBackup\YachtSource_' + MyAppVersion + '\"')
;#Define result Exec('cmd /c dir/p', '"e:\My Documents"')

[Setup]
;version explorer displays for setup.exe, recovered with VB6 app.major & app.minor
VersionInfoVersion={#MyAppVersion}
;minimum windows version sofware will run on (0=no Win98, 4.0= nt or 2000,XP upwards)
;MinVersion= 4.10,4.0
MinVersion= 0,5.0
AppName=RacingSignals
AppId=RacingSignals
;CreateUninstallRegKey=no
;UpdateUninstallLogAppName=no
;On INNO installer "This will install Ais Decoder Version x.x.x.x on your computer"
AppVerName=RacingSignals
AppPublisher=Neal Arundale
AppPublisherURL=http://web.arundale.co.uk/docs/ais/sp_map.html
;where the users files are placed
DefaultDirName={pf}\Arundale\RacingSignals
DefaultGroupName=RacingSignals
;UsePreviousAppDir=No
;UninstallDisplayIcon=E:\jna\arundale\website\docs\arundale.ico
;20/6/15 UninstallDisplayIcon=arundale.ico
;outputdir=E:\jna\Arundale\website\docs\ais\
;outputdir=C:\website\
;outputdir="C:\Users\Admin\My Documents\DirectNic\Live Parent (ArundaleCom)\docs\ais\"
outputdir="C:\Users\Admin\Documents\ais\YachtSource"
OutputBaseFilename= RacingSignals_setup_{#MyAppVersion}
setuplogging=yes
;20/6/15 SetupIconFile=arundale.ico
;directory icon for setup file (Icon file expected in YachtSource)
SetupIconFile=bluepeter.ico
;required for vbfiles installation
PrivilegesRequired=admin
LicenseFile=license.txt
;FileDateTimeString= (#MyFileDateTimeString)
AppMutex="RacingSignals"

[Dirs]
;only required if creating an empty directory [files] creates the directory
;these get copied to userappdata when AisDecoderns new version
;Required for RacingSignals/Results .csv file will not write out to My Documents!!
Name: "{userappdata}\Arundale\RacingSignals"
[Files]
; begin VB system files
;dll'S CANNOT BE IN SYSTEM DIRECTORY
Source: "vbfiles\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vbfiles\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "vbfiles\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\Support\vb6stkit.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\msstdfmt.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\Support\msvcrt.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Dont think I need this for RacingSignals
;Source: "vbfiles\Support\scrrun.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\capicom.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver

Source: "{#MySys32}\MSWINSCK.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\mscomctl.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\ComDlg32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSHFlxGd.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSINET.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
;Windows 8
Source: "{#MySys32}\richtx32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

; end VB system files
Source: "{#MyAppName}"; DestDir: "{app}" ;flags: replacesameversion ignoreversion
;Source: "E:\My Documents\Ais\Decoder_v3\{#MyAppVersion}.txt"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Files" ;flags: replacesameversion ignoreversion
Source: "arundale.ico"; DestDir: "{app}"  ;flags: replacesameversion ignoreversion
;Source: "com0com\setup_com0com-3.0.0.0-i386-and-x64-unsigned.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\setup_com0com_W7_x64_signed.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\setup_com0com_W7_x86_signed.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\setup_com0com-2.2.2.0-x64-fre-signed.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\com0com_setup_driver.bat"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\com0com_setup_port.bat"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\com0com_remove_port.bat"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
Source: "C:\Documents and Settings\All Users\Application Data\Arundale\RacingSignals\Sequences\*.ini"; DestDir: "{commonappdata}\Arundale\RacingSignals\Sequences"  ;flags: replacesameversion ignoreversion
Source: "C:\Documents and Settings\All Users\Application Data\Arundale\RacingSignals\Sounds\*.wav"; DestDir: "{commonappdata}\Arundale\RacingSignals\Sounds"  ;flags: replacesameversion ignoreversion
Source: "C:\Documents and Settings\All Users\Application Data\Arundale\RacingSignals\SignalImages\*.gif"; DestDir: "{commonappdata}\Arundale\RacingSignals\SignalImages"  ;flags: replacesameversion ignoreversion

;help
;Source: "Help\NmeaRouter.chm"; DestDir: "{app}\Help"  ;flags: replacesameversion ignoreversion
;above need uncommenting
;Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme ignoreversion

;Creates the Shortcuts
[Icons]
;20/6/15 Name: "{group}\RacingSignals"; Filename: "{app}\RacingSignals.exe"; IconFilename:"{app}\arundale.ico"
;Icon is set up as property in frmMain (see Project Properties > Make)
Name: "{group}\RacingSignals"; Filename: "{app}\RacingSignals.exe"
; NOTE: Most apps do not need registry entries to be pre-created. If you
; don't know what the registry is or if you need to use it, then chances are
; you don't need a [Registry] section.
;Name: "{userdocs}\Ais Decoder"; Filename: "{userappdata}\Arundale\Ais Decoder\Output"; Flags: foldershortcut ; IconFilename:"{app}\arundale.ico" ;Comment:"AisDecoder Files"

[InstallDelete]
Type: files; Name: "{app}\RacingSignals.exe"

[Registry]
Root: HKCU; Subkey: "Software\Arundale"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\Arundale\RacingSignals"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Arundale\RacingSignals\Profiles"; Flags: uninsdeletekey
[Run]
;Filename: "{app}\license.txt"; Description: "View the README file"; Flags: postinstall shellexec unchecked skipifsilent
;line below causes error invalid control array index in NmeaRouter
;Filename: "{app}\com0com\setup.exe";  Parameters: "/S /D={app}\com0com\"; Description: "Install Virtual Com Port (VCP) support"; Flags: postinstall nowait skipifsilent
;Filename: "{app}\com0com\setup.exe";  Parameters: "/S /D={app}\com0com\"; StatusMsg: "Installing Virtual Com Port (VCP) driver ..."; Flags: runminimized
;Filename: "{app}\com0com\com0com_setup_driver.bat";  Parameters: "/S /D={app}\com0com\"; StatusMsg: "Installing Virtual Com Port (VCP) driver ..."; Flags: runminimized

;Remove these old ini files silently (if the exist)
;Filename: "cmd.exe"; workingdir: "{commonappdata}\Arundale\RacingSignals\Sequences"; parameters: "/C ""del ScarboroughMultiple.ini"""; Flags:skipifdoesntexist shellexec runhidden
;Filename: "cmd.exe"; workingdir: "{commonappdata}\Arundale\RacingSignals\Sequences"; parameters: "/C ""del ScarboroughSingle.ini"""; Flags:skipifdoesntexist shellexec runhidden

Filename: "{app}\RacingSignals.exe"; Description: "Launch application"; Flags: postinstall nowait skipifsilent

