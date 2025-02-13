; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "IPA-AreaCalculations"
#define MyAppVersion "1.00"
#define MyAppPublisher "IPA Architecture and More"
#define MyAppExeName "MyProg-x64.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{75B951C9-A0E1-43AA-BB76-DDEAF1781D38}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
CreateAppDir=no
LicenseFile=B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\LICENSE.txt
; Uncomment the following line to run in non administrative install mode (install for current user only).
;PrivilegesRequired=lowest
OutputDir=B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\Releases
OutputBaseFilename=IPA-AreaCalculationsV1.00
SetupIconFile=B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\img\installerIcon.ico
Password=ipaMipa
Encryption=yes
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\bin\Release\AreaCalculations.dll"; DestDir: "C:\Program Files\IPA\AreaCalculations\"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\AreaCalculations.addin"; DestDir: "C:\ProgramData\Autodesk\Revit\Addins\2021"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\AreaCalculations.addin"; DestDir: "C:\ProgramData\Autodesk\Revit\Addins\2022"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\AreaCalculations.addin"; DestDir: "C:\ProgramData\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\AreaCalculations.addin"; DestDir: "C:\ProgramData\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\img\areacIcon.png"; DestDir: "C:\Program Files\IPA\AreaCalculations\"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\img\areaIcon.png"; DestDir: "C:\Program Files\IPA\AreaCalculations\"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\img\excelIcon.png"; DestDir: "C:\Program Files\IPA\AreaCalculations\"; Flags: ignoreversion
Source: "B:\06. BIM AUTOMATION\02. C#\AREA CALCULATIONS\NewAreaCalculationsProject\img\plotIcon.png"; DestDir: "C:\Program Files\IPA\AreaCalculations\"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

