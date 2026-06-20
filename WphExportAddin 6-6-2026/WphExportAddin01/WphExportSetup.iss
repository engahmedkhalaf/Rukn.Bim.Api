; ==========================================================
; WPH Export - Installer
; Revit 2023
; ==========================================================

#define MyAppName "WPH Export"
#define MyAppVersion "1.0.0"
#define PublishDir "D:\API Khalaf\Revit API\Publish\NWC EXPORTER Scope"

[Setup]
AppId={{9F2C5D4A-3B71-4E88-9C21-7D5A6E0F1234}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={commonappdata}\Autodesk\Revit\Addins\2023\WphExport
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=WphExportSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "{commonappdata}\Autodesk\Revit\Addins\2023\WphExport"

[Files]
; Add-in manifest
Source: "{#PublishDir}\WphExportAddin01.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion

; All application files (DLL, PDB, Resources subfolder)
Source: "{#PublishDir}\WphExport\*"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023\WphExport"; Flags: ignoreversion recursesubdirs createallsubdirs

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
