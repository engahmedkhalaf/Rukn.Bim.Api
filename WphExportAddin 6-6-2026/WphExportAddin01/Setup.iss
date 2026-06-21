; Inno Setup script for RUKN BIM API Export Revit Add-in
; Supports Revit 2023
; Uses absolute paths so it compiles successfully from any directory or drive

[Setup]
AppName=RUKN BIM API Export Revit Add-In
AppVersion=1.0.0
AppPublisher=RUKN BIM API
DefaultDirName={commonappdata}\Autodesk\Revit\Addins\2023\WphExport
DefaultGroupName=RUKN BIM API Export
DirExistsWarning=no
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputBaseFilename=RuknBimApiExportSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Copy the Revit Add-In manifest file to the main Addins folder (using absolute path)
Source: "d:\API Khalaf\Rukn.Bim.Api\WIP\REVIT\WphExportAddin 6-6-2026\WphExportAddin01\WphExportAddin01\WphExportAddin01.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion

; Copy the binaries to the subfolder (using absolute paths)
Source: "d:\API Khalaf\Rukn.Bim.Api\WIP\REVIT\WphExportAddin 6-6-2026\WphExportAddin01\WphExportAddin01\bin\Release\WphExportAddin01.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "d:\API Khalaf\Rukn.Bim.Api\WIP\REVIT\WphExportAddin 6-6-2026\WphExportAddin01\WphExportAddin01\bin\Release\WphExportAddin01.pdb"; DestDir: "{app}"; Flags: ignoreversion

; Copy the resource icons folder (using absolute path)
Source: "d:\API Khalaf\Rukn.Bim.Api\WIP\REVIT\WphExportAddin 6-6-2026\WphExportAddin01\WphExportAddin01\bin\Release\Resources\*"; DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs createallsubdirs

[Code]
// Helper code can go here if you need to detect Revit installation or close Revit before installation.
