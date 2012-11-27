#define SetupBaseName   "SetupGeoSan-v."
#define AppVersionFile  "06.00.07.12"

[Setup]
AppName=GeoSan
Compression=lzma
AllowNoIcons=no
AlwaysRestart=no
AlwaysShowComponentsList=yes
; equivale ao Product Version em propriedades do Setup gerado
AppVerName=GeoSan {#AppVersionFile}
; equivale ao file version em propriedades do Setup gerado
VersionInfoVersion={#AppVersionFile} 
VersionInfoTextVersion=GeoSan {#AppVersionFile}
AppCopyright=NEXUS GeoEngenharia

DefaultDirName=C:\Arquivos de Programas\GeoSan
DefaultGroupName=GeoSan
UninstallDisplayIcon={app}\GeoSan.exe
AppMutex=GeoSan

SolidCompression=yes
OutputDir=Output
OutputBaseFilename={#SetupBaseName + AppVersionFile}

[Language]
MessagesFile=compiler:BrazilianPortuguese.isl

[Files]

; CopyMode: alwaysskipifsameorolder para manter caso não exista diferença de versão
; CopyMode: alwaysoverwrite para sobrescrever tudo

;ARQUIVOS NEXUS - *************************************************************************************************************************
Source: "ArquivosInstGeoSan\GeoSan.exe";          DestDir: "{app}";             CopyMode: alwaysoverwrite;
;Source: "ArquivosInstGeoSan\GeoSan.bat";          DestDir: "{app}";             CopyMode: alwaysoverwrite;
;Source: "ArquivosInstGeoSan\Instruções para configuração de atualizador automatico de versões.txt";          DestDir: "{app}";             CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\00. Leiame.txt";          DestDir: "{app}";             CopyMode: alwaysoverwrite;

Source: "ArquivosInstGeoSan\Exporte EPANet.exe";  DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\ValidaBase.exe";      DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\NSecurity.dll";       DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NUsers.dll";          DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NexusConnection.dll"; DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NxViewManager.ocx";   DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NexusPM4.ocx";        DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror

;Source: "ArquivosInstGeoSan\LoozeXP.ocx";        DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "ArquivosInstGeoSan\databaseImage.png";  DestDir: "{app}";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "ArquivosInstGeoSan\Manual do Usuário Geosan.doc"; CopyMode: alwaysoverwrite; DestDir: "{app}"; Flags: isreadme
;Source: "MyProg.chm"; DestDir: "{app}"


;ARQUIVOS VISUAL BASIC 6 - ****************************************************************************************************************
Source: "ArquivosInstGeoSan\ASYCFILT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\COMCAT.DLL";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\COMDLG32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSADO26.TLB";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regtypelib noregerror
Source: "ArquivosInstGeoSan\MSCHRT20.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSCOMCT2.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSCOMCTL.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSDBRPT.DLL";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSFLXGRD.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSMASK32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSSTDFMT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSVBVM60.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\MSVCRT.DLL";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\OLEAUT32.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\OLEPRO32.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\RICHTX32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\STDOLE2.TLB";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regtypelib noregerror
Source: "ArquivosInstGeoSan\TABCTL32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\VB6STKIT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror


;ARQUIVOS TERRALIB 3.3.1.7 - *******************************************************************************************************************

;A Tecom será instalada por instalador "3 - TerraComponents 3.3.1.7". Os arquivos tecom não serão incluidos no instalador do Geosan

;Source: "TECOM 3.3.1.5\TeComCanvas.dll";       DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComDatabase.dll";     DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComExport.dll";       DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComImport.dll";       DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComNetwork.dll";      DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComPrinter.dll";      DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComViewDatabase.dll"; DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "TECOM 3.3.1.5\TeComViewManager.dll";  DestDir: "{app}\Controles"; CopyMode: alwaysoverwrite; Flags: regserver noregerror


[Icons]
;Name: "{group}\My Program"; Filename: "{app}\MyProg.exe"
Name: "{commonprograms}\GeoSan"; Filename: "{app}\GeoSan.exe"
Name: "{commondesktop}\GeoSan";  Filename: "{app}\GeoSan.exe"
