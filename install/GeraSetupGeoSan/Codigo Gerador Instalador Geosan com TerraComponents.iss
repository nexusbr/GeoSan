#define SetupBaseName   "SetupGeoSan-v."
#define AppVersionFile  "06.10.28"

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
;para não aparecer ao usuário para ele entrar o diretório de instalação do GeoSan
DisableDirPage=yes              
;para não aparecer o nome do Grupo no menu Início do Windows para o usuário selecionar
DisableProgramGroupPage=yes

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
Source: "ArquivosInstGeoSan\GeoSanIni.exe";       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\NSecurity.dll";       DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NUsers.dll";          DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NexusConnection.dll"; DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NxViewManager2.ocx";  DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\NexusPM4.ocx";        DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror

;Source: "ArquivosInstGeoSan\LoozeXP.ocx";        DestDir: "{app}\Controles";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "ArquivosInstGeoSan\databaseImage.png";  DestDir: "{app}";   CopyMode: alwaysoverwrite; Flags: regserver noregerror
;Source: "ArquivosInstGeoSan\Manual do Usuário Geosan.doc"; CopyMode: alwaysoverwrite; DestDir: "{app}"; Flags: isreadme
;Source: "MyProg.chm"; DestDir: "{app}"


;ARQUIVOS VISUAL BASIC 6 - ****************************************************************************************************************
Source: "ArquivosInstGeoSan\ASYCFILT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\COMCAT.DLL";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\COMDLG32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSADO26.TLB";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regtypelib noregerror 32bit
Source: "ArquivosInstGeoSan\MSCHRT20.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSCOMCT2.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSCOMCTL.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSDBRPT.DLL";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSFLXGRD.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSMASK32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSSTDFMT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSVBVM60.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSVCRT.DLL";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\OLEAUT32.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\OLEPRO32.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\RICHTX32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\STDOLE2.TLB";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regtypelib noregerror 32bit
Source: "ArquivosInstGeoSan\TABCTL32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\VB6.OLB";      DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\VB6STKIT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\scrrun.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\cdosys.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSINET.OCX";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\mswsock.dll";  DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit
Source: "ArquivosInstGeoSan\MSCOMCTL.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit

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
Name: "{commonprograms}\GeoSan"; Filename: "{app}\GeoSanIni.exe"
Name: "{commondesktop}\GeoSan";  Filename: "{app}\GeoSanIni.exe"
