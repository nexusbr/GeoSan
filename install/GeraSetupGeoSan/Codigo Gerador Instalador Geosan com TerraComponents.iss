#define SetupBaseName   "SetupGeoSan-v."
#define AppVersionFile  "06.10.36"

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

;ARQUIVOS TERRALIB 4.2.0 - *******************************************************************************************************************
Source: "ArquivosInstGeoSan\TeCom4.2.0\fbclient.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\gdal110.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\geotiff.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\iconv.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\intl.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libeay32.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libexpat.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libiconv-2.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libintl-2.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libmysql.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libpq.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libtiff.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\msadox.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\msvcr71.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\oci.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\oledb32.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\oraociicus10.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\qt-mt338.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\qwt500.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\ssleay32.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\terralib.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\terralib_ado.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\terralib_shp.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\terralib_spl.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
; dos antigos
Source: "ArquivosInstGeoSan\TeCom4.2.0\ijl15.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libiconv2.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\libxml2.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\msjava.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\msvcp80.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\msvcp80d.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\msvcr80.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\shapelib.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\SIBPRO2.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\tiff.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\zlib.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\TeCom4.2.0\zlib1.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;

; dos antigos
;Source: "ArquivosInstGeoSan\TeCom4.2.0\TECOMP~1.oca";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite;  Flags: regserver noregerror

Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComCanvas.dll";       DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComConnection.dll";   DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComDatabase.dll";     DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComExport.dll";       DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComGeometry.dll";     DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComImport.dll";       DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComNetwork.dll";      DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComPrinter.dll";      DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComViewDatabase.dll"; DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror
Source: "ArquivosInstGeoSan\TeCom4.2.0\TeComViewManager.dll";  DestDir: "C:\Arquivos de Programas\NEXUS\TeCom4.2.0"; CopyMode: alwaysoverwrite; Flags: regserver noregerror

;ARQUIVOS NEXUS - *************************************************************************************************************************
Source: "ArquivosInstGeoSan\GeoSan.exe";          DestDir: "{app}";             CopyMode: alwaysoverwrite;
;Source: "ArquivosInstGeoSan\GeoSan.bat";          DestDir: "{app}";             CopyMode: alwaysoverwrite;
;Source: "ArquivosInstGeoSan\Instruções para configuração de atualizador automatico de versões.txt";          DestDir: "{app}";             CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\00. Leiame.txt";          DestDir: "{app}";             CopyMode: alwaysoverwrite;

Source: "ArquivosInstGeoSan\Exporte EPANet.exe";  DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\GeoSanIni.exe";       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\ValidaBase.exe";      DestDir: "{app}"; CopyMode: alwaysoverwrite;
; arquivos de atualização contínua do GeoSan
Source: "ArquivosInstGeoSan\GeoSan.exe";          DestDir: "c:\tempApp"; CopyMode: alwaysoverwrite;     
Source: "ArquivosInstGeoSan\GeoSanIni.exe";       DestDir: "c:\tempApp"; CopyMode: alwaysoverwrite;
Source: "ArquivosInstGeoSan\Updates.txt";         DestDir: "c:\tempApp"; CopyMode: alwaysoverwrite;
; dlls GeoSan
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



[Icons]
;Name: "{group}\My Program"; Filename: "{app}\MyProg.exe"
Name: "{commonprograms}\GeoSan"; Filename: "{app}\GeoSanIni.exe"
Name: "{commondesktop}\GeoSan";  Filename: "{app}\GeoSanIni.exe"
