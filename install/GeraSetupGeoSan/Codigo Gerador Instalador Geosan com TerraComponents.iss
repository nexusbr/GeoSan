#define SetupBaseName   "SetupGeoSan-v."
#define AppVersionFile  "08.01.00"

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

[Run]
; add the Parameters, WorkingDir and StatusMsg as you wish, just keep here
; the conditional installation Check
; Referência: http://stackoverflow.com/questions/11137424/how-to-make-vcredist-x86-reinstall-only-if-not-yet-installed
Filename: "{tmp}\vcredist_x86.exe"; Check: VCRedistNeedsInstall; Parameters: "/passive /Q:a /c:""msiexec /qb /i vcredist.msi"" "; StatusMsg: Instalando componetes C++ VS2010 ...

[Code]
#IFDEF UNICODE
  #DEFINE AW "W"
#ELSE
  #DEFINE AW "A"
#ENDIF
type
  INSTALLSTATE = Longint;
const
  INSTALLSTATE_INVALIDARG = -2;  // An invalid parameter was passed to the function.
  INSTALLSTATE_UNKNOWN = -1;     // The product is neither advertised or installed.
  INSTALLSTATE_ADVERTISED = 1;   // The product is advertised but not installed.
  INSTALLSTATE_ABSENT = 2;       // The product is installed for a different user.
  INSTALLSTATE_DEFAULT = 5;      // The product is installed for the current user.

  VC_2005_REDIST_X86 = '{A49F249F-0C91-497F-86DF-B2585E8E76B7}';
  VC_2005_REDIST_X64 = '{6E8E85E8-CE4B-4FF5-91F7-04999C9FAE6A}';
  VC_2005_REDIST_IA64 = '{03ED71EA-F531-4927-AABD-1C31BCE8E187}';
  VC_2005_SP1_REDIST_X86 = '{7299052B-02A4-4627-81F2-1818DA5D550D}';
  VC_2005_SP1_REDIST_X64 = '{071C9B48-7C32-4621-A0AC-3F809523288F}';
  VC_2005_SP1_REDIST_IA64 = '{0F8FB34E-675E-42ED-850B-29D98C2ECE08}';
  VC_2005_SP1_ATL_SEC_UPD_REDIST_X86 = '{837B34E3-7C30-493C-8F6A-2B0F04E2912C}';
  VC_2005_SP1_ATL_SEC_UPD_REDIST_X64 = '{6CE5BAE9-D3CA-4B99-891A-1DC6C118A5FC}';
  VC_2005_SP1_ATL_SEC_UPD_REDIST_IA64 = '{85025851-A784-46D8-950D-05CB3CA43A13}';

  VC_2008_REDIST_X86 = '{FF66E9F6-83E7-3A3E-AF14-8DE9A809A6A4}';
  VC_2008_REDIST_X64 = '{350AA351-21FA-3270-8B7A-835434E766AD}';
  VC_2008_REDIST_IA64 = '{2B547B43-DB50-3139-9EBE-37D419E0F5FA}';
  VC_2008_SP1_REDIST_X86 = '{9A25302D-30C0-39D9-BD6F-21E6EC160475}';
  VC_2008_SP1_REDIST_X64 = '{8220EEFE-38CD-377E-8595-13398D740ACE}';
  VC_2008_SP1_REDIST_IA64 = '{5827ECE1-AEB0-328E-B813-6FC68622C1F9}';
  VC_2008_SP1_ATL_SEC_UPD_REDIST_X86 = '{1F1C2DFC-2D24-3E06-BCB8-725134ADF989}';
  VC_2008_SP1_ATL_SEC_UPD_REDIST_X64 = '{4B6C7001-C7D6-3710-913E-5BC23FCE91E6}';
  VC_2008_SP1_ATL_SEC_UPD_REDIST_IA64 = '{977AD349-C2A8-39DD-9273-285C08987C7B}';
  VC_2008_SP1_MFC_SEC_UPD_REDIST_X86 = '{9BE518E6-ECC6-35A9-88E4-87755C07200F}';
  VC_2008_SP1_MFC_SEC_UPD_REDIST_X64 = '{5FCE6D76-F5DC-37AB-B2B8-22AB8CEDB1D4}';
  VC_2008_SP1_MFC_SEC_UPD_REDIST_IA64 = '{515643D1-4E9E-342F-A75A-D1F16448DC04}';

  VC_2010_REDIST_X86 = '{196BB40D-1578-3D01-B289-BEFC77A11A1E}';
  VC_2010_REDIST_X64 = '{DA5E371C-6333-3D8A-93A4-6FD5B20BCC6E}';
  VC_2010_REDIST_IA64 = '{C1A35166-4301-38E9-BA67-02823AD72A1B}';
  VC_2010_SP1_REDIST_X86 = '{F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}';
  VC_2010_SP1_REDIST_X64 = '{1D8E6291-B0D5-35EC-8441-6616F567A0F7}';
  VC_2010_SP1_REDIST_IA64 = '{88C73C1C-2DE5-3B01-AFB8-B46EF4AB41CD}';

function MsiQueryProductState(szProduct: string): INSTALLSTATE;
external 'MsiQueryProductState{#AW}@msi.dll stdcall';

function VCVersionInstalled(const ProductID: string): Boolean;
begin
  Result := MsiQueryProductState(ProductID) = INSTALLSTATE_DEFAULT;
end;

function InitializeSetup: Boolean;
var
  S: string;
  State: INSTALLSTATE;
begin
  Result := True;
  State := MsiQueryProductState(VC_2010_SP1_REDIST_X86);
  case State of
    INSTALLSTATE_INVALIDARG: S := 'INSTALLSTATE_INVALIDARG: An invalid parameter was passed to the function.';
    INSTALLSTATE_UNKNOWN: S := 'INSTALLSTATE_UNKNOWN: The product is neither advertised or installed.';
    INSTALLSTATE_ADVERTISED: S := 'INSTALLSTATE_ADVERTISED: The product is advertised but not installed.';
    INSTALLSTATE_ABSENT: S := 'INSTALLSTATE_ABSENT: The product is installed for a different user.';
    INSTALLSTATE_DEFAULT: S := 'INSTALLSTATE_DEFAULT: The product is installed for the current user.';
  else
    S := IntToStr(State) + 'Unexpected result';
  end;
//  MsgBox(S, mbInformation, MB_OK);
end;

function VCRedistNeedsInstall: Boolean;
begin
  // here the Result must be True when you need to install your VCRedist
  // or False when you don't need to, so now it's upon you how you build
  // this statement, the following won't install your VC redist only when
  // the Visual C++ 2010 Redist (x86) and Visual C++ 2010 SP1 Redist(x86)
  // are installed for the current user
  Result := not (VCVersionInstalled(VC_2010_SP1_REDIST_X86) and 
    VCVersionInstalled(VC_2010_SP1_REDIST_X86));
end;

[Files]
; Instala vcredist_x86.exe - Microsoft Visual C++ 2010 x86 Redistributable 10.0.40219, necessário para as componentes do Visual Studio 2010 do TeCanvas
Source: "ArquivosInstGeoSan\vcredist_x86.exe"; DestDir: {tmp}; Flags: deleteafterinstall; 

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
; Microsoft Winsock Control DLL versão 6.00.81694
Source: "ArquivosInstGeoSan\MSWINSCK.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: regserver noregerror 32bit   

[Icons]
;Name: "{group}\My Program"; Filename: "{app}\MyProg.exe"
Name: "{commonprograms}\GeoSan"; Filename: "{app}\GeoSan.exe"
Name: "{commondesktop}\GeoSan";  Filename: "{app}\GeoSan.exe"
