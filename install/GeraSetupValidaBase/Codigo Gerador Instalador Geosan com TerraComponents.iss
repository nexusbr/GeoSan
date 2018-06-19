#define SetupBaseName   "SetupValidaBase-v."
#define AppVersionFile  "07.00.00"

[Setup]
AppName=ValidaBase
Compression=lzma
AllowNoIcons=no
AlwaysRestart=no
AlwaysShowComponentsList=yes
; equivale ao Product Version em propriedades do Setup gerado
AppVerName=ValidaBase {#AppVersionFile}
; equivale ao file version em propriedades do Setup gerado
VersionInfoVersion={#AppVersionFile} 
VersionInfoTextVersion=ValidaBase {#AppVersionFile}
AppCopyright=NEXUS GeoEngenharia

DefaultDirName=C:\Arquivos de Programas\GeoSan
DefaultGroupName=ValidaBase
UninstallDisplayIcon={app}\ValidaBase.exe
AppMutex=ValidaBase

SolidCompression=yes
OutputDir=Output
OutputBaseFilename={#SetupBaseName + AppVersionFile}

[Language]
MessagesFile=compiler:BrazilianPortuguese.isl

[Files]

; CopyMode: alwaysskipifsameorolder para manter caso não exista diferença de versão
; CopyMode: alwaysoverwrite para sobrescrever tudo

;ARQUIVOS NEXUS - *************************************************************************************************************************
Source: "ArquivosInstValidaBase\ValidaBase.exe";          DestDir: "{app}";             CopyMode: alwaysoverwrite;

;ARQUIVOS VISUAL BASIC 6 - ****************************************************************************************************************


[Icons]
Name: "{commonprograms}\ValidaBase"; Filename: "{app}\ValidaBase.exe"
Name: "{commondesktop}\ValidaBase";  Filename: "{app}\ValidaBase.exe"
