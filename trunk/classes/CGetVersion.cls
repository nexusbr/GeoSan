VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGetVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Para obter a versão de um arquivo .exe
Option Explicit
Private linha() As String

'declarações para poder descobrir a versão de um arquivo executável
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Public Function ObtemVersao(dirNomeArquivo As String) As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    ObtemVersao = fso.GetFileVersion(dirNomeArquivo)
End Function
' Lê o arquivo com todas as atualizações disponíveis
' O formato do arquivo é:
'
' nome do arquivo a atualizar;diretorio local;versão
'
' VerificaAtualizacoes - número de atualizações a serem realizadas (número de linhas lidas)
' dirNomeArquivo - diretório completo e nome do arquivo que será lido
'
' vbCr is the carriage return (return to line beginning),
' vbLf is the line feed (go to next line)
' vbCrLf is the carriage return / line feed (similar to pressing Enter)
'
Public Function VerificaAtualizacoes(dirNomeArquivo As String) As Integer
    On Error GoTo Trata_Erro:
    Dim contador As Integer
    Dim arquivoLido As String
    Dim linhaLida As String

    Open dirNomeArquivo For Input As #1
    Input #1, arquivoLido
    Close #1
    linha() = Split(arquivoLido, vbLf)
    VerificaAtualizacoes = UBound(linha()) + 1
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CGetVersion", "VerificaAtualizacoes", CStr(Err.Number), CStr(Err.Description), True, True
        VerificaAtualizacoes = False
    End If
End Function
' Dada uma linha lida no arquivo remoto de atualizações a serem realizada, faz um parse para ver o que tem que realizar no computador local para uma determinada atualização
' Segue o formato:
'
' nome do arquivo a atualizar;diretorio local;versão;diretório de instalação
'
' atualizacao - string contendo todos os parâmetros separados por ;
' nomeArquivo - nome do arquivo a ser atualizado
' diretorio - localização do mesmo no computador cliente
' versao - número da nova versão para comparar com a existente e ver se necessita atualizar
' dirInstalacao - diretório local onde a atualização deverá ser realizada para o cliente
'
Public Sub SplitAtualizacoes(numeroLinha As Integer, ByRef nomeArquivo As String, ByRef diretorio As String, ByRef versao As String, ByRef dirInstalacao As String)
    On Error GoTo Trata_Erro:
    Dim splitAtualizacao() As String
    
    splitAtualizacao = Split(linha(numeroLinha), ";")
    nomeArquivo = splitAtualizacao(0)
    diretorio = splitAtualizacao(1)
    versao = splitAtualizacao(2)
    dirInstalacao = splitAtualizacao(3)
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CGetVersion", "SplitAtualizacoes", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Sub


Public Function ExisteArquivo(sFile As String) As Boolean
    On Error Resume Next
    ExisteArquivo = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
' Obtem o número da versão de um determinado arquivo
'
' FileName - nome completo do arquivo com o drive e diretório
'
Public Function ObtemVersaoArquivo(ByVal FileName As String) As String
   On Error GoTo Trata_Erro:
   Dim nDummy       As Long
   Dim sBuffer()    As Byte
   Dim nBufferLen   As Long
   Dim lplpBuffer   As Long
   Dim udtVerBuffer As VS_FIXEDFILEINFO
   Dim puLen        As Long
      
   nBufferLen = GetFileVersionInfoSize(FileName, nDummy)
   If nBufferLen > 0 Then
        ReDim sBuffer(nBufferLen) As Byte
        Call GetFileVersionInfo(FileName, 0&, nBufferLen, sBuffer(0))
        Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
        Call CopyMemory(udtVerBuffer, ByVal lplpBuffer, Len(udtVerBuffer))
        ObtemVersaoArquivo = udtVerBuffer.dwFileVersionMSh & "." & udtVerBuffer.dwFileVersionMSl & "." & udtVerBuffer.dwFileVersionLSh & "." & udtVerBuffer.dwFileVersionLSl
    End If
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CGetVersion", "ObtemVersaoArquivo", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Function
