VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArquivoIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FUNÇÕES PARA LER E GRAVAR NO ARQUIVO .INI-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Funções de leitura do arquivo de inicialização do GeoSan


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private dirConfigGeoSan As String                       'nome do diretório onde está localizado o arquivo de configurações iniciais do GeoSan, GEOSAN.INI

Property Let dirGeoSanIni(dirG As String)
    dirConfigGeoSan = dirG & "\GEOSAN.INI"
End Property
' Lê o arquivo de inicialização do GeoSan e retorna o parâmetro solicitado do mesmo
'
' Secao - o nome da seção presente no arquivo .ini entre colchetes []
' Entrada - nome do parâmetro ao qual se deseja obter a informação de entrada, o qual está dentro da seção apontada. Fica antes do sinal de igual
'
Public Function ReadINI(Secao As String, Entrada As String) As String
    On Error GoTo Trata_Erro:
    Dim retlen As String
    Dim Ret As String
    
    Ret = String$(255, 0)                                                           'string que conterá o parâmetro de retorno. Preenche com o caractere ASCII 0 255 vezes
    retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), dirConfigGeoSan)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CArquivoIni", "ReadINI", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
' Escreve no arquivo de inicialização
'
' Secao - O que esta entre []
' Entrada - nome do que se encontra antes do sinal de igual
' texto - valor que vem depois do igual
'
Public Sub WriteINI(Secao As String, Entrada As String, Texto As String)
    WritePrivateProfileString Secao, Entrada, Texto, dirConfigGeoSan
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CArquivoIni", "WriteINI", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
