VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe com operações em arquivos e diretórios
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

'Declarações necessárias para a função GetMyDocumentsDirectory()
Const REG_SZ = 1
Const REG_BINARY = 3
Const HKEY_CURRENT_USER = &H80000001
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = &H20000
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
    ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' Sugestão de rotina para o usuário selecionar um arquivo em que serão salvas informações
'    Dim selecionaArquivo As New CArquivo                                                    'para obter o nome e diretório onde o arquivo será salvo
'    Dim nomeArquivo As String                                                               'nome completo do arquivo com o drive e diretório no qual será salvo
'    Dim diretorioMyDocuments As String                                                      'diretório meus documentos inicial
'    Dim filelocation As String                                                              'nome completo do arquivo onde será salvo o relatório
'
'    nomeArquivo = selecionaArquivo.ConfiguraNomeArquivo("exportação_dados_redes", "txt", diretorioMyDocuments)  'obtem o nome do arquivo sugerido e o diretório meus documentos do usuário
'    CommonDialog1.Filter = "Texto (.txt)|*.TXT|Todos tipos (*.*)|*.*|"                                          'configura o filtro do arquivo
'    CommonDialog1.filename = nomeArquivo                                                                        'informa a caixa de diálogo que será aberta o nome do arquivo inicial sugerido
'    CommonDialog1.InitDir = diretorioMyDocuments                                                                'sugero o diretório inicial
'    CommonDialog1.ShowSave                                                                                      'abre a caixa de diálogo par ao usuário digitar o nome do arquivo e selecionar o diretório, se desejar
'    filelocation = CommonDialog1.filename

' Função para definir previamente um nome de arquivo
'
' Retorna um diretório junto com o nome do arquivo
' nomeArquivo - nome do arquivo que será concatenado junto ao restante das informações
' extensao - extensao do arquivo sem o ponto
' diretoiro - em que o arquivo será salvo. Retorna este valor
'
Public Function ConfiguraNomeArquivo(nomeArquivo As String, extensao As String, ByRef diretorio As String) As String
    diretorio = GetMyDocumentsDirectory()
    ConfiguraNomeArquivo = Format(Now, "YYYY-MM-DD-HHMMSS") & "-" & nomeArquivo & "." & extensao
End Function


' Função para a seleção de um diretório por parte do usuário
'
' Retorna a string contendo o nome do diretório completo
'
Public Function SelecionaDiretorio() As String
    Dim lpIDList As Long ' Declare Varibles
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = "Selecione o diretório onde serão exportados os arquivos"
    ' Text to appear in the the gray area under the title bar
    ' telling you what to do
    
    With tBrowseInfo
        .hWndOwner = FrmMain.hwnd ' Owner Form
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        sBuffer = "falhou"
    End If

    SelecionaDiretorio = sBuffer
End Function
' Função para retornar o prefixo com a data e hora em que os arquivos foram gerados
'
' Retorna a data seguida da hora para ser incorporada no arquivo de exportação
'
Public Function prefixo() As String
    Dim dataHora As String
    
    dataHora = DateValue(Now) & " " & TimeValue(Now)               'define o nome completo do prefixo do arquivo, incluíndo a data e hora em que o mesmo será gerado pela primeira vez
    dataHora = Replace(dataHora, "/", "-")                        'troca caractere / especial que não é aceito como parte do nome do arquivo
    dataHora = Replace(dataHora, ":", "-")                        'troca caractere : especial que não é aceito como parte do nome do arquivo
    dataHora = dataHora + " - "
    prefixo = dataHora
End Function
'Obtem o nome do diretório dos Meus Documentos do usuário que está logado
'
'GetMyDocumentsDirectory() - retorna o caminho do diretório
'
Private Function GetMyDocumentsDirectory() As String
    Dim lRes As Long
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    Dim strData As Integer
    RegOpenKeyEx HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, lRes
    lResult = RegQueryValueEx(lRes, "Personal", 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(lRes, "Personal", 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                GetMyDocumentsDirectory = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        End If
    End If
    RegCloseKey lRes
End Function
