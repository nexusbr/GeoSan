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
' Classe com opera��es em arquivos e diret�rios
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
' Fun��o para a sele��o de um diret�rio por parte do usu�rio
'
' Retorna a string contendo o nome do diret�rio completo
'
Public Function SelecionaDiretorio() As String
    Dim lpIDList As Long ' Declare Varibles
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = "Selecione o diret�rio onde ser�o exportados os arquivos"
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
' Fun��o para retornar o prefixo com a data e hora em que os arquivos foram gerados
'
' Retorna a data seguida da hora para ser incorporada no arquivo de exporta��o
'
Public Function prefixo() As String
    Dim dataHora As String
    
    dataHora = DateValue(Now) & "  " & TimeValue(Now)               'define o nome completo do prefixo do arquivo, inclu�ndo a data e hora em que o mesmo ser� gerado pela primeira vez
    dataHora = Replace(dataHora, "/", "-")                        'troca caractere / especial que n�o � aceito como parte do nome do arquivo
    dataHora = Replace(dataHora, ":", "-")                        'troca caractere : especial que n�o � aceito como parte do nome do arquivo
    dataHora = dataHora & " - "
    prefixo = dataHora
End Function