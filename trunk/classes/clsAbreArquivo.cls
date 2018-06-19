VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbreArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Esta classe é responsável por operações de abertura de arquivos com aplicações associadas as suas respectivas extensões
Option Explicit

'Declara função para iniciar uma aplicação associada com uma extensão de um documento sem saber o nome da aplicação associada
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long   'retorna um handle para o desktop

'Definição de constantes de erro para o a função ShellExecute
Const SW_SHOWNORMAL = 1
Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&

'Função para abrir um arquivo com a aplicação associada a extensão do mesmo
'
' DocName - nome do arquivo a ser aberto
'
Private Function StartDoc(DocName As String) As Long
          On Error GoTo Trata_Erro:
          Dim Scr_hDC As Long
          Scr_hDC = GetDesktopWindow()
          StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _
          "", "C:\", SW_SHOWNORMAL)
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "clsAbreArquivo", "StartDoc", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
'Subrotina para abrir um arquivo com a aplicação associada a extensão do mesmo.
'Esta função verifica se houve erro na abertura e mostra mensagem de erro que ocorreu
'
' nomeArquivo - nome do arquivo, podendo conter o número do disco (c:\ por exemplo) a ser aberto pela aplicação associada a extensão do mesmo.
'
Public Sub Abre(nomeArquivo As String)
    On Error GoTo Trata_Erro:
    Dim r As Long, msg As String
    r = StartDoc(nomeArquivo)         'chama o arquivo a ser aberto
    If r <= 32 Then
        'There was an error
        Select Case r
            Case SE_ERR_FNF
                msg = "GeoSan: O arquivo não foi encontrado"
            Case SE_ERR_PNF
                msg = "GeoSan: O diretório não foi encontrado"
            Case SE_ERR_ACCESSDENIED
                msg = "GeoSan: Acesso negado"
            Case SE_ERR_OOM
                msg = "GeoSan: Disco cheio"
            Case SE_ERR_DLLNOTFOUND
                msg = "GeoSan: DLL não encontrada"
            Case SE_ERR_SHARE
                msg = "GeoSan: Ocorreu uma violação de compartilhamento"
            Case SE_ERR_ASSOCINCOMPLETE
                msg = "GeoSan: Associação de arquivo incompleta ou inválida"
            Case SE_ERR_DDETIMEOUT
                msg = "GeoSan: Tempo limite de DDE atingido"
            Case SE_ERR_DDEFAIL
                msg = "GeoSan: Transação de DDE falhou"
            Case SE_ERR_DDEBUSY
                msg = "GeoSan: DDE em uso/ocupada"
            Case SE_ERR_NOASSOC
                msg = "GeoSan: Não existe associação para a extensão do arquivo"
            Case ERROR_BAD_FORMAT
                msg = "GeoSan: Erro de arquivo executável inválido ou erro na imagem do arquivo executável"
            Case Else
                msg = "GeoSan: Erro desconhecido"
        End Select
        MsgBox msg
    End If
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "clsAbreArquivo", "Abre", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

