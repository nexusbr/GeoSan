Attribute VB_Name = "Header"
'variáveis púbicas para vários módulos

Public Type GlobalVariables
    diretorioGeoSan As String                           'diretório onde encontra-se o GeoSan.exe que está rodando
    enviaEmails As Boolean                              'se é ou não para enviar mensagens de email quando ocorrer um erro
End Type

Public glo As GlobalVariables

'Salva as variáveis globais em um arquivo de controle
Public Sub SaveLoadGlobalData(filename As String, Save As Boolean)
    Dim filenum As Integer, isOpen As Boolean
    On Error GoTo Error_Handler
    filenum = FreeFile
    Open filename For Binary As filenum
    isOpen = True
    If Save Then
        Put #filenum, , glo
    Else
        Get #filenum, , glo
    End If
Error_Handler:
    If isOpen Then Close #filenum
End Sub
