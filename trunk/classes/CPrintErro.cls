VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrintErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para informar o erro para o usuário e registrar o mesmo no arquivo de log
' Recebe os dados do erro, mostra mensagem ao usuário e registra no arquivo de log do sistema
'
' Modulo - string que contém em arquivo VB o erro ocorreu
' EVENTO - string que contém em que rotina o erro ocorreu
' ErrDescr - string com a descrição do erro ocorrido
' ExibeMensagem - se é para exibir ou não uma mensagem para o usuário
' enviaEmail - verdadeiro ou falso, indicando se irá enviar ou não um email para NEXUS com o erro que ocorreu
' linha - número da linha ou descrição da localização em que o erro ocorreu
'
' Exemplo de uso desta classe: ErroUsuario CStr(Me.Name), "nome do evento que o mesmo ocorreu, através de mensagem digitada pelo programador", CStr(Err.Number), CStr(Err.Description), True, True
' Códigos de erros a serem tratados:
' 0  - The operation completed successfully
' 20 - The system cannot find the device specified
'
' Exemplo de código para tratar o erro
'
'   Exit Function
'   Exit Sub
'Trata_Erro:
'    If Err.Number = 0 Or Err.Number = 20 Then
'        Resume Next
'    Else
'       ErroUsuario.Registra "CCoordIniFinLinha", "GetStartEndPointsTrecho", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
'    End If
'
Public Function Registra(ByVal Modulo As String, ByVal EVENTO As String, ByVal ErrNum As String, ByVal ErrDescr As String, ByVal ExibeMensagem As Boolean, ByVal enviaEmail As Boolean, Optional ByVal linha As String = "Não registrada")
    Dim nomeArquivoErro As String
    Dim sucesso As String
    
    nomeArquivoErro = Environ("HOMEDRIVE") & Environ("HOMEPATH")
    'nomeArquivoErro = Environ$("USERPROFILE") & "\Meus Documentos" & Environ("LOCALAPPDATA")
    nomeArquivoErro = nomeArquivoErro + "\GeoSanLogErro.txt"
    Close #1 'FECHA O ARQUIVO DE LOG
    Open nomeArquivoErro For Append As #1
    Print #1, "DATA"; Tab(16); Now
    Print #1, "USUÁRIO"; Tab(16); strUser
    Print #1, "VERSÃO"; Tab(16); Versao_Geo
    Print #1, "ARQUIVO"; Tab(16); Modulo
    Print #1, "ROTINA"; Tab(16); EVENTO
    Print #1, "LINHA"; Tab(16); linha
    Print #1, "CÓDIGO ERRO"; Tab(16); ErrNum
    Print #1, "DESCRIÇÃO"; Tab(16); ErrDescr
    Print #1, ""
    Print #1, "-----------------------------------------------------------------------------------------------------"
    Print #1, ""
    Close #1 'FECHA O ARQUIVO
    'SE O PARÂMETRO ExibeMensagem = True , EXIBE MENSAGEM PARA O USUÁRIO
    If enviaEmail = True Then
        sucesso = Email.enviaEmail(nomeArquivoErro)
        If sucesso <> "ok" Then
            MsgBox "Não conseguiu enviar a mensagem do erro ocorrido para a NEXUS. Mensagem: " & sucesso & " Favor entrar em contato com o suporte técnico da NEXUS."
        End If
    End If
    If ExibeMensagem = True Then
        MsgBox "A operação não pode ser completada, favor enviar o arquivo: " & nomeArquivoErro & " para o suporte técnico para maiores detalhes.", vbInformation
    End If
End Function
