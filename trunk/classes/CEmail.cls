VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe responsável pelo envio de emails
'
'
' Envia um email genério
'
' sTo - email de quem vai receber. p.ex. tec@nexusbr.com
' sSubject - texto com o assunto do email
' sFrom - nome de quem está enviando. p.ex. NEXUS <tec@nexusbr.com>
' sBody - texto com o corpo do email
' sSmtpServer - endereço smtp de envio. p.ex. para o GMail: smtp.gmail.com
' iSmtpPort - porta de envio do email. p. ex. para o GMail: 465
' sSmtpUser - usuário de login do smtp. p. ex. para o GMail: tec@nexusbr.com
' sSmtpPword - senha do sSmtpUser para acessar
' sFilePath - caminho do arquivo que será anexado
' bSmtpSSL - se é para enviar com segurança ou não. p. ex. para o GMail: True
'
' Obs.: Ao debugar pode acontecer se estiver indo linha a linha de ele não enviar o email
'
Private Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.message             'Entrar em Project References e incluir a biblioteca cdosys.dll
    Set lobj_cdomsg = New CDO.message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function
' Metodo para enviar uma mensagem de erro que ocorreu no GeoSan
'
' arquivoParaEnviar - nome do arquivo completo a ser enviado em anexo ao email. p. ex. C:\user\pinheiro\geosanLogErro.txt
'
Public Function enviaEmail(arquivoParaAnexar As String) 'As String
    Dim retval As String
    Dim textoErro As String
    
    textoErro = "Segue em anexo o arquivo de log do erro ocorrido no GeoSan."
    retval = SendMail("tec@nexusbr.com", "GeoSan - Aviso de erro", "NEXUS <tec@nexusbr.com>", textoErro, "smtp.gmail.com", 465, "tec@nexusbr.com", "nexus243", arquivoParaAnexar, True)
    enviaEmail = retval
End Function


