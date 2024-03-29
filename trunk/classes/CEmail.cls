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
' Classe respons�vel pelo envio de emails
'
'
Private emailDestinatario As String             'email de quem vai receber. p.ex. tec@nexusbr.com
Private emailRemetente As String                'nome de quem est� enviando. p.ex. NEXUS <tec@nexusbr.com>
Private enderecoSMTP As String                  'endere�o smtp de envio. p.ex. para o GMail: smtp.gmail.com
Private portaSMTP As Integer                    'porta de envio do email. p. ex. para o GMail: 465
Private usuarioSMTP As String                   'usu�rio de login do smtp. p. ex. para o GMail: tec@nexusbr.com
Private senhaSMTP As String                     'senha do sSmtpUser para acessar
Private SslSmtp As Boolean                      'se � para enviar com seguran�a ou n�o. p. ex. para o GMail: True

' Envia um email gen�rio
'
' sTo - email de quem vai receber. p.ex. tec@nexusbr.com
' sSubject - texto com o assunto do email
' sFrom - nome de quem est� enviando. p.ex. NEXUS <tec@nexusbr.com>
' sBody - texto com o corpo do email
' sSmtpServer - endere�o smtp de envio. p.ex. para o GMail: smtp.gmail.com
' iSmtpPort - porta de envio do email. p. ex. para o GMail: 465
' sSmtpUser - usu�rio de login do smtp. p. ex. para o GMail: tec@nexusbr.com
' sSmtpPword - senha do sSmtpUser para acessar
' sFilePath - caminho do arquivo que ser� anexado
' bSmtpSSL - se � para enviar com seguran�a ou n�o. p. ex. para o GMail: True
'
' Obs.: Ao debugar pode acontecer se estiver indo linha a linha de ele n�o enviar o email
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
    retval = SendMail(emailDestinatario, "GeoSan - Aviso de erro", emailRemetente, textoErro, enderecoSMTP, portaSMTP, usuarioSMTP, senhaSMTP, arquivoParaAnexar, SslSmtp)
    enviaEmail = retval
End Function
' M�todo para ler as configura��es de envio de email
'
'
Public Function leConfiguracoesEmail()
    Dim caminhoGeoSanIni As String
    
    caminhoGeoSanIni = App.path & "\CONTROLES\GEOSAN.INI"
    emailDestinatario = ReadINI("EMAIL", "EMAILDESTINATARIO", caminhoGeoSanIni)
    emailRemetente = ReadINI("EMAIL", "EMAILREMETENTE", caminhoGeoSanIni)
    enderecoSMTP = ReadINI("EMAIL", "ENDERECOSMTP", caminhoGeoSanIni)
    portaSMTP = CInt(IIf(ReadINI("EMAIL", "PORTASMTP", caminhoGeoSanIni) = "", 0, ReadINI("EMAIL", "PORTASMTP", caminhoGeoSanIni)))     'caso n�o encontre retorna a porta zero
    usuarioSMTP = ReadINI("EMAIL", "USUARIOSMTP", caminhoGeoSanIni)
    senhaSMTP = ReadINI("EMAIL", "SENHASMTP", caminhoGeoSanIni)
    SslSmtp = CBool(IIf(ReadINI("EMAIL", "SSLSMTP", caminhoGeoSanIni) = "", "TRUE", ReadINI("EMAIL", "SSLSMTP", caminhoGeoSanIni)))     'caso exista erro na leitura da seguran�a ssl, define como padr�o TRUE, sen�o o que foi entrado no arquivo .ini
    If emailDestinatario = "" Or emailRemetente = "" Or enderecoSMTP = "" Or portaSMTP = 0 Or usuarioSMTP = "" Or senhaSMTP = "" Then   'define como padr�o o email da NEXUS, pois n�o consegui ler corretamente no arquivo .ini
        emailDestinatario = "tec@nexusbr.com"
        emailRemetente = "NEXUS <tec@nexusbr.com>"
        enderecoSMTP = "smtp.gmail.com"
        portaSMTP = 465
        usuarioSMTP = "tec@nexusbr.com"
        senhaSMTP = "nexus243"
        SslSmtp = True
    Else
        'conseguiu ler com sucesso todas informa��es e est� pronto para enviar emails caso ocorra um erro na aplica��o
    End If
End Function
