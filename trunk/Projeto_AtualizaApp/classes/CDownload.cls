VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private b() As Byte                             'contem o arquivo em que será realizado o download
Private nomeUrl As String                       'caminho em que está o arquivo em que será realizado o download
Private proxyNome As String                     'nome do servidor Proxy da empresa, caso exista
Private proxyUsuario As String                  'login do usuário para acessar o Proxy interno da empresa, caso exista
Private proxySenha As String                    'senha do usuário para acessar o Proxy interno da empresa, caso exista
Private proxyPort As String                     'número da porta do servidor Proxy interno da empresa
Private diretorioComputador As String                'nome do diretório local onde será realizado o download da atualização
Private nomeArquivoAtualizacoes As String       'nome do arquivo remoto que contem as atualizações a serem realizadas
Private diretorioRemoto As String               'nome do diretório remoto onde estão todas as atualizações


Property Let url(nomeU As String)               'recebe o caminho em que está o arquivo em que será realizado o download
    nomeUrl = nomeU
End Property

Property Let proxy(proxyN As String)            'recebe o nome da porta do proxy
    proxyNome = proxyN
End Property

Property Let proxyUser(proxyU As String)        'recebe o nome do usuário que acessa o proxy
    proxyUsuario = proxyU
End Property

Property Let proxyPass(proxyP As String)        'recebe a senha do usuário que acessa o proxy
    proxySenha = proxyP
End Property

Property Let proxyPorta(proxyP As String)       'recebe o número da porta do proxy
    proxyPort = proxyP
End Property
Property Let diretorioLocal(diretorioL As String)   'nome do diretório local onde serão baixadas todas as atualizações
    diretorioComputador = diretorioL
End Property

Property Let diretorioServidor(diretorioS As String) 'nome do diretório remoto onde estão todas as atualizações
    diretorioRemoto = diretorioS
End Property
Public Function DownloadArquivo(arquivoDownload As String) As Boolean
    On Error GoTo Trata_Erro:
    Dim url As String
    Dim nomeDownloadLocal As String
    
    Set ErroUsuario = New CPrintErro
    url = nomeUrl & ":" & proxyPort & diretorioRemoto & "/" & arquivoDownload
    frmAtualiza.Inet1.url = url
    b() = frmAtualiza.Inet1.OpenURL(url, icByteArray)
    nomeDownloadLocal = diretorioComputador & "\" & arquivoDownload
    Open nomeDownloadLocal For Binary Access Write As #1
    Put #1, , b()
    Close #1
    DownloadArquivo = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CDownload", "DownloadArquivo", CStr(Err.Number), CStr(Err.Description), False, True
        DownloadArquivo = False
    End If
End Function

Public Function VerificaAtualizacoes() As Boolean

End Function

