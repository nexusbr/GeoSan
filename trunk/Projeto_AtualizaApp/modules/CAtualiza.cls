VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAtualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para atualizar tanto remoto para servidor quanto servidor para cliente
Private carrega As CDownload                                'para fazer o download de novas atualizações
' Para a leitura das configurações do usuário
'
'
'
Private Function LeConfiguracoes()
    On Error GoTo Trata_Erro:
    
    Set conf = New CArquivoIni                                                              'leitura e escrita no arquivo GeoSan.ini
    Set carrega = New CDownload
    conf.dirGeoSanIni = "C:\Arquivos de programas\GeoSan\Controles"
    carrega.atualizacaoWeb = conf.ReadINI("ATUALIZACAO", "WEB")                             'se a atualização será realizada pela web ou não
    carrega.diretorioServidor = conf.ReadINI("ATUALIZACAO", "DIRETORIO")                    '\download\GeoSan
    carrega.portaWww = conf.ReadINI("ATUALIZACAO", "PORTAWWW")                              'porta do servidor na web em que estão os arquivos
    carrega.url = conf.ReadINI("ATUALIZACAO", "URL")                                        'c:\tempFtp ou http://www.nexusbr.com
    carrega.proxyPorta = conf.ReadINI("ATUALIZACAO", "PROXYPORTA")
    carrega.proxy = conf.ReadINI("ATUALIZACAO", "PROXY")
    carrega.diretorioLocal = conf.ReadINI("ATUALIZACAO", "DIRETORIOLOCAL")                  'c:\tempApp
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CAtualiza", "LeConfiguracoes", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Function
' Compara os arquivos de um repositório remoto da NEXUS e atualiza no diretório local todos os arquivos do diretório remoto da NEXUS para o diretório local do servidor do cliente
'
' dirDownloadLocal - diretório local onde estão as últimas atualizações disponíveis para os usuários do GeoSan
'
Public Function AtualizaDirRemoto() As Boolean
    On Error GoTo Trata_Erro:
    Dim numeroAtualizacoes  As Integer                      'número total de atualizações a serem realizadas
    Dim nomeArquivo As String                               'nome do arquivo a ser atualizado
    Dim diretorio As String                                 'nome do drive e diretório onde o arquivo será atualizado (salvo)
    Dim versaoNova As String                                'numero da versão nova a ser atualizada
    Dim versaoLocal As String                               'numero da versão local do arquivo instalada no computador do usuário
    Dim dirInstalacaoLocal As String                        'diretório local onde será realizada a instalação da atualização
    Dim retorno As Boolean
    Dim i As Integer
    
    LeConfiguracoes                                         'Realiza as configurações de atualização do usuário
    If versao.ExisteArquivo(carrega.diretorioLocal & "\" & "Updates.txt") Then
        Kill (carrega.diretorioLocal & "\" & "Updates.txt")
    End If
    retorno = carrega.DownloadArquivo("Updates.txt")                                        'obtem a lista de atualizações disponíveis
    numeroAtualizacoes = versao.VerificaAtualizacoes(carrega.diretorioLocal & "\" & "Updates.txt")
    frmAtualiza.ProgressBar1.Max = numeroAtualizacoes + 1
    mensagem = mensagem & vbCrLf & vbCrLf & "www.nexusbr.com ---> " & carrega.diretorioLocal
    frmAtualiza.Text1 = mensagem
    For i = 0 To numeroAtualizacoes - 1                                                     'enquanto existirem atualizações para se fazer download
        frmAtualiza.ProgressBar1.Value = i + 1
        versao.SplitAtualizacoes i, nomeArquivo, diretorio, versaoNova, dirInstalacaoLocal  'obtem os dados detalhados das atualizações remotas que não foram baixadas ainda
        If versao.ExisteArquivo(carrega.diretorioLocal & "\" & nomeArquivo) Then                        'verifica se existe o arquivo localmente referente a atualização a ser realizada
            versaoLocal = versao.ObtemVersaoArquivo(carrega.diretorioLocal & "\" & nomeArquivo)
            mensagem = mensagem & vbCrLf & Str(i + 1) & " - " & carrega.diretorioLocal & "\" & nomeArquivo & " - " & versaoLocal & " - www.nexusbr.com\" & nomeArquivo & " - " & versaoNova
            frmAtualiza.Text1 = mensagem                                        'isto deveria estar fora da classe, pois inviabiliza a reutilização
            If versaoLocal <> versaoNova Then
                retorno = carrega.DownloadArquivo(nomeArquivo)                  'faz o download para o diretório local, da atualização
                If retorno = True Then
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - Download: " & nomeArquivo & " ---> " & carrega.diretorioLocal & " - substituindo o existente"
                Else
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - Não conseguiu realizar o download do arquivo: " & nomeArquivo & " para o diretório local: " & carrega.diretorioLocal
                End If
                frmAtualiza.Text1 = mensagem
            End If
        Else                                                                    'é um arquivo novo e não existe nada locamente no diretório do usuário
            retorno = carrega.DownloadArquivo(nomeArquivo)                      'faz o download para o diretório local, da atualização
            If retorno = True Then
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - Download: " & nomeArquivo & " ---> " & carrega.diretorioLocal & " - Arquivo novo adicionado"
                Else
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - Não conseguiu realizar o download do arquivo: " & nomeArquivo & " para o diretório local: " & carrega.diretorioLocal
            End If
            frmAtualiza.Text1 = mensagem
        End If
        nomeArquivo = ""
        diretorio = ""
        versaoNova = ""
        dirInstalacaoLocal = ""
    Next
    AtualizaDirRemoto = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        AtualizaDirRemoto = False
        ErroUsuario.Registra "CAtualiza", "AtualizaDirRemoto", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Function
' Atualiza a aplicação local do usuário depois que o diretório do servidor foi atualizado
'
'
'
Public Function AtualizaAplicacaoLocal() As Boolean
    On Error GoTo Trata_Erro:
    Dim numeroAtualizacoes  As Integer                      'número total de atualizações a serem realizadas
    Dim nomeArquivo As String                               'nome do arquivo a ser atualizado
    Dim diretorio As String                                 'nome do drive e diretório onde o arquivo será atualizado (salvo)
    Dim versaoNova As String                                'numero da versão nova a ser atualizada
    Dim dirInstalacaoLocal As String                        'diretório local onde será realizada a instalação da atualização
    Dim retorno As Boolean
    Dim i As Integer
    Dim fs As New Scripting.FileSystemObject                'para poder copiar arquivos
    
    Set fs = New Scripting.FileSystemObject                 'inicializa objeto de copia de arquivos
    numeroAtualizacoes = versao.VerificaAtualizacoes(carrega.diretorioLocal & "\" & "Updates.txt")
    mensagem = mensagem & vbCrLf & vbCrLf & carrega.diretorioLocal & " ---> c:\Arquivo de Programas\GeoSan"
    frmAtualiza.Text1 = mensagem
    For i = 0 To numeroAtualizacoes - 1                                                     'enquanto existirem atualizações para se fazer download
        versao.SplitAtualizacoes i, nomeArquivo, diretorio, versaoNova, dirInstalacaoLocal  'obtem os dados detalhados das atualizações remotas que não foram baixadas ainda
        If versao.ExisteArquivo(dirInstalacaoLocal & "\" & nomeArquivo) And versao.ExisteArquivo(carrega.diretorioLocal & "\" & nomeArquivo) Then                        'verifica se existem os arquivos localmente e remotamente (diretório no servidor local) referente a atualização a ser realizada
            versaoLocal = versao.ObtemVersaoArquivo(dirInstalacaoLocal & "\" & nomeArquivo)
            mensagem = mensagem & vbCrLf & Str(i + 1) & " - " & carrega.diretorioLocal & "\" & nomeArquivo & " - " & versaoNova & " - " & dirInstalacaoLocal & "\" & nomeArquivo & " - " & versaoLocal
            frmAtualiza.Text1 = mensagem
            If versaoLocal <> versaoNova Then
                If nomeArquivo <> "GeoSanIni.exe" Then                                      'não pode copiar sobre a aplicação que está rodando, terá que ser realizado mais tarde quando entrar no GeoSan
                    fs.CopyFile carrega.diretorioLocal & "\" & nomeArquivo, dirInstalacaoLocal & "\" & nomeArquivo, True
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - " & carrega.diretorioLocal & "\" & nomeArquivo & " ---> " & dirInstalacaoLocal & "\" & nomeArquivo & " - realizada copia para atualização"
                    frmAtualiza.Text1 = mensagem
                Else
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - " & carrega.diretorioLocal & "\" & nomeArquivo & " ---> " & dirInstalacaoLocal & "\" & nomeArquivo & " - não realizou a cópia, pois é a própria aplicação de atualização. Será realizada quando entrar no GeoSan"
                    frmAtualiza.Text1 = mensagem
                End If
            Else
                mensagem = mensagem & vbCrLf & Str(i + 1) & " - " & carrega.diretorioLocal & "\" & nomeArquivo & " ---> " & dirInstalacaoLocal & "\" & nomeArquivo & " - não foi preciso copiar, versão já estava atualizada"
                frmAtualiza.Text1 = mensagem
            End If
        Else                                                                                'é um arquivo novo e não existe nada locamente no diretório do usuário
                If versao.ExisteArquivo(carrega.diretorioLocal & "\" & nomeArquivo) Then    'verifica antes se foi baixado realmente do servidor o arquivo que será copiado para o diretório da aplicação do usuário
                    fs.CopyFile carrega.diretorioLocal & "\" & nomeArquivo, dirInstalacaoLocal & "\" & nomeArquivo, True
                    mensagem = mensagem & vbCrLf & Str(i + 1) & " - " & carrega.diretorioLocal & "\" & nomeArquivo & " ---> " & dirInstalacaoLocal & "\" & nomeArquivo & " - Novo arquivo incorporado"
                    frmAtualiza.Text1 = mensagem
                End If
        End If
        nomeArquivo = ""
        diretorio = ""
        versaoNova = ""
        dirInstalacaoLocal = ""
    Next
    AtualizaAplicacaoLocal = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        AtualizaAplicacaoLocal = False
        ErroUsuario.Registra "CAtualiza", "AtualizaAplicacaoLocal", CStr(Err.Number), CStr(Err.Description), True, True
    End If
End Function
