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
'Atualiza a aplicação GeoSanIni.exe quando necessário

' Atualiza a aplicação local do usuário depois que o diretório do servidor foi atualizado
'
' AtualizaAplicacaoLocal - retorna verdadeiro se atualizou com sucesso e falso se ocorreu algum problema na atualização
'
Public Function AtualizaAplicacaoLocal() As Boolean
    On Error GoTo Trata_Erro:
    Dim versaoAtu As CGetVersion        'para obter a versão da aplicação
    Dim mensagemAtu As Boolean
    Dim diretorioServLocal As String    'diretório local onde estão todas as atualizações baixadas da web p.ex: c:\tempApp
    Dim versaoAplicacao As String       'número da versão atual da aplicação GeoSanIni.exe
    Dim versaoServidor As String        'número da versão que foi atualizada e está disponível
    Dim retornoAtu As Boolean           'se a cópia foi realizada com sucesso ou não
    
    Set versaoAtu = New CGetVersion
    diretorioServLocal = ReadINI("ATUALIZACAO", "DIRETORIOLOCAL", App.path & "\CONTROLES\GEOSAN.INI")
    If versaoAtu.ExisteArquivo(App.path & "\GeoSanIni.exe") Then
        versaoAplicacao = versaoAtu.ObtemVersao(App.path & "\GeoSanIni.exe")
    Else
        ErroUsuario.Registra "GeoSan.exe-CAtualiza", "AtualizaAplicacaoLocal", CStr(Err.Number), CStr(Err.Description), False, True, "Não encontrou o arquivo: " & App.path & "\GeoSanIni.exe - A aplicação continuou rodando normalmente."
        AtualizaAplicacaoLocal = False
        Exit Function
    End If
    If versaoAtu.ExisteArquivo(diretorioServLocal & "\GeoSanIni.exe") Then
        versaoServidor = versaoAtu.ObtemVersao(diretorioServLocal & "\GeoSanIni.exe")
    Else
        ErroUsuario.Registra "GeoSan.exe-CAtualiza", "AtualizaAplicacaoLocal", CStr(Err.Number), CStr(Err.Description), False, True, "Não encontrou o arquivo: " & diretorioServLocal & "\GeoSanIni.exe - A aplicação continuou rodando normalmente."
        AtualizaAplicacaoLocal = False
        Exit Function
    End If
    If versaoAplicacao <> versaoServidor Then
        FileCopy diretorioServLocal & "\GeoSanIni.exe", App.path & "\GeoSanIni.exe"
        ErroUsuario.Registra "GeoSan.exe-CAtualiza", "AtualizaAplicacaoLocal", CStr(Err.Number), CStr(Err.Description), False, True, "Copiou o arquivo " & diretorioServLocal & "\GeoSanIni.exe ---> " & App.path & "\GeoSanIni.exe - Atualizando o mesmo com sucesso."
    End If
    AtualizaAplicacaoLocal = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "GeoSan.exe-CAtualiza", "AtualizaAplicacaoLocal", CStr(Err.Number), CStr(Err.Description), False, glo.enviaEmails
    End If
End Function
