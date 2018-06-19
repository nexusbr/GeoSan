VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAcertaZsDosNos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe responsável por acertar todos os z's dos nós das redes de água
'
'
'

Private getZNo As New CObtemZdoNo
' Subrotina que atribui a cota Z a todos os nós das redes
'
'
'
Public Sub AtribuiZs()
    On Error GoTo Trata_Erro
    Dim RSNos As ADODB.Recordset
    Dim rsAtributoNos As ADODB.Recordset
    Dim xNo As Double
    Dim yNo As Double
    Dim objIdNo As String
    Dim zNo As Double
    Dim diretorio As String                                                 'Diretório onde será salva a exportação das cotas calculadas para conferir a anterior com a nova
    Dim retorno As Boolean
    Dim arquivoExportacao As String                                         'nome do arquivo junto com o diretório onde serão exportadas as cotas para conferência
    Dim zNoStr As String                                                    'coordenada Z do no convertida para string. Isso para tratar a virgula que o SQL nao permite no UPDATE
    
    diretorio = arquivo.SelecionaDiretorio
    If diretorio = "falhou" Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set RSNos = New ADODB.Recordset
    RSNos.Open "Select x, y, object_id from Points2", Conn, adOpenKeyset, adLockOptimistic
    arquivoExportacao = diretorio & "\GeoSan_atualização_cotas.txt"
    Open arquivoExportacao For Output As #10                                                                            'Abre o arquivo novo
    Close #10
    retorno = cGeoDatabase.geoDatabase.setCurrentLayer("mdt")                                                       'ativa o layer do MDT
    If retorno = False Then
        Screen.MousePointer = vbNormal
        MsgBox ("É necesário antes de iniciar esta operação abrir uma vista do mapa.")
        Exit Sub
    End If
    Do While RSNos.EOF = False
        DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            varGlobais.pararExecucao = False
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        xNo = RSNos("x").value
        yNo = RSNos("y").value
        objIdNo = RSNos("object_id").value
        zNo = getZNo.ObtemZ(xNo, yNo, objIdNo)
        RSNos.MoveNext
        Set rsAtributoNos = New ADODB.Recordset
        rsAtributoNos.Open "SELECT GROUNDHEIGHT FROM WATERCOMPONENTS WHERE OBJECT_ID_ = " & objIdNo, Conn, adOpenKeyset, adLockOptimistic
        Open arquivoExportacao For Append As #10
        Print #10, objIdNo & " , " & xNo & " , " & yNo & " , " & rsAtributoNos.Fields("GROUNDHEIGHT").value & " , " & CStr(zNo)
        Close #10
        FrmMain.sbStatusBar.Panels(2).Text = "Z nó " & objIdNo & " calculado = " & zNo
        rsAtributoNos.Close
        zNoStr = CStr(zNo)                          'converte a cota do no para string para poder atualizar no SQL como ponto
        zNoStr = Replace(zNoStr, ",", ".")          'troca a vírgula por ponto
        Conn.execute "UPDATE WATERCOMPONENTS SET GROUNDHEIGHT = " & zNoStr & " WHERE OBJECT_ID_ = " & objIdNo
    Loop
    Close #10
    Screen.MousePointer = vbNormalObtemZ
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Screen.MousePointer = vbNormal
        ErroUsuario.Registra "CAcertaZsDosNos", "AtribuiZs", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub


