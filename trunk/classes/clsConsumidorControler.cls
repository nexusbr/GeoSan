VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsumidorControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvartcs As TeCanvas 'local copy
Private mvarTerraEvent As TypeGeometryEvent 'local copy
Private mvartdbcon As TeDatabase 'local copy
Private mvarCGeo As clsGeoReference 'local copy

' Atribui localmente o objeto TeCanvas da OCX TeCanvas para poder dentro da classe executar
' os métodos disponíveis no mesmo
'
Public Property Set tcs(ByVal vData As TeCanvas)
    Set mvartcs = vData
    
End Property
' Recebe o objeto TeCanvas da OCX TeCanvas para poder dentro da classe executar
' os métodos disponíveis no mesmo
'
Public Property Get tcs() As TeCanvas
    Set tcs = mvartcs
End Property
' Recebe o tipo de evento que está ocorrendo para a classe cConsumidor e salva internamente
' este evento na classe para saber depois o que está acontecendo
'
Public Property Let TerraEvent(ByVal vData As TypeGeometryEvent)
    mvarTerraEvent = vData
End Property

Public Property Get TerraEvent() As TypeGeometryEvent
    TerraEvent = mvarTerraEvent
End Property

Public Property Set tdbcon(ByVal vData As TeDatabase)
    Set mvartdbcon = vData
    'Set cgeo.tdbcon = vData
End Property

Public Property Get tdbcon() As TeDatabase
    Set tdbcon = mvartdbcon
End Property

Public Property Set cgeo(ByVal vData As clsGeoReference)
    Set mvarCGeo = vData
End Property

Public Property Get cgeo() As clsGeoReference
    Set cgeo = mvarCGeo
End Property
'Cadastra a linha do ramal da ligação de água
'
Public Function Move()
    tcs.setCurrentLayer ("NXGS_V_LIG_COMERCIAL_GPS")
    tcs.moveGeometries (0)
End Function
Public Sub Init()
    
End Sub
'Insere um texto
'
Public Function InsereTexto(objIdPontoGPS As String)
         Dim coordXcentroGeometria As Double, coordYcentroGeometria As Double
         Dim gpsConsumidorSql As String
         Dim numeroCasa As String
         
         numeroCasa = ObtemNumeroCasa(objIdPontoGPS)
         SaveInDatabase
         tdbcon.setCurrentLayer "NXGS_V_LIG_COMERCIAL_GPS"                                                                      'definie que o layer que estamos trabalhando é o que possui os pontos das coordenadas GPS
         gpsConsumidorSql = "Delete From Texts" & GetLayerID("NXGS_V_LIG_COMERCIAL_GPS") & "  Where Object_id = '" & objIdPontoGPS & "'"
         Conn.execute gpsConsumidorSql
         tdbcon.getCenterGeometry CLng(objIdPontoGPS), objIdPontoGPS, tpPOINTS, coordXcentroGeometria, coordYcentroGeometria    'localiza as coordenadas X e Y do ponto GPS selecionado
         tdbcon.InsertTextInDatabase objIdPontoGPS, numeroCasa, 1, coordXcentroGeometria, coordYcentroGeometria, 0, TeHorzLeft, TeHorzLeft, True
         gpsConsumidorSql = "UPDATE NXGS_V_LIG_COMERCIAL_GPS SET USUARIO_LOG = '" & strUser & "', DATA_LOG = '" & Format(Now, "DD/MM/YY HH:MM") & "', DATALOG = GETDATE() WHERE " & ObtemNomeColunaObjId("NXGS_V_LIG_COMERCIAL_GPS") & " = '" & objIdPontoGPS & "'"                  'atualiza quem realizou a última modificação na posição do consumidor
         Conn.execute (gpsConsumidorSql)
         tcs.plotView
End Function

' Função chamada para salvar no banco de dados
'
Public Function SaveInDatabase() As Boolean
    On Error GoTo Trata_Erro
    Dim USR As String
    
    With tcs
        Select Case TerraEvent                          'A variavel TerraEvent determina o evento/funcão que o usuário está utilizando
            Case tg_MoveGpsPoint
                tcs.saveOnMemory
                tcs.SaveInDatabase
                Screen.MousePointer = vbHourglass
                Screen.MousePointer = vbNormal
        End Select
    End With
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf mid(Err.Description, 1, 9) = "ORA-03114" Then
        PrintErro "clsConsumidor", "Public Function SaveInDatabase()", CStr(Err.Number), CStr(Err.Description), True
        End
    Else
        PrintErro "clsConsumidor", "Public Function SaveInDatabase()", CStr(Err.Number), CStr(Err.Description), True
    End If
    Screen.MousePointer = vbNormal
End Function

' Obtem a identificação do layer
' Retorna o número de identificação do layer
'
' LayerName_ - nome do layer
'
Private Function GetLayerID(LayerName_ As String) As Integer
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset

    Set rs = Conn.execute(convertQuery("SELECT Layer_id from Te_Layer where name='" & LayerName_ & "' or name='" & UCase(LayerName_) & "'", CInt(typeconnection)))
    If rs.EOF = False Then
        GetLayerID = rs(0).value            'obtem o número do layer (id)
    End If
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsConsumidorControler", "Public Function GetLayerID", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

'Obtem o número da casa do consumidor levantado em campo com o coletor do software comercial
'
'
Private Function ObtemNumeroCasa(objIDLigacao As String) As String
On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim consumidorGpsSql As String
    
    consumidorGpsSql = "SELECT NRO_CASA from NXGS_V_LIG_COMERCIAL_GPS where " & ObtemNomeColunaObjId("NXGS_V_LIG_COMERCIAL_GPS") & " = '" & objIDLigacao & "'"
    Set rs = Conn.execute(consumidorGpsSql)
    If rs.EOF = False Then
        ObtemNumeroCasa = rs(0).value               'obtem o número da casa
    Else
        ObtemNumeroCasa = "s/n"                     'não encontrou o registro com o número da casa. Não deveria nunca entrar aqui
    End If
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsConsumidorControler", "Public Function ObtemNumeroCasa", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

'Obtem o endereço completo do consumidor levantado em campo com o coletor do software comercial
'
'
Public Function ObtemEnderecoCompleto(objIDLigacao As String) As String
On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim consumidorGpsSql As String
    
    consumidorGpsSql = "SELECT ENDERECO, CONSUMIDOR, NRO_LIGACA from NXGS_V_LIG_COMERCIAL_GPS where " & ObtemNomeColunaObjId("NXGS_V_LIG_COMERCIAL_GPS") & " = '" & objIDLigacao & "'"
    Set rs = Conn.execute(consumidorGpsSql)
    If rs.EOF = False Then
        ObtemEnderecoCompleto = rs(0).value & " - " & rs(1).value & " - " & rs(2).value             'obtem o número da casa
    Else
        ObtemEnderecoCompleto = "Sem endereço"                     'não encontrou o registro com o número da casa. Não deveria nunca entrar aqui
    End If
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsConsumidorControler", "Public Function ObtemEnderecoCompleto", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function


'Obtem o nome da coluna do object_id de um layer
'
'
Private Function ObtemNomeColunaObjId(nomeLayer As String) As String
On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim consumidorGpsSql As String
    
    consumidorGpsSql = "SELECT unique_id from te_layer_table where attr_table = '" & nomeLayer & "'"
    Set rs = Conn.execute(consumidorGpsSql)
    If rs.EOF = False Then
        ObtemNomeColunaObjId = rs(0).value          'obtem o nome da coluna que contem o object_id
    Else
        ObtemNomeColunaObjId = "Falhou"             'não encontrou o registro com o nome da coluna, o banco de dados está inconsistente
    End If
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsConsumidorControler", "Public Function ObtemNomeColunaObjId", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

'Insere um texto com o número da casa
'Esta função é utilizada para a atualização inicial de todos os números das casas
'
Private Sub InsereTextoQualquerNoPontoGps(objIdPontoGPS As String, textoParaInserir As String)
         Dim coordXcentroGeometria As Double, coordYcentroGeometria As Double
         Dim gpsConsumidorSql As String
         Dim numeroCasa As String
                                                                        
         gpsConsumidorSql = "Delete From Texts" & GetLayerID("NXGS_V_LIG_COMERCIAL_GPS") & "  Where Object_id = '" & objIdPontoGPS & "'"
         Conn.execute gpsConsumidorSql
         tdbcon.setCurrentLayer "NXGS_V_LIG_COMERCIAL_GPS"
         tdbcon.getCenterGeometry CLng(objIdPontoGPS), objIdPontoGPS, tpPOINTS, coordXcentroGeometria, coordYcentroGeometria    'localiza as coordenadas X e Y do ponto GPS selecionado
         tdbcon.InsertTextInDatabase objIdPontoGPS, textoParaInserir, 1, coordXcentroGeometria, coordYcentroGeometria, 0, TeHorzLeft, TeHorzLeft, True
         tcs.plotView
End Sub


'Atualiza todos os números das casas dos pontos levantados em campo com GPS
'
Public Sub AtualizaNumeroGpsCasas()
    Dim rs As ADODB.Recordset
    Dim consumidorGpsSql As String
    Dim consumidorEndereco As String
    Dim consumidorNumero As String
    Dim consumidorComplemento As String
    Dim consumidorObjId As String
    
    consumidorGpsSql = "SELECT " & ObtemNomeColunaObjId("NXGS_V_LIG_COMERCIAL_GPS") & ", ENDERECO, NRO_CASA, COMP_LOG from NXGS_V_LIG_COMERCIAL_GPS"
    Set rs = Conn.execute(consumidorGpsSql)
    Do While rs.EOF = False
        consumidorObjId = rs(0).value
        consumidorEndereco = IIf(IsNull(rs!ENDERECO), "", rs!ENDERECO)
        consumidorNumero = IIf(IsNull(rs!NRO_CASA), "", rs!NRO_CASA)
        consumidorComplemento = IIf(IsNull(rs!COMP_LOG), "", rs!COMP_LOG)
        Call InsereTextoQualquerNoPontoGps(consumidorObjId, consumidorNumero)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsConsumidorControler", "Public Function ObtemNomeColunaObjId", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub
