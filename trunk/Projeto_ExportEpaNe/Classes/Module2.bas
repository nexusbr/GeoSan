Attribute VB_Name = "Module2"
'Option Explicit
'Private MyTime As Date
'Private QdeLida As Integer
'
'
'Private tb As New TeDatabase
'Private conn As New ADODB.Connection
'Private rsTrechosExportados As New ADODB.Recordset
'Private rsCoordinates As New ADODB.Recordset
'Private rsPipes As New ADODB.Recordset
'Private rsJunctions As New ADODB.Recordset
'Private rsPumps As New ADODB.Recordset
'Private rsValves As New ADODB.Recordset
'Private rsReservoirs As New ADODB.Recordset
'Private rsTanks As New ADODB.Recordset
'Private rsVertices As New ADODB.Recordset
'Private rsNosTmp As New ADODB.Recordset
'Private Layer_id As Integer
'Private rsNos As Recordset
'Sub ExportaEPA(rsTrechos As ADODB.Recordset, mconn As ADODB.Connection)
'   Set conn = mconn
'   Dim a As Integer, rsOutroTrecho As ADODB.Recordset, NO As String, NoI As String, NoF As String
'   Dim Trechoi As String, Trechof As String, TrechoiCompr As Double, TrechofCompr As Double
'
'
'   tb.Provider = Provider
'   tb.Connection = conn
'   tb.setCurrentLayer "waterlines"
'   AbrirTrechosExportados
'   AbrirEstruturaExporteRede
'   Screen.MousePointer = vbHourglass
'   Layer_id = GetLayerID("watercomponents")
'
'   Set rsNos = New ADODB.Recordset
'
'   rsNos.Open "Select OBJECT_ID_," & _
'               "x , y, id_TYPE, GROUNDHEIGHT, DEMAND, Pattern " & _
'               "from watercomponents inner join points" & Layer_id & " on object_id_=object_id", conn
'
'
'   While Not rsNos.EOF
'      With rsNosTmp
'         .AddNew
'         .Fields("ID").Value = rsNos.Fields("OBJECT_ID_").Value
'         .Fields("X").Value = rsNos.Fields("X").Value
'         .Fields("Y").Value = rsNos.Fields("Y").Value
'         .Fields("Tipo").Value = rsNos.Fields("ID_TYPE").Value
'         .Fields("Cota").Value = rsNos.Fields("GROUNDHEIGHT").Value
'         .Fields("Demanda").Value = rsNos.Fields("demand").Value
'         .Fields("Padrao").Value = IIf(IsNull(rsNos.Fields("PATTERN").Value), 0, rsNos.Fields("PATTERN").Value)
'      End With
'      rsNos.MoveNext
'   Wend
'   rsNos.Close
'   Set rsNos = Nothing
'
'
'
'
'   While Not rsTrechos.EOF
'      For a = 1 To 2
'         If a = 1 Then
'            NO = rsTrechos("Initialcomponent")
'         Else
'            NO = rsTrechos("finalcomponent")
'         End If
'         'pega a tubulação
'         If TubulacaoNaoCadastrada(rsTrechos.Fields("object_id_").Value) = True Then
'            rsNosTmp.Filter = "id='" & NO & "'"
'            If rsNosTmp.EOF Then
'               Exit For
'            End If
'            Select Case rsNosTmp.Fields("tipo").Value
'               Case No_Valvulas
'                  'Criar Funcão Retornando comprimento do Trecho1 e Dois
'                  Select Case Sub_Tipo_Valvula(NO)
'                     Case val_Desconhecida, Val_Esferica, Val_Gaveta, Val_Retencao 'Valvulas de secionamento ou Retenção or desconhecida
'                        RetornaNosAdjacentes NO, NoI, NoF, Trechoi, Trechoi, TrechoiCompr, TrechofCompr
'                        AddJunction NoI
'                        AddJunction NoF
'
'                        AddPipe rsTrechos.Fields("Object_id_").Value, _
'                              NoI, NoF, TrechoiCompr + TrechofCompr, _
'                              rsTrechos("internaldiameter").Value, _
'                              rsTrechos("roughness").Value, 0, StatusValve(NO)
'                        Exit For
'                     Case Else 'Valvulas de Controle
'                        RetornaNosAdjacentes NO, NoI, NoF, Trechoi, Trechof, TrechoiCompr, TrechofCompr
'                        AddVitualNosVB NO, NoI, NoF
'                        AddValves NO, NoI & "A", NoF & "A"
'                        If Not E_Nos_Especial(NoI) Then
'                           AddPipe Trechoi, NoI, NoI & "A", TrechoiCompr, rsTrechos("internaldiameter").Value, rsTrechos("roughness").Value, 0, "Open"
'                        End If
'                        If Not NoUltimo(NoF) And Not E_Nos_Especial(NoF) Then
'                           AddPipe Trechof, NoF, NoF & "A", TrechofCompr, rsTrechos("internaldiameter").Value, rsTrechos("roughness").Value, 0, "Open"
'                           AddJunction NoF
'                        End If
'
'                        Exit For
'                  End Select
'               Case No_Bombas
'
'                  RetornaNosAdjacentes NO, NoI, NoF, Trechoi, Trechof, TrechoiCompr, TrechofCompr
'                  If AddVitualNosVB(NO, NoI, NoF) Then
'                     AddPumps NO, NoI & "A", NoF & "A"
'                     AddPipe Trechoi, NoI, NoI & "A", TrechoiCompr, rsTrechos("internaldiameter").Value, rsTrechos("roughness").Value, 0, "Open"
'                     If Not NoUltimo(NoF) And Not E_Nos_Especial(NoF) Then
'                        AddPipe Trechof, NoF, NoF & "A", TrechofCompr, rsTrechos("internaldiameter").Value, rsTrechos("roughness").Value, 0, "Open"
'                     End If
'                     Exit For
'                  Else
'                     a = 1
'                  End If
'               Case No_Reservatorios
'                  AddReservoirs NO
'                  If a = 2 Then
'                     RetornaNosAdjacentes NO, NoI, NoF, Trechoi, Trechoi, TrechoiCompr, TrechofCompr
'                     If VerificaBombaOuValvula(NoI) Then
'                        NoI = NO & "A"
'                     End If
'                     If VerificaBombaOuValvula(NoF) Then
'                        NoF = NO & "A"
'                     End If
'                     AddPipe rsTrechos.Fields("Object_id_").Value, _
'                              NoI, _
'                              NoF, _
'                              IIf(rsTrechos("Length").Value = 0, rsTrechos("LengthCalculated").Value, rsTrechos("Length").Value), _
'                              rsTrechos("internaldiameter").Value, _
'                              rsTrechos("roughness").Value, 0, "Open"
'                  End If
'
'                  rsNos.Close
'                  Exit For
'               Case No_Tanques
'                  AddTanks NO
'
'                  If a = 2 Then
'                     RetornaNosAdjacentes NO, NoI, NoF, Trechoi, Trechoi, TrechoiCompr, TrechofCompr
'                     If VerificaBombaOuValvula(NoI) Then
'                        NoI = NO & "A"
'                     End If
'                     If VerificaBombaOuValvula(NoF) Then
'                        NoF = NO & "A"
'                     Else
'                        NoF = NO
'                     End If
'                     AddPipe rsTrechos.Fields("Object_id_").Value, _
'                              NoI, _
'                              NoF, _
'                              IIf(rsTrechos("Length").Value = 0, rsTrechos("LengthCalculated").Value, rsTrechos("Length").Value), _
'                              rsTrechos("internaldiameter").Value, _
'                              rsTrechos("roughness").Value, 0, "Open"
'                  End If
'                  Exit For
'               Case Else 'Outros Nos
'                  AddJunction NO
'                  If a = 2 Then
'                     'RetornaNosAdjacentes no, noi, nof
'                     NoI = rsTrechos("initialcomponent").Value
'                     NoF = rsTrechos("finalcomponent").Value
'                     If VerificaBombaOuValvula(NoI) Then
'                        NoI = NO & "A"
'                     End If
'                     If VerificaBombaOuValvula(NoF) Then
'                        NoF = NO & "A"
'                     End If
'                     AddPipe rsTrechos.Fields("Object_id_").Value, _
'                              NoI, _
'                              NoF, _
'                              IIf(rsTrechos("Length").Value = 0, rsTrechos("LengthCalculated").Value, rsTrechos("Length").Value), _
'                              rsTrechos("internaldiameter").Value, _
'                              rsTrechos("roughness").Value, 0, "Open"
'                  End If
'            End Select
'         Else
'
'         End If
'      Next
'      a = 1
'      Set rsNos = Nothing
'      rsTrechos.MoveNext
'      frmOdometro.Caption = frmOdometro.ProgressBar1.Value & " até " & frmOdometro.ProgressBar1.Max
'      frmOdometro.ProgressBar1.Value = frmOdometro.ProgressBar1.Value + 1
'      DoEvents
'   Wend
'   Set rsNosTmp = Nothing
'   rsTrechos.Close
'   Set rsTrechos = Nothing
'   rsTrechosExportados.Close
'   Set rsTrechosExportados = Nothing
'   Set rsNos = Nothing
'   GeraArquivo_de_Saida
'   Screen.MousePointer = vbNormal
'   MsgBox "ok"
'End Sub
'
''Function TubulacaoNaoCadastrada(Object_id_ As String) As Boolean
''   rsTrechosExportados.Filter = "object_id_='" & Object_id_ & "'"
''   If rsTrechosExportados.EOF Then
''      TubulacaoNaoCadastrada = True
''   Else
''      TubulacaoNaoCadastrada = False
''   End If
''End Function
'
''Function NoNaoExisteCoordinates(Object_id_ As String) As Boolean
''   rsCoordinates.Filter = "id='" & Object_id_ & "'"
''   If rsCoordinates.EOF Then
''      NoNaoExisteCoordinates = True
''   Else
''      NoNaoExisteCoordinates = False
''   End If
''End Function
'
'Function NoUltimo(id As String) As Boolean
'   Dim Rs As ADODB.Recordset
'   Set Rs = conn.Execute("Select count(*) from waterlines " & _
'           "where initialcomponent = '" & id & "' or finalcomponent = '" & id & "'")
'   If Rs(0).Value = 1 Then
'      NoUltimo = True
'   End If
'   Rs.Close
'   Set Rs = Nothing
'End Function
'
'
'Function Sub_Tipo_Valvula(Object_id_ As String) As Integer
'   Dim Rs As ADODB.Recordset
'   Set Rs = conn.Execute("Select value_ from watercomponentsdata where object_id_='" & Object_id_ & "' and id_subtype=1")
'   If Rs.EOF Then
'      Sub_Tipo_Valvula = 0
'   Else
'      Sub_Tipo_Valvula = Rs.Fields(0).Value
'   End If
'   Rs.Close
'   Set Rs = Nothing
'End Function
'
'
'Sub AbrirTrechosExportados()
'   rsTrechosExportados.Fields.Append "object_id_", adVarChar, 255
'   rsTrechosExportados.Open
'End Sub
'
'Sub AbrirEstruturaExporteRede()
'   rsCoordinates.Fields.Append "id", adVarChar, 255
'   rsCoordinates.Fields.Append "x", adDouble
'   rsCoordinates.Fields.Append "y", adDouble
'   rsCoordinates.Open
'
'   rsPipes.Fields.Append "id", adVarChar, 255
'   rsPipes.Fields.Append "node1", adVarChar, 255
'   rsPipes.Fields.Append "node2", adVarChar, 255
'   rsPipes.Fields.Append "length", adDouble, 255
'   rsPipes.Fields.Append "diameter", adDouble, 255
'   rsPipes.Fields.Append "roughness", adDouble, 255
'   rsPipes.Fields.Append "minorloss", adDouble, 255
'   rsPipes.Fields.Append "status", adVarChar, 255
'   rsPipes.Open
'
'   rsJunctions.Fields.Append "id", adVarChar, 255
'   rsJunctions.Fields.Append "elev", adVarChar, 255
'   rsJunctions.Fields.Append "demand", adDouble, 255
'   rsJunctions.Fields.Append "pattern", adVarChar, 255
'   rsJunctions.Open
'
'   rsPumps.Fields.Append "id", adVarChar, 255
'   rsPumps.Fields.Append "node1", adVarChar, 255
'   rsPumps.Fields.Append "node2", adVarChar, 255
'   rsPumps.Fields.Append "parameters", adVarChar, 255
'   rsPumps.Open
'
'   rsValves.Fields.Append "id", adVarChar, 255
'   rsValves.Fields.Append "node1", adVarChar, 255
'   rsValves.Fields.Append "node2", adVarChar, 255
'   rsValves.Fields.Append "diameter", adDouble
'   rsValves.Fields.Append "type", adVarChar, 255
'   rsValves.Fields.Append "setting", adVarChar, 255
'   rsValves.Fields.Append "minorloss", adVarChar, 255
'   rsValves.Open
'
'   rsReservoirs.Fields.Append "ID", adVarChar, 255
'   rsReservoirs.Fields.Append "Head", adVarChar, 255
'   rsReservoirs.Fields.Append "Pattern", adVarChar, 255
'   rsReservoirs.Open
'
'   rsTanks.Fields.Append "ID", adVarChar, 255
'   rsTanks.Fields.Append "Elevation", adVarChar, 255
'   rsTanks.Fields.Append "InitLevel", adDouble
'   rsTanks.Fields.Append "MinLevel", adDouble
'   rsTanks.Fields.Append "MaxLevel", adDouble
'   rsTanks.Fields.Append "Diameter", adDouble
'   rsTanks.Fields.Append "MinVol", adDouble
'   rsTanks.Fields.Append "VolCurve", adDouble
'   rsTanks.Open
'
'   rsVertices.Fields.Append "ID", adVarChar, 255
'   rsVertices.Fields.Append "X-Coord", adDouble
'   rsVertices.Fields.Append "Y-Coord", adDouble
'   rsVertices.Open
'
'   rsNosTmp.Fields.Append "ID", adVarChar, 255
'   rsNosTmp.Fields.Append "X", adDouble
'   rsNosTmp.Fields.Append "Y", adDouble
'   rsNosTmp.Fields.Append "Tipo", adInteger
'   rsNosTmp.Fields.Append "Padrao", adInteger
'   rsNosTmp.Fields.Append "Curva", adInteger
'   rsNosTmp.Fields.Append "Diametro", adDouble
'   rsNosTmp.Fields.Append "Cota", adDouble
'   rsNosTmp.Fields.Append "NivelMin", adDouble
'   rsNosTmp.Fields.Append "NivelMax", adDouble
'   rsNosTmp.Fields.Append "VolumeMin", adDouble
'   rsNosTmp.Fields.Append "CurvaVol", adDouble
'   rsNosTmp.Fields.Append "Parametros", adDouble
'   rsNosTmp.Fields.Append "setting", adDouble
'   rsNosTmp.Fields.Append "type", adDouble
'   rsNosTmp.Fields.Append "demanda", adDouble
'
'   rsNosTmp.Open
'
'
'End Sub
'
'Sub AddPumps(id As String, Node1 As String, Node2 As String)
'   Dim Rs As ADODB.Recordset
'   Dim CURVE As String
'   Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
'                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
'                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
'                         "where object_id_ = '" & id & "'")
'   While Not Rs.EOF
'      Select Case Rs.Fields("EPAREF").Value
'         Case "CURVE"
'            CURVE = Rs.Fields("VALUE_").Value
'            'IMPLEMENTAR?
'      End Select
'      Rs.MoveNext
'   Wend
'   Rs.Close
'   Set Rs = Nothing
'   rsPumps.AddNew
'   rsPumps.Fields("id").Value = id
'   rsPumps.Fields("Node1").Value = Node1
'   rsPumps.Fields("Node2").Value = Node2
'   rsPumps.Fields("Parameters").Value = " HEAD " & CURVE
'End Sub
'
'Sub AddPipe(id As String, Node1 As String, Node2 As String, length As Double, Diameter As String, roughness As Double, MinorLoss As Double, status As String)
'   If Node1 = "" Or Node2 = "" Then Exit Sub
'   rsPipes.AddNew
'   rsPipes.Fields("id").Value = id
'   rsPipes.Fields("node1").Value = Node1
'   rsPipes.Fields("node2").Value = Node2
'   rsPipes.Fields("length").Value = Replace(length, ",", ".")
'   rsPipes.Fields("diameter").Value = Diameter
'   rsPipes.Fields("roughness").Value = Replace(roughness, ",", ".")
'   rsPipes.Fields("minorloss").Value = MinorLoss
'   rsPipes.Fields("status").Value = status
'   AddVertices id
'   rsTrechosExportados.AddNew
'   rsTrechosExportados.Fields("object_id_").Value = id
'End Sub
'
'
'Sub AddVertices(id As String)
'   Dim a As Integer, Qde As Integer, x As Double, y As Double
'   Qde = tb.getQuantityPointsLine(0, id)
'   If Qde > 2 Then
'      For a = 1 To Qde - 2
'         With rsVertices
'            If tb.getPointOfLine(0, id, a, x, y) = 1 Then
'               .AddNew
'               .Fields("id").Value = id
'               .Fields("x-coord").Value = Replace(x, ",", ".")
'               .Fields("y-coord").Value = Replace(y, ",", ".")
'            End If
'         End With
'      Next
'   End If
'End Sub
'
'
'Sub AddJunction(id As String)
'   If id = "" Then Exit Sub
'   If NoNaoExisteCoordinates(id) Then
'      rsJunctions.AddNew
'      rsJunctions.Fields("id").Value = id
'      rsJunctions.Fields("elev").Value = IIf(IsNull(rsNosTmp("cota").Value), 0, rsNosTmp("cota").Value)
'      rsJunctions.Fields("demand").Value = ((((rsNosTmp("demanda").Value * 1000) / 24) / 30) / 3600)
'      rsJunctions.Fields("pattern").Value = IIf(IsNull(rsNosTmp("padrao").Value), "", rsNosTmp("padrao").Value)
'      AddCoordinate id
'   End If
'End Sub
'
'Sub AddCoordinate(id As String)
'   Dim x As Double, y As Double
'   rsCoordinates.AddNew
'   rsCoordinates.Fields("id").Value = id
'   rsCoordinates.Fields("x").Value = rsNosTmp("x").Value
'   rsCoordinates.Fields("y").Value = rsNosTmp("y").Value
'End Sub
'
'Sub AddReservoirs(id As String)
'   If NoNaoExisteCoordinates(id) Then
'      rsReservoirs.AddNew
'      rsReservoirs.Fields("ID").Value = id
'      rsReservoirs.Fields("Head").Value = ""
'      rsReservoirs.Fields("Pattern").Value = ""
'      AddCoordinate id
'   End If
'End Sub
'
'Sub AddTanks(id As String)
'   If NoNaoExisteCoordinates(id) Then
'      rsTanks.AddNew
'      rsTanks.Fields("ID").Value = id
'      rsTanks.Fields("Elevation").Value = 0
'      rsTanks.Fields("InitLevel").Value = 0
'      rsTanks.Fields("MinLevel").Value = 0
'      rsTanks.Fields("MaxLevel").Value = 0
'      rsTanks.Fields("Diameter").Value = 0
'      rsTanks.Fields("MinVol").Value = 0
'      rsTanks.Fields("VolCurve").Value = 0
'      AddCoordinate id
'   End If
'End Sub
'
'Sub AddValves(id As String, Node1 As String, Node2 As String)
'   Dim Rs As ADODB.Recordset
'   Dim PumpDiameter As Double, PumpType As String, PumpSetting As String, PumpMinorLoss As String
'   Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
'                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
'                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
'                         "where object_id_ = '" & id & "'")
'   While Not Rs.EOF
'      Select Case Rs.Fields("EPAREF").Value
'         Case "TYPE"
'            PumpType = Rs.Fields("DESCRIPTION_").Value
'         Case "SETTING"
'            PumpSetting = Rs.Fields("VALUE_").Value
'         Case "DIAMETER"
'            PumpDiameter = Rs.Fields("VALUE_").Value
'         Case "NINORLOSS"
'            'IMPLEMENTAR?
'      End Select
'      Rs.MoveNext
'   Wend
'   Rs.Close
'   Set Rs = Nothing
'
'   rsValves.AddNew
'   rsValves.Fields("ID").Value = id
'   rsValves.Fields("Node1").Value = Node1
'   rsValves.Fields("Node2").Value = Node2
'   rsValves.Fields("Diameter").Value = PumpDiameter
'   rsValves.Fields("Type").Value = PumpType
'   rsValves.Fields("Setting").Value = PumpSetting
'   rsValves.Fields("MinorLoss").Value = PumpMinorLoss
'
'End Sub
'
'Function StatusValve(NO As String) As String
'   Dim Rs As ADODB.Recordset
'   Set Rs = conn.Execute("Select w.id_subtype, w.value_ as valor, s.description_ as descricao from watercomponentsdata w " & _
'                         "left join WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as int)=s.value_ " & _
'                         "where object_id_ = " & NO)
'   While Not Rs.EOF
'      Select Case Rs.Fields("id_subtype").Value
'         Case 1 'Tipo
'         Case 2 'Estado
'            StatusValve = Rs.Fields("Descricao").Value
'         Case 4 'Diametro
'         Case 5 'Coeficiente
'      End Select
'      Rs.MoveNext
'   Wend
'   Rs.Close
'   Set Rs = Nothing
'End Function
'
'
'Sub GeraArquivo_de_Saida()
'
'   Dim a As Integer, str As String
'   Open FrmEPANET.txtArquivo.Text For Output As #1
'
'      With rsJunctions
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, "[JUNCTIONS]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'      With rsReservoirs
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[RESERVOIRS]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'      With rsTanks
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[TANKS]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'
'      With rsPumps
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[PUMPS]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'      With rsValves
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[VALVES]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'      With rsPipes
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[PIPES]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'      Dim MyArray() As String
'      Dim rsPatterns As ADODB.Recordset
'      Set rsPatterns = conn.Execute("Select * from x_patterns")
'
'      With rsPatterns
'         Print #1, "[PATTERNS]"
'         Print #1, ";ID" & Chr(vbKeyTab) & Chr(vbKeyTab) & "Multipliers"
'         Print #1, ";" & rsPatterns("descricao").Value
'         While Not .EOF
'            MyArray = Split(rsPatterns("Padrao").Value, ";", 25)
'            Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(0), ",", ".") & _
'                      Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(1), ",", ".") & _
'                      Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(2), ",", ".") & _
'                      Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(3), ",", ".") & _
'                      Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(4), ",", ".") & _
'                      Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(5), ",", ".")
'            If MyArray(6) <> "" Then
'               Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(6), ",", ".") & _
'                         Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(7), ",", ".") & _
'                         Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(8), ",", ".") & _
'                         Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(9), ",", ".") & _
'                         Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(10), ",", ".") & _
'                         Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(11), ",", ".")
'               If MyArray(12) <> "" Then
'                  Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(12), ",", ".") & _
'                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(13), ",", ".") & _
'                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(14), ",", ".") & _
'                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(15), ",", ".") & _
'                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(16), ",", ".") & _
'                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(17), ",", ".")
'                  If MyArray(18) <> "" Then
'                     Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(18), ",", ".") & _
'                               Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(19), ",", ".") & _
'                               Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(20), ",", ".") & _
'                               Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(21), ",", ".") & _
'                               Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(22), ",", ".") & _
'                               Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(23), ",", ".")
'                  End If
'               End If
'            End If
'            rsPatterns.MoveNext
'         Wend
'      End With
'      rsPatterns.Close
'      Set rsPatterns = Nothing
'      Dim b As Integer
'      Dim MyArray_x() As String
'      Dim MyArray_y() As String
'      Dim rsCurves As ADODB.Recordset
'      Set rsCurves = conn.Execute("Select * from x_Curves order by tipo")
'
'      With rsCurves
'
'         Print #1, "[CURVES]"
'         Print #1, ";ID" & Chr(vbKeyTab) & Chr(vbKeyTab) & "X-Value" & Chr(vbKeyTab) & Chr(vbKeyTab) & "Y-Value"
'         For b = 1 To 4
'            If b = 1 Then
'               rsCurves.Filter = "Tipo = 'Bomba'"
'               If Not rsCurves.EOF Then Print #1, ";PUMPS:" & rsCurves.Fields("descricao").Value
'            ElseIf b = 2 Then
'               rsCurves.Filter = "Tipo = 'Rendimento'"
'               If Not rsCurves.EOF Then Print #1, ";EFFICIENCY:" & rsCurves.Fields("descicao").Value
'            ElseIf b = 3 Then
'               rsCurves.Filter = "Tipo = 'Volume'"
'               If Not rsCurves.EOF Then Print #1, ";VOLUME:" & rsCurves.Fields("descicao").Value
'            ElseIf b = 4 Then
'               rsCurves.Filter = "Tipo = 'Perda de Carga'"
'               If Not rsCurves.EOF Then Print #1, ";HEADLOSS:" & rsCurves.Fields("descicao").Value
'            End If
'            While Not .EOF
'               MyArray_x = Split(rsCurves("Coordenada_x").Value, ";", 50)
'               MyArray_y = Split(rsCurves("Coordenada_y").Value, ";", 50)
'               For a = 0 To 49
'                  If MyArray_x(a) = "" Then
'                     Exit For
'                  Else
'                     Print #1, .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray_x(a), ",", ".") & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray_y(a), ",", ".")
'                  End If
'               Next
'               rsCurves.MoveNext
'            Wend
'          Next
'      End With
'      rsCurves.Close
'      Set rsCurves = Nothing
'
'
'      With rsCoordinates
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[COORDINATES]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'      With rsVertices
'         .Filter = ""
'         If .RecordCount > 0 Then
'            .MoveFirst
'            For a = 0 To .Fields.Count - 1
'               str = str & .Fields(a).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'            Next
'            Print #1, ""
'            Print #1, "[VERTICES]"
'            Print #1, ";" & str
'            str = ""
'            While Not .EOF
'               For a = 0 To .Fields.Count - 1
'                  str = str & IIf(IsNumeric(.Fields(a).Value), _
'                           Replace(.Fields(a).Value, ",", "."), _
'                           .Fields(a).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'               Next
'               Print #1, str & ";"
'               str = ""
'               .MoveNext
'            Wend
'         End If
'      End With
'
'
'
'
'
''      GravaArquivo "TITLE"
''      GravaArquivo "JUNCTIONS"
''      GravaArquivo "RESERVOIRS"
''      GravaArquivo "TANKS"
''      GravaArquivo "PUMPS"
''      GravaArquivo "VALVES"
''      GravaArquivo "PIPES"
''      GravaArquivo "TAGS"
''      GravaArquivo "DEMANDS"
''      GravaArquivo "PATTERNS"
''      GravaArquivo "CURVES"
''      GravaArquivo "CONTROLS"
''      GravaArquivo "RULES"
''      GravaArquivo "ENERGY"
''      GravaArquivo "EMITTERS"
''      GravaArquivo "SOURCES"
''      GravaArquivo "REACTIONS"
''      GravaArquivo "MIXING"
''      GravaArquivo "REACTIONS"
''      GravaArquivo "TIMES"
''      GravaArquivo "REPORT"
''      GravaArquivo "OPTIONS"
''      GravaArquivo "COORDINATES"
''      GravaArquivo "VERTICES"
''      GravaArquivo "BACKDROP"
''      GravaArquivo "END"
'   Close #1
'End Sub
'
'Sub RetornaNosSeValvula(NO As String)
'   Dim RsRetornaNos As ADODB.Recordset
'   Set RsRetornaNos = conn.Execute("Select * from watercomponents inner join points39 on object_id=object_id_ where object_id_='" & NO & "'")
'   Select Case RsRetornaNos("id_type").Value
'      Case 1
'         NO = NO & "A"
'   End Select
'   RsRetornaNos.Close
'   Set RsRetornaNos = Nothing
'End Sub
'
'
'Sub RetornaNosAdjacentes(NO As String, NoI As String, NoF As String, Trechoi As String, Trechof As String, TrechoiCompr As Double, TrechofCompr As Double)
'   Dim rsOutroTrecho As ADODB.Recordset
'   Set rsOutroTrecho = conn.Execute("Select * from waterlines " & _
'           "where initialcomponent = '" & NO & "' or finalcomponent = '" & NO & "'")
'   If rsOutroTrecho.Fields("initialcomponent") <> NO Then
'      NoI = rsOutroTrecho.Fields("initialcomponent")
'   ElseIf rsOutroTrecho.Fields("finalcomponent") <> NO Then
'      NoI = rsOutroTrecho.Fields("finalcomponent")
'   End If
'   Trechoi = rsOutroTrecho("object_id_").Value
'   TrechoiCompr = IIf(rsOutroTrecho("Length").Value = 0, rsOutroTrecho("LengthCalculated").Value, rsOutroTrecho("Length").Value)
'   rsOutroTrecho.MoveNext
'   If rsOutroTrecho.EOF Then
'      NoF = NO 'SE NO ADJACENTES NÃO FOR VALCULA OU BOMBA
'      Exit Sub
'   End If
'   Trechof = rsOutroTrecho("object_id_").Value
'   TrechofCompr = IIf(rsOutroTrecho("Length").Value = 0, rsOutroTrecho("LengthCalculated").Value, rsOutroTrecho("Length").Value)
'   If rsOutroTrecho.Fields("initialcomponent") <> NO Then
'      NoF = rsOutroTrecho.Fields("initialcomponent")
'   ElseIf rsOutroTrecho.Fields("finalcomponent") <> NO Then
'      NoF = rsOutroTrecho.Fields("finalcomponent")
'   End If
'   rsOutroTrecho.Close
'   Set rsOutroTrecho = Nothing
'   Set rsOutroTrecho = conn.Execute("Select count(*) from waterlines " & _
'           "where initialcomponent = '" & NoI & "' or finalcomponent = '" & NoI & "'")
'   If rsOutroTrecho(0).Value > 1 Then
'      rsTrechosExportados.AddNew
'      rsTrechosExportados(0).Value = Trechoi
'   End If
'   Set rsOutroTrecho = conn.Execute("Select count(*) from waterlines " & _
'           "where initialcomponent = '" & NoF & "' or finalcomponent = '" & NoF & "'")
'   If rsOutroTrecho(0).Value > 1 Then
'      rsTrechosExportados.AddNew
'      rsTrechosExportados(0).Value = Trechof
'   End If
'End Sub
'
'
'Function AddVitualNosVB(NO As String, ByRef NoI As String, ByRef NoF As String) As Boolean
'   Dim rsCoordenada As ADODB.Recordset, x As Double, y As Double, compr As Double
'
'
'   rsPumps.Filter = "id='" & NO & "'"
'   If rsPumps.EOF Then
'      Set rsCoordenada = conn.Execute("Select object_id_ from waterlines " & _
'              "where initialcomponent = '" & NO & "' or finalcomponent = '" & NO & "'")
'
'      If Not NoNaoExisteCoordinates(NO & "A") And E_Nos_Especial(NoI) Then
'         NoI = NO
'      Else
'         rsJunctions.AddNew
'         rsJunctions.Fields("id").Value = NoI & "A"
'         rsJunctions.Fields("elev").Value = 0
'         rsJunctions.Fields("demand").Value = 0
'         rsJunctions.Fields("pattern").Value = ""
'
'         tb.getCenterGeometry 0, rsCoordenada(0).Value, 2, x, y
'         rsCoordinates.AddNew
'         rsCoordinates.Fields("id").Value = NoI & "A"
'         rsCoordinates.Fields("x").Value = x
'         rsCoordinates.Fields("y").Value = y
'
'
'      End If
'      rsCoordenada.MoveNext
'
'
'      If Not NoNaoExisteCoordinates(NO & "A") And E_Nos_Especial(NoF) Then
'         NoF = NO
'      Else
'         rsJunctions.AddNew
'         rsJunctions.Fields("id").Value = NoF & "A"
'         rsJunctions.Fields("elev").Value = 0
'         rsJunctions.Fields("demand").Value = 0
'         rsJunctions.Fields("pattern").Value = ""
'
'         tb.getCenterGeometry 0, rsCoordenada(0).Value, 2, x, y
'         rsCoordinates.AddNew
'         rsCoordinates.Fields("id").Value = NoF & "A"
'         rsCoordinates.Fields("x").Value = x
'         rsCoordinates.Fields("y").Value = y
'         rsCoordenada.Close
'      End If
'
'      AddVitualNosVB = True
'   End If
'   Set rsCoordenada = Nothing
'End Function
'
'
'Sub AddVitualNoEntreValvulas(Trecho As String, v1 As String, v2 As String)
'   Dim rsCoordenada As ADODB.Recordset, x As Double, y As Double, NoI As String, NoF
'   tb.Provider = 1
'   tb.Connection = conn
'   tb.setCurrentLayer "waterlines"
'   tb.getCenterGeometry 0, Trecho, 2, x, y
'   rsJunctions.AddNew
'   rsJunctions.Fields("id").Value = v1 & "A" & v1
'   rsJunctions.Fields("elev").Value = 0
'   rsJunctions.Fields("demand").Value = 0
'   rsJunctions.Fields("pattern").Value = ""
'   rsCoordinates.AddNew
'   rsCoordinates.Fields("id").Value = v1 & "A" & v1
'   rsCoordinates.Fields("x").Value = x
'   rsCoordinates.Fields("y").Value = y
'End Sub
'
'Function VerificaBombaOuValvula(NO) As Boolean
'   Dim Rs As ADODB.Recordset
'   Set Rs = conn.Execute("select id_type from watercomponents where object_id_='" & NO & "'")
'   If Not Rs.EOF Then
'      Select Case Rs("id_type").Value
'         Case 1, 20
'            VerificaBombaOuValvula = True
'      End Select
'   End If
'   Rs.Close
'   Set Rs = Nothing
'End Function
'
'Public Function GetLayerID(LayerName_ As String) As Integer
'   Dim Rs As ADODB.Recordset
'   Set Rs = conn.Execute("Select Layer_id from Te_Layer where name='" & LayerName_ & "'")
'   GetLayerID = Rs(0).Value
'   Rs.Close
'   Set Rs = Nothing
'End Function
'
'Function E_Nos_Especial(NO As String) As Boolean
'   Dim RsRetornaNos As ADODB.Recordset
'   Set RsRetornaNos = conn.Execute("Select * from watercomponents inner join points" & Layer_id & " on object_id=object_id_ where object_id_='" & NO & "'")
'   Select Case RsRetornaNos("id_type").Value
'      Case No_Bombas, No_Valvulas
'         E_Nos_Especial = True
'      Case Else
'         E_Nos_Especial = False
'   End Select
'   RsRetornaNos.Close
'   Set RsRetornaNos = Nothing
'End Function
'
'
'
'
'
'
