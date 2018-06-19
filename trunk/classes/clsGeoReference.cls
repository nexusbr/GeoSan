VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeoReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum LayerTypeRefence
   Poligonos = 0
   Trecho_Rede_Agua = 1
   Componente_Rede_Agua = 2
   Trecho_Rede_esgoto = 3
   Componente_Rede_Esgoto = 4
   Trecho_Rede_Drenagem = 5
   Componente_Rede_Drenagem = 6
   RAMAIS_AGUA = 7
   RAMAIS_ESGOTO = 8
   DOCUMENTOS = 9
   AMARRACAO = 10
   LOTES = 11
   OUTROS = 12
   CONSUMIDOR_GPS = 13
End Enum

Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim ii As String
Dim j As String
Dim k As String
Dim l As String
Dim vm As String



'Public stopProcess
Public object_ids As String
Private p As Boolean, i As Integer, MyLines() As String
Private mvartdb As TeDatabase 'local copy
Private mvartdbcon As TeDatabase 'local copy
Private mvartcs As TeCanvas
Public Property Set tdb(ByVal vData As TeDatabase)
    Set mvartdb = vData
End Property

Public Property Get tdb() As TeDatabase
    Set tdb = mvartdb
End Property

Public Property Set tdbcon(ByVal vData As TeDatabase)
    Set mvartdbcon = vData
End Property

Public Property Get tdbcon() As TeDatabase
    Set tdbcon = mvartdbcon
End Property
'Propriedade do Canvas
'
'
Public Property Set tcs(ByVal vData As TeCanvas)
    Set mvartcs = vData
End Property

Public Property Get tcs() As TeCanvas
    Set tcs = mvartcs
End Property

'###############################################################
' Autor     : Luis Claudio Rodrigues Domingues    Data: 20/12/05
' Nome      : GetReferenceLayer
' Descrição : Retorna o layer de referencia
'###############################################################

Public Function GetReferenceLayer(LayerName As String, Optional LayerLines As Boolean) As String
On Error GoTo Trata_Erro
   Dim rs As ADODB.Recordset, strsql As String
  
  If frmCanvas.TipoConexao <> 4 Then
    strsql = "SELECT * From X_LayersComponents where (layerline='" & LayerName & "' or layerline='" & UCase(LayerName) & "' or layercomponent='" & LayerName & "' or layercomponent='" & UCase(LayerName) & "') "
   If LayerLines Then
        strsql = strsql & " and layerline='" & LayerName & "'"
   End If
   
    Set rs = Conn.execute(strsql)
   
   If Not rs.EOF Then
      If LCase(rs.Fields("LAYERLINE").value) = LCase(LayerName) Then
         GetReferenceLayer = UCase(rs.Fields("LAYERCOMPONENT").value)
      Else
         GetReferenceLayer = UCase(rs.Fields("LAYERLINE").value)
      End If
   Else
      MsgBox "Não foi encontrado " & LayerName & " no banco de dados.", vbInformation
   End If
   
   Else
Dim va As String
Dim vm As String
Dim vi As String
Dim vo As String
Dim vu As String
Dim vc As String
va = "X_LAYERSCOMPONENTS"
vm = "LAYERLINE"
vi = "LAYERCOMPONENT"


'MsgBox Conn.ConnectionString

  strsql = "SELECT * From " + """" + va + """" + " where (" + """" + vm + """" + "='" & UCase(LayerName) & "' or " + """" + vm + """" + "='" & UCase(LayerName) & "' or " + """" + vi + """" + "='" & UCase(LayerName) & "' or " + """" + vi + """" + "='" & UCase(LayerName) & "')"
    
      If LayerLines Then

        strsql = strsql & " and " + """" + vm + """" + "='" & LayerName & "'"
   End If
       
       
  
       
       
       
       'rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   Set rs = Conn.execute(strsql)
   
   If Not rs.EOF Then
      If (rs.Fields("LAYERLINE").value) = LayerName Then
         GetReferenceLayer = (rs.Fields("LAYERCOMPONENT").value)
      Else
         GetReferenceLayer = (rs.Fields("LAYERLINE").value)
      End If
   Else
      MsgBox "Não foi encontrado " & LayerName & " no banco de dados.", vbInformation
   End If
   
   End If
  
  ' If LayerLines Then
    '    strsql = strsql & " and layerline='" & LayerName & "'"
  ' End If
  
   rs.Close
   Set rs = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
      PrintErro "clsGeoReference", "Public Function GetReferenceLayer", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Function

'###############################################################
' Autor     : Luis Claudio Rodrigues Domingues    Data: 02/10/06
' Nome      : GetReferenceLayer
' Descrição : Retorna o layer de referencia
'###############################################################

Public Function GetTypeText(LayerName As String) As Long
On Error GoTo Trata_Erro
   Dim rs As ADODB.Recordset
   a = "X_LAYERSCOMPONENTS"
   b = "LAYERLINE"
     If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT * From X_LayersComponents where layerline='" & LayerName & "'")
   Else
    Set rs = Conn.execute("SELECT * From " + """" + a + """" + " where " + """" + b + """" + "='" & LayerName & "'")
   End If
   If Not rs.EOF Then
      GetTypeText = rs.Fields("TypeText").value
   End If
   rs.Close
   Set rs = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
      PrintErro "clsGeoReference", "Public Function GetTypeText", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Function

'###############################################################
' Autor     : Luis Claudio Rodrigues Domingues    Data: 02/10/06
' Nome      : GetReferenceLayer
' Descrição : Retorna o layer de referencia
'###############################################################

'Public Function GetTypeReference(layerName As String) As Long
'   Dim rs As ADODB.Recordset
'   Set rs = Conn.execute("SELECT * From X_LayersComponents where layerline='" & layerName & "'")
'   If Not rs.EOF Then
'      GetTypeReference = rs.Fields("TypeReference").value
'   End If
'   rs.Close
'   Set rs = Nothing
'End Function


Public Function InsertTextInPoint(geom_id As Long, object_id As String, Layer As String, Optional CotaI As Double, Optional CotaF As Double) As Boolean
   
   Dim X As Double, Y As Double, Distancia As Integer
   
   Dim TbRefTexto As String
   Dim SQL As String
   Dim rs As ADODB.Recordset
   
   'LOCALIZA A TABELA DE TEXTOS PARA PEGAR AS COTAS E DEPOIS EXCLUIR OS REGISTROS
   
   Set rs = New ADODB.Recordset
    If frmCanvas.TipoConexao <> 4 Then
   SQL = "SELECT TR.GEOM_TABLE FROM TE_REPRESENTATION TR WHERE GEOM_TYPE = 128 AND TR.LAYER_ID IN (SELECT LAYER_ID FROM TE_LAYER WHERE NAME = '" & Layer & "')"
   Else
   Dim va As String
   Dim ve As String
   Dim vi As String
   Dim vo As String
   Dim vu As String
   Dim vc As String
va = "te_representation"
vm = "geom_table"
vi = "geom_type"
vo = "layer_id"
vu = "te_layer"
vc = "name"
   
   SQL = "SELECT " + """" + va + """" + "." + """" + vm + """" + " FROM " + """" + va + """" + "  WHERE " + """" + vi + """" + " = '128' AND " + """" + va + """" + "." + """" + vo + """" + " IN (SELECT " + """" + vo + """" + " FROM " + """" + vu + """" + " WHERE " + """" + vc + """" + " = '" & Layer & "')"
   End If
   
      rs.Open SQL, Conn, adOpenDynamic, adLockOptimistic
   
   If rs.EOF = False Then
      
      TbRefTexto = rs!GEOM_TABLE
      rs.Close
      
      'EXCLUI O REGISTRO
a = LCase(TbRefTexto)

c = "object_id"

      If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "DELETE FROM " & TbRefTexto & " WHERE OBJECT_ID = '" & object_id & "'"

      Else
      Conn.execute "DELETE FROM " + """" + a + """" + " WHERE " + """" + c + """" + " = '" & object_id & "'"
      End If
      
      Set rs = New ADODB.Recordset
       If frmCanvas.TipoConexao <> 4 Then
      SQL = "SELECT OBJECT_ID_, GROUNDHEIGHT AS " + """" + "INI" + """" + ", GROUNDHEIGHTFINAL AS " + """" + "FIM" + """" + " FROM " & Layer & " WHERE OBJECT_ID_ = '" & object_id & "' ORDER BY COMPONENT_ID"
      Else
      a = "OBJECT_ID_"
b = "INITIALGROUNDHEIGHT"
c = "FINALGROUNDHEIGHT"
d = "COMPONENT_ID"
     SQL = "SELECT " + """" + a + """" + "," + """" + b + """" + " AS INI, " + """" + c + """" + " AS FIM FROM " + """" + Layer + """" + " WHERE " + """" + a + """" + " = '" & object_id & "' ORDER BY " + """" + d + """" + ""
    
      
      End If
      
        rs.Open SQL, Conn, adOpenDynamic, adLockOptimistic

      If rs.EOF = False Then
         tdbcon.setCurrentLayer Layer
         
         'LOCALIZA O CENTRO DA GEOMETRIA DO PONTO EM X E Y
         tdbcon.getCenterGeometry CLng(rs!Object_id_), rs!Object_id_, tpPOINTS, X, Y
                  
         tdbcon.InsertTextInDatabase rs!Object_id_, Round(rs!ini, 3), 1, X, Y + 2, 0, , , True
         tdbcon.InsertTextInDatabase rs!Object_id_, Round(rs!fim, 3), 1, X, Y - 3, 0, , , True
         
         'INSERE OS TEXTOS COM A DIFERENÇA ENTRE COTA TERRENO E COTA PROFUNDIDADE NUMA NOVA LOCALIZAÇÃO
         
         Dim dblDIF As Double
         Dim strDIF As String
         
         dblDIF = rs!ini - rs!fim
         strDIF = Round(dblDIF, 3)
         
         tdbcon.InsertTextInDatabase rs!Object_id_, strDIF, 1, X + 3, Y - 0.5, 0, , , True
         
      End If

   End If
  
End Function

'Public Function UpdateTextInPoint(LayerName As String, object_id As String) As Boolean
'
'   InsertTextInPoint
'
'   On Error GoTo Trata_Erro
'   '####### DECLARAÇÃO DE VARIAVEIS LOCAIS ##########
'   If LCase(LayerName) = "watercomponents" Then Exit Function
'   Dim RsOut As ADODB.Recordset, RsInt As ADODB.Recordset, a As Integer
'   Set RsInt = Conn.execute("SELECT * from texts" & GetLayerID(LayerName) & " Where object_id='" & object_id & "'")
'   Set RsOut = Conn.execute("SELECT * from " & LayerName & " Where object_id_='" & object_id & "'")
'
'   '####### ROTINA DE ATUALIZAÇÃO ##############
'   For a = 1 To 2
'      Conn.execute "UPDATE TEXTS" & GetLayerID(LayerName) & _
'         " SET TEXT_VALUE = " & IIf(a = 1, Replace(RsOut!GroundHeight, ",", "."), Replace(RsOut!GroundHeightFinal, ",", ".")) & _
'         " WHERE GEOM_ID=" & RsInt!geom_id
'      RsInt.MoveNext
'
'   Next
'   RsInt.Close
'   RsOut.Close
'   Set RsInt = Nothing
'   Set RsOut = Nothing
'
'Trata_Erro:
'    If Err.Number = 0 Or Err.Number = 20 Then
'       Resume Next
'    Else
'       'xxxxx mantendo o código antigo 01/08/2008 xxxxx
'       If Not (RsOut Is Nothing) Then
'           If RsOut.State = adStateOpen Then RsOut.Close
'       End If
'       If Not (RsInt Is Nothing) Then
'          If RsInt.State = adStateOpen Then RsInt.Close
'       End If
'       Set RsInt = Nothing
'       Set RsOut = Nothing
'       'xxxxx mantendo o código antigo 01/08/2008 FIM xxxxx
'       Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'       Print #1, Now & " " & strUser & " " & Versao_Geo & " - clsGeoReference - Public Function UpdateTextInPoint - " & Err.Number & " - " & Err.Description
'       Close #1
'       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'    End If
'
'End Function

Public Function UpdateTextsInLines(LayerName_ As String, object_id As String)
   On Error GoTo Trata_Erro
   
   '####### DECLARAÇÃO DE VARIAVEIS LOCAIS ##########
   Dim rsTexts As ADODB.Recordset, rsWLines As ADODB.Recordset, LayId As Integer
   Dim comprimento As Double, Material As String, Diametro As String, CotaI As Double, CotaF As Double
   '####### ROTINA  ##############
   Dim L5 As String

   LayId = GetLayerID(LayerName_)
   
   L5 = str(LayId)
   L5 = Trim(L5)
  If frmCanvas.TipoConexao <> 4 Then
   Set rsTexts = Conn.execute(convertQuery("SELECT * From Texts" & LayId & " where object_id='" & object_id & "' Order by geom_id", CInt(typeconnection)))
   Else
   Dim au As String
    Dim ae As String
     Dim ai As String
      Dim ao As String
      Dim uu As String
      uu = str(LayId)
      uu = Trim(uu)
      au = "texts"
      ae = "object_id"
      ai = "geom_id"
      
    'Set rsTexts = Conn.execute(convertQuery("SELECT * From " + """" + au + uu + """" + " where " + """" + ae + """" + "='" & object_id & "' Order by " + """" + ai + """" + "", CInt(typeconnection)))
 
   Set rsTexts = New ADODB.Recordset
     rsTexts.Open "SELECT * From " + """" + au + L5 + """" + " where " + """" + ae + """" + "='" & object_id & "' Order by " + """" + ai + """" + "", Conn, adOpenDynamic, adLockOptimistic

  
   
   
   End If
   
   
   If Not rsTexts.EOF Then
      
      If frmCanvas.TipoConexao <> 4 Then
      
      Set rsWLines = Conn.execute("SELECT * From  " & LayerName_ & " left JOIN x_Material on material=materialid where object_id_='" & object_id & "'")
      Else
     Dim ay As String
    Dim ax As String
     Dim az As String
     Dim aa As String
     ax = "X_MATERIAL"
      ay = "OBJECT_ID_"
      az = "MATERIALID"
     aa = "MATERIAL"
        Set rsWLines = Conn.execute("SELECT * From  " + """" + LayerName_ + """" + " left JOIN " + """" + ax + """" + " on " + """" + aa + """" + "=" + """" + az + """" + " where " + """" + ay + """" + "='" & object_id & "'")
        
      End If
      
      
      Diametro = Chr$(216) & " " & IIf(IsNull(rsWLines!INTERNALDIAMETER), 0, rsWLines!INTERNALDIAMETER)
      
      Material = IIf(IsNull(rsWLines!MATERIALNAME), "", rsWLines!MATERIALNAME)
      If frmCanvas.TipoConexao = 4 Then
      
      If Material = "" Then
      Material = "0"
      End If
      End If
      comprimento = IIf(rsWLines!Length = 0, rsWLines!LENGTHCALCULATED, rsWLines!Length)
      
         If LayerName_ = "SEWERLINES" Then
            CotaI = rsWLines!INITIALTUBEDEEPNESS
            CotaF = rsWLines!FINALTUBEDEEPNESS

         Else
            CotaI = rsWLines!INITIALGROUNDHEIGHT
            CotaF = rsWLines!FINALGROUNDHEIGHT
         
         End If
      
a = "texts"
b = LayId
c = "b"
d = "text_value"
e = CotaI
f = CotaF
g = comprimento

ii = "geom_id"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"


  
     If GetTypeText(LayerName_) = 2 Then  'Rede Esgoto / Drenagem
         'Inclinação
         If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "Update Texts" & LayId & " Set text_value = '" & Round(((CotaI - CotaF) / comprimento) * 100, 2) & "%" & "' Where geom_id=" & rsTexts.Fields("geom_id")
         Else
         a = "texts"
b = LayId
c = "b"
d = "text_value"
e = CotaI
f = CotaF
g = comprimento

ii = "geom_id"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"
         'MsgBox "Update " + """" + a + LayId + """" + " Set " + """" + d + """" + "='" & Round(((CotaI - CotaF) / comprimento) * 100, 2) & "%" & "'  Where " + """" + ii + """" + "='" & rsTexts.Fields("geom_id") & "'"
        Conn.execute "Update " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + "='" & Round(((CotaI - CotaF) / comprimento) * 100, 2) & "%" & "'  Where " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         End If
         rsTexts.MoveNext
         'Diamentro
          If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "Update Texts" & LayId & " Set text_value = '" & Diametro & "' Where geom_id=" & rsTexts.Fields("geom_id")
         Else
         a = "texts"
b = LayId
c = "b"
d = "text_value"
e = CotaI
f = CotaF
g = comprimento

ii = "geom_id"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"
         'MsgBox "Update " + """" + a + LayId + """" + " Set " + """" + d + """" + " = '" & Diametro & "' Where " + """" + ii + """" + "='" & rsTexts.Fields("geom_id") & "'"
         Conn.execute "Update " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " = '" & Diametro & "' Where " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         End If
         rsTexts.MoveNext
         'Material
          If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "Update Texts" & LayId & " Set text_value = '" & Material & "' Where geom_id=" & rsTexts.Fields("geom_id")
         Else
         a = "texts"
b = LayId
c = "b"
d = "text_value"
e = CotaI
f = CotaF
g = comprimento

ii = "geom_id"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"
         'MsgBox "Update " + """" + a + """" + LayId + """" + " Set " + """" + d + """" + " = '" & Material & "' Where " + """" + ii + """" + "='" & rsTexts.Fields("geom_id") & "'"
         Conn.execute "Update " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " = '" & Material & "' Where " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         End If
         rsTexts.MoveNext
         'Comprimento
         If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "Update Texts" & LayId & " Set text_value = '" & comprimento & "' Where geom_id='" & rsTexts.Fields("geom_id") & "'"
         
         Else
         a = "texts"
b = LayId
c = "b"
d = "text_value"
e = CotaI
f = CotaF
g = comprimento

ii = "geom_id"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"

         'MsgBox "Update" + """" + "texts" + LayId + """" + " Set " + """" + "text_value" + """" + " = '" & comprimento & "' Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         Conn.execute "Update" + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " = '" & comprimento & "' Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         End If
         rsTexts.MoveNext
         

a = "texts"
b = LayId
c = "b"
d = "TEXT_VALUE"
e = Round(CotaI, 2)
f = "'e'"
g = rsTexts.Fields("geom_id")
h = "'g'"
ii = "GEOM_ID"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"

            'Cota Inicial
             If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "Update Texts" & LayId & " Set text_value = '" & Round(CotaI, 2) & "' Where geom_id=" & rsTexts.Fields("geom_id")
            Else
            Conn.execute "Update " + """" + "texts" + L5 + """" + "  Set " + """" + "text_value" + """" + " = '" & Round(CotaI, 2) & "' Where " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
            End If
            rsTexts.MoveNext
            'Cota Final
a = "texts"
b = LayId
c = "b"
d = "TEXT_VALUE"
e = Round(CotaF, 2)
f = "'e'"
g = rsTexts.Fields("geom_id")
h = "'g'"
ii = "GEOM_ID"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"
            If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "Update Texts" & LayId & " Set text_value = '" & Round(CotaF, 2) & "' Where geom_id=" & rsTexts.Fields("geom_id")
            Else
            Conn.execute "Update " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " =  '" & Round(CotaF, 2) & "' Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") + "'"
            End If
         
a = "texts"
b = LayId
c = "b"
d = "TEXT_VALUE"
e = Round(CotaI, 2)
f = "'e'"
g = rsTexts.Fields("geom_id")
h = "'g'"

         If Not rsTexts.EOF Then
            rsTexts.MoveNext
            If Not rsTexts.EOF Then
             If frmCanvas.TipoConexao <> 4 Then
               Conn.execute "Update Texts" & LayId & " Set text_value = '" & Round(CotaI, 2) & "' Where geom_id=" & rsTexts.Fields("geom_id")
               Else
               Conn.execute "Update " + """" + "texts" + L5 + """" + "  Set " + """" + "TEXT_VALUE" + """" + " =  '" & Round(CotaI, 2) & "' Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
               End If
            End If
         End If
         If Not rsTexts.EOF Then
            rsTexts.MoveNext
a = "Texts"
b = LayId
c = "b"
d = "TEXT_VALUE"
e = Round(CotaF, 2)
f = "'e'"
g = rsTexts.Fields("geom_id")
h = "'g'"
            If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "Update Texts" & LayId & " Set text_value = '" & Round(CotaF, 2) & "' Where geom_id=" & rsTexts.Fields("geom_id")
            Else
            Conn.execute "Update " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " =  '" & Round(CotaF, 2) & "' Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
            End If
         End If
      Else 'Rede de Agua
         'Diametro
         'Dim nSQL As String
a = "texts"
b = LayId
c = "b"
d = "TEXT_VALUE"
e = Round(CotaF, 2)
f = "'e'"
g = rsTexts.Fields("geom_id")
h = "'g'"
         'nSQL = "Update Texts" & LayId & " Set text_value = '" & Diametro & "' Where geom_id= " & rsTexts.Fields("geom_id")
         If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "Update Texts" & LayId & " Set text_value = '" & Diametro & "' Where geom_id= " & rsTexts.Fields("geom_id")
         Else
           Conn.execute "Update " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " = '" & Diametro & "'" + " Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         End If
         rsTexts.MoveNext
         'Material
         If frmCanvas.TipoConexao <> 4 Then
         Conn.execute "Update Texts" & LayId & " Set text_value = '" & Material & "' Where geom_id= " & rsTexts.Fields("geom_id")
         Else
          Conn.execute "Update  " + """" + "texts" + L5 + """" + " Set " + """" + "text_value" + """" + " = '" & Material & "'" + " Where  " + """" + "geom_id" + """" + "='" & rsTexts.Fields("geom_id") & "'"
         End If
         rsTexts.MoveNext
         'Comprimento
     
      End If
      
      rsWLines.Close
   End If
   Set rsWLines = Nothing
   rsTexts.Close
   Set rsTexts = Nothing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 3021 Then
       Resume Next
    Else
       
      PrintErro "clsGeoReference", "Public Function UpdateTextsInLines", CStr(Err.Number), CStr(Err.Description), True
       
    End If
   
End Function
' Calcula os posicionamentos dos textos na linha e insere os mesmos
'
' geom_id - da linha
' object_id - da linha
' Diametro - do trecho de rede
' Material - do trecho de rede
' comprimento - do trecho de rede
' InsCotas - se serão colocadas as cotas
' CotaI - cota inicial do trecho
' CotaF - cota final do trecho
'
Public Function InsertTextInDatabase(geom_id As Long, object_id As String, _
                  Diametro As String, Material As String, comprimento As Double, Optional InsCotas As Boolean = False, Optional CotaI As Double, Optional CotaF As Double) As Boolean
    On Error GoTo Trata_Erro
    '####### DECLARAÇÃO DE VARIAVEIS LOCAIS ##########
    'On Error GoTo InsertTextInDatabase_Err
    Dim SEG As Integer, QtdePontos As Integer, ANGULO As Double, Side As TeGeometrySide
    Dim X As Double, Y As Double, l As Double, Distancia As Integer
    '####### ROTINA DE INSERÇÃO ##############
    
    With tdb
        'retorna em x, y as coordenadas da bissetriz da cadeia de linhas (meio dela),
        'para sabermos onde colocaremos os textos da linha (comprimento, etc.)
        .getCenterGeometry geom_id, object_id, tpLINES, X, Y
        'obtem a quantidade de pontos (vértices) em uma cadeia de linhas
        QtdePontos = .getQuantityPointsLine(geom_id, object_id)
        'divide por dois para saber em que segmento de linha iremos inserir o texto. se a divisao nao for exata pegaremos o segmento mais a esquerda.
        SEG = CInt(QtdePontos / 2)
        'obtem o angulo do segmento escolhido
        ANGULO = .getAngle(geom_id, object_id, SEG, 2)
        .getLengthOfLine geom_id, object_id, l
        'obtem a informação se a linha foi desenhada da direita para esquerda, ou da esquerda para direita, indefinido ou leftRight (???)
        Side = .getSide(X, Y, object_id, geom_id, tpLINES)
        If Side = TeRight Then
            Distancia = 6
        Else
            Distancia = 6
            'Distancia = -1
        End If
        If Not InsCotas Then
            '--- Diametro
            ' obtem as coords. x e y perpendiculares a linha (object_id), no meio da linha + 1 metro (L/2+1),
            ' a uma distância (-Distancia) perpendicular a linha
            'L = comprimento
            If l > 10 Then 'textos do diametro e material um ao lado do outro
                If ANGULO <= 90 Or ANGULO > 270 Then 'quadrante 1 e 4
                    .getPerpendicularPoint geom_id, object_id, (l / 3) * 2, 2, X, Y
                ElseIf ANGULO = 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l / 3), -2, X, Y
                Else
                    .getPerpendicularPoint geom_id, object_id, (l / 3), -2, X, Y
                End If
                .InsertTextInDatabase object_id, Diametro, 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
                ' --- Material em baixo, esquerda
                If ANGULO <= 90 Or ANGULO > 270 Then 'Q1 Q4
                    .getPerpendicularPoint geom_id, object_id, (l / 3), 2, X, Y
                ElseIf ANGULO = 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l / 3) * 2, -2, X, Y
                Else 'Q2 Q3
                    .getPerpendicularPoint geom_id, object_id, (l / 3) * 2, -2, X, Y
                End If
                .InsertTextInDatabase object_id, IIf(Material = "0", "", Material), 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
            Else
                If ANGULO <= 90 Or ANGULO > 270 Then 'quadrante 1 e 4
                    .getPerpendicularPoint geom_id, object_id, (l / 2), 2, X, Y
                ElseIf ANGULO = 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l / 2), -2, X, Y
                Else
                    .getPerpendicularPoint geom_id, object_id, (l / 2), -2, X, Y
                End If
                .InsertTextInDatabase object_id, Diametro, 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
                ' --- Material em baixo, esquerda
                If ANGULO <= 90 Or ANGULO > 270 Then 'Q1 Q4
                    .getPerpendicularPoint geom_id, object_id, (l / 2), 3.5, X, Y
                ElseIf ANGULO = 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l / 2), -3.5, X, Y
                Else 'Q2 Q3
                    .getPerpendicularPoint geom_id, object_id, (l / 2), -3.5, X, Y
                End If
                .InsertTextInDatabase object_id, IIf(Material = "0", "", Material), 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
            End If
            '--- Comprimento em cima da linha , no meio
            If ANGULO <= 90 Or ANGULO > 270 Then 'Q1 Q4
            .getPerpendicularPoint geom_id, object_id, (l / 2), -0.5, X, Y
            ElseIf ANGULO = 270 Then
            .getPerpendicularPoint geom_id, object_id, (l / 2), 0.5, X, Y
            Else 'Q2 Q3
            .getPerpendicularPoint geom_id, object_id, (l / 2), 0.5, X, Y
            End If
            .InsertTextInDatabase object_id, Round(comprimento, 2), 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
        Else
            '--- Inclinação a esquerda
            If ANGULO <= 90 Or ANGULO > 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 3), 2, X, Y
            ElseIf ANGULO = 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 3) * 2, -2, X, Y
            Else
                .getPerpendicularPoint geom_id, object_id, (l / 3) * 2, -2, X, Y
            End If
            If Not (CotaI = 0 Or CotaF = 0) Then
                .InsertTextInDatabase object_id, Round(((CotaI - CotaF) / comprimento) * 100, 2) & "%", 1, X, Y, ANGULO, , , True
            Else
                .InsertTextInDatabase object_id, 0 & "%", 1, X, Y, ANGULO, , , True
            End If
            '--- Diametro em baixo, direita
            If ANGULO <= 90 Or ANGULO > 270 Then 'quadrante 1 e 4
                .getPerpendicularPoint geom_id, object_id, (l / 3) * 2, 2, X, Y
            ElseIf ANGULO = 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 3), -2, X, Y
            Else
                .getPerpendicularPoint geom_id, object_id, (l / 3), -2, X, Y
            End If
            .InsertTextInDatabase object_id, Diametro, 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
            '--- Material em baixo, meio
            If ANGULO <= 90 Or ANGULO > 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 2), 2, X, Y
            ElseIf ANGULO = 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 2), -2, X, Y
            Else
                .getPerpendicularPoint geom_id, object_id, (l / 2), -2, X, Y
            End If
            .InsertTextInDatabase object_id, IIf(Material = "0", "", Material), 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
            '--- Comprimento em cima da linha , no meio
            If ANGULO <= 90 Or ANGULO > 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 2), -0.5, X, Y
            ElseIf ANGULO = 270 Then
                .getPerpendicularPoint geom_id, object_id, (l / 2), 0.5, X, Y
            Else
                .getPerpendicularPoint geom_id, object_id, (l / 2), 0.5, X, Y
            End If
            .InsertTextInDatabase object_id, Round(comprimento, 2), 1, X, Y, IIf(ANGULO = 270, 90, ANGULO), , , True
            With tdb
                '--- CotaInicial
                If ANGULO <= 90 Or ANGULO > 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l * 0.15), -0.5, X, Y
                ElseIf ANGULO = 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l * 0.15), 0.5, X, Y
                Else
                    .getPerpendicularPoint geom_id, object_id, (l * 0.15), 0.5, X, Y
                End If
                .InsertTextInDatabase object_id, CotaI, 1, X, Y, ANGULO, , , True
                '--- CotaFinal
                If ANGULO <= 90 Or ANGULO > 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l * 0.85), -0.5, X, Y
                ElseIf ANGULO = 270 Then
                    .getPerpendicularPoint geom_id, object_id, (l * 0.85), 0.5, X, Y
                Else
                    .getPerpendicularPoint geom_id, object_id, (l * 0.85), 0.5, X, Y
                End If
                .InsertTextInDatabase object_id, CotaF, 1, X, Y, ANGULO, , , True
            End With
        End If
    End With
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsGeoReference", "Public Function InsertTextInDatabase", CStr(Err.Number), CStr(Err.Description), True
    End If
    'InsertTextInDatabase_Err:
    'MsgBox "Não foi possível inserir o texto automaticamente,contate o suporte:(011) 381602034 - Nexus ", vbExclamation, "Atenção"
End Function
' Obtem a identificação do layer
' Retorna o número de identificação do layer
'
' LayerName_ - nome do layer
'
Public Function GetLayerID(LayerName_ As String) As Integer
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim a As String
    Dim b As String
    Dim c As String
    
    a = "layer_id"
    b = "te_layer"
    c = "name"
    If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute(convertQuery("SELECT Layer_id from Te_Layer where name='" & LayerName_ & "' or name='" & UCase(LayerName_) & "'", CInt(typeconnection)))
        If rs.EOF = False Then
            GetLayerID = rs(0).value            'obtem o número do layer (id)
        End If
        rs.Close
        Set rs = Nothing
    Else
        Set rs = Conn.execute(convertQuery("SELECT " + """" + a + """" + " from " + """" + b + """" + " where " + """" + c + """" + "='" & LayerName_ & "'", CInt(typeconnection)))
        If rs.EOF = False Then
            GetLayerID = rs(0).value            'obtem o número do layer (id)
        End If
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsGeoReference", "Public Function GetLayerID", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

Public Function GetLayerAttrib(LayerName_ As String, layer_id As Integer, Attr_Link As String) As Boolean
On Error GoTo Trata_Erro
   Dim rs As ADODB.Recordset
   Set rs = Conn.execute(convertQuery("SELECT l.Layer_id, t.attr_link from Te_Layer_Table t inner join te_layer l on l.layer_id=t.layer_id where l.name='" & LayerName_ & "'", CInt(typeconnection)))
   If rs.EOF Then
      GetLayerAttrib = False
   Else
      layer_id = rs(0).value
      Attr_Link = rs(1).value
      GetLayerAttrib = True
   End If
   rs.Close
   Set rs = Nothing

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
      PrintErro "clsGeoReference", "Public Function GetLayerAttrib", CStr(Err.Number), CStr(Err.Description), True
       
    End If

End Function

Public Function DeleteTextObjectWithInsertText(LayerName As String, geom_id As Long, Object_id_ As String) As Boolean
On Error GoTo Trata_Erro
   Dim rs As ADODB.Recordset, IDLayer As Integer
   Dim gus As String
   Dim Diameter As String, Lenght As Double, Material As String
a = "layer_id"
b = "te_layer"
c = "name"
d = "OBJECT_ID"
e = Object_id_
f = "'e'"
g = "X_MATERIAL"
h = "MATERIALID"
ii = "MATERIAL"
   If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute(convertQuery("SELECT layer_id from te_layer where name='" & LayerName & "'", CInt(typeconnection)))
   Else
   Set rs = Conn.execute(convertQuery("SELECT " + """" + "layer_id" + """" + " from " + """" + "te_layer" + """" + " where " + """" + "name" + """" + "='" & LayerName & "'", CInt(typeconnection)))
   
   End If
   If Not rs.EOF Then
      IDLayer = rs(0)
      If frmCanvas.TipoConexao = 4 Then
      gus = str(IDLayer)
      gus = Trim(gus)
      End If
      
      rs.Close
      Set rs = Nothing
      'problema
      'A linha abaixo foi comentada para correção no dia 31/05/2007
      'Conn.execute "DeleteTXTVisual " & IDLayer & "," & object_id_
      DeleteTXTVisual IDLayer, Object_id_
      'Conn.execute "Delete From texts" & IDLayer & "_txvisual where geom_id in(SELECT geom_id from texts" & IDLayer & " where object_id = '" & object_id_ & "')"
      
      
      
      
a = "texts"
b = IDLayer
c = "b"
d = "object_id"
e = "OBJECT_ID_"
f = "'e'"
g = "X_MATERIAL"
h = "MATERIALID"
ii = "MATERIAL"
Dim zz As String
Dim za As String
za = a + b
zz = Trim(za)
      If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "Delete From Texts" & IDLayer & " Where object_id = " & Object_id_
      Else
     ' MsgBox "Delete From " + """" + zz + """" + " Where " + """" + d + """" + " =  '" & Val(Object_id_) & "'"
      Conn.execute "Delete From " + """" + "texts" + gus + """" + " Where " + """" + d + """" + " =  '" & Val(Object_id_) & "'"
      End If
      
      If frmCanvas.TipoConexao <> 4 Then
      Set rs = Conn.execute("SELECT * from " & LayerName & " Left Join X_Material on MaterialId=Material where object_id_='" & Object_id_ & "'")
      Else
          Set rs = Conn.execute("SELECT * from " + """" + LayerName + """" + " Left Join" + """" + g + """" + " on " + """" + h + """" + "=" + """" + ii + """" + " where " + """" + e + """" + "='" & Val(Object_id_) & "'")
      End If
      
      
  
      If Not rs.EOF Then
         Diameter = Chr$(216) & " " & rs.Fields("InternalDiameter")
         Lenght = IIf(rs.Fields("length") = 0, rs.Fields("lengthcalculated"), rs.Fields("length"))
         Material = IIf(IsNull(rs.Fields("MaterialName")), "0", rs.Fields("MaterialName"))
         If UCase(LayerName) = "WATERLINES" Then
            InsertTextInDatabase geom_id, Object_id_, Diameter, Material, Lenght
         Else
            InsertTextInDatabase geom_id, Object_id_, Diameter, Material, Lenght, True, rs!INITIALTUBEDEEPNESS, rs!FINALTUBEDEEPNESS
         End If
      End If
   End If
   rs.Close
   Set rs = Nothing
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
       
      PrintErro "clsGeoReference", "Public Function DeleteTextObjectWithInsertText", CStr(Err.Number), CStr(Err.Description), True
   
   End If
   
End Function
' Apaga uma rede de água, esgoto ou drenagem, eliminando os trechos e nós das redes
'
'
Public Function DeleteRede(LayerName As String, object_id As String)
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset, RsCount As ADODB.Recordset, strsql As String
    Dim LayId As Integer, LayIdRef As Integer
    Dim localErro As String

    localErro = "0"
    If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute(convertQuery("SELECT Layer_id from Te_Layer where name='" & GetReferenceLayer(LayerName) & "'", CInt(typeconnection)))
    Else
        Dim aa As String
        Dim ab As String
        Dim ac As String
        aa = "layer_id"
        ab = "te_layer"
        ac = "name"
        Set rs = Conn.execute(convertQuery("SELECT " + """" + aa + """" + " from " + """" + ab + """" + " where " + """" + ac + """" + "='" & GetReferenceLayer(LayerName) & "'", CInt(typeconnection)))
    End If
    localErro = "1"
    LayId = rs(0).value
    If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute(convertQuery("SELECT Layer_id from Te_Layer where name='" & LayerName & "'", CInt(typeconnection)))
    Else
        Set rs = Conn.execute(convertQuery("SELECT " + """" + aa + """" + " from " + """" + ab + """" + " where " + """" + ac + """" + "='" & LayerName & "'", CInt(typeconnection)))
    End If
    LayIdRef = rs(0).value
    a = "INITIALCOMPONENT"
    b = "FINALCOMPONENT"
    c = "OBJECT_ID_"
    If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute("SELECT InitialComponent, FinalComponent From " & LayerName & _
        " Where Object_id_ = " & object_id)
    Else
        Set rs = Conn.execute("SELECT " + """" + "INITIALCOMPONENT" + """" + ", " + """" + "FINALCOMPONENT" + """" + " From " + """" + LayerName + """" + "Where " + """" + "OBJECT_ID_" + """" + " = '" & object_id & "'")
    End If
    localErro = "2"
    If Not rs.EOF Then
        If frmCanvas.TipoConexao <> 4 Then
            localErro = "2A"
            Set RsCount = Conn.execute("SELECT count(*) as " + """" + "Qde" + """" + " From " & LayerName & _
            " Where InitialComponent=" & rs.Fields("InitialComponent") & _
            " or FinalComponent = " & rs.Fields("InitialComponent"))
            If RsCount!Qde = 1 Then
                strsql = "'" & rs!INITIALCOMPONENT & "'"
            End If
            RsCount.Close
            Set RsCount = Conn.execute("SELECT count(*) as " + """" + "Qde" + """" + " From " & LayerName & _
            " Where InitialComponent=" & rs.Fields("FinalComponent") & _
            " or FinalComponent = " & rs.Fields("FinalComponent"))
            If RsCount!Qde = 1 Then
                strsql = strsql & IIf(strsql <> "", ",", "") & "'" & rs!FinalComponent & "'"
            End If
            RsCount.Close
        Else
            localErro = "2B"
            Dim aq As String
            aq = LayerName
            Set RsCount = Conn.execute("SELECT count(*) as " + """" + "Qde" + """" + " From " + """" + aq + """" + _
            " Where " + """" + a + """" + "='" & rs.Fields("INITIALCOMPONENT") & "'" + _
            " or " + """" + b + """" + " = '" & rs.Fields("INITIALCOMPONENT") & "'")
            If RsCount!Qde = 1 Then
                strsql = "'" & rs!INITIALCOMPONENT & "'"
            End If
            RsCount.Close
            Set RsCount = Conn.execute("SELECT count(*) as " + """" + "Qde" + """" + " From " + """" + aq + """" + _
            " Where " + """" + a + """" + "='" & rs.Fields("FINALCOMPONENT") & "'" + _
            " or " + """" + b + """" + " = '" & rs.Fields("FINALCOMPONENT") & "'")
            If RsCount!Qde = 1 Then
                strsql = strsql & IIf(strsql <> "", ",", "") & "'" & rs!FinalComponent & "'"
            End If
            RsCount.Close
        End If
    End If
    rs.Close
    localErro = "3"
    Set rs = Nothing
    Dim comandDelete As String
    Dim CMD_Delete As String
    Dim RefLayer As String
    RefLayer = GetReferenceLayer(LayerName)
    If strsql <> "" Then
        a = RefLayer
        b = "+a+"
        c = "OBJECT_ID_"
        d = "points"
        e = LayId
        f = "e"
        g = "DATA"
        h = "texts"
        ii = "object_id"
        Dim d2 As String
        d2 = "object_id_"
        If frmCanvas.TipoConexao <> 4 Then
            localErro = "3A"
            CMD_Delete = "Delete From " & RefLayer & "data where object_id_ In (" & strsql & ")"
            Conn.execute (CMD_Delete)
            CMD_Delete = "Delete From " & RefLayer & " where object_id_ In (" & strsql & ")"
            Conn.execute (CMD_Delete)
            CMD_Delete = "Delete From points" & LayId & " where object_id In (" & strsql & ")"
            Conn.execute (CMD_Delete)
            tdb.setCurrentLayer RefLayer
            'VERIFICA SE OS NÓS POSSUEM TEXTOS PARA SEREM DELETADOS
            If Not tdb.existsRepresentation(128) = 0 Then
                CMD_Delete = "Delete From texts" & LayId & " where object_id In (" & strsql & ")"
                Conn.execute (CMD_Delete)
            End If
        Else
            localErro = "3B"
            CMD_Delete = "Delete From " + """" + a + g + """" + " where " + """" + c + """" + " In (" & strsql & ")"
            Conn.execute (CMD_Delete)
            CMD_Delete = "Delete From " + """" + a + """" + " where " + """" + "OBJECT_ID_" + """" + " In (" & strsql & ")"
            Conn.execute (CMD_Delete)
            CMD_Delete = "Delete From " + """" + d + e + """" + " where " + """" + ii + """" + " In (" & strsql & ")"
            Conn.execute (CMD_Delete)
            tdb.setCurrentLayer RefLayer
            'VERIFICA SE OS NÓS POSSUEM TEXTOS PARA SEREM DELETADOS
            If Not tdb.existsRepresentation(128) = 0 Then
                CMD_Delete = "Delete From " + """" + h + e + """" + " where " + """" + ii + """" + " In (" & strsql & ")"
            End If
            Conn.execute (CMD_Delete)
        End If
    End If
    localErro = "4"
    If frmCanvas.TipoConexao <> 4 Then
        Conn.execute "Delete From " & LayerName & "Data where object_id_ = '" & object_id & "'"
        Conn.execute "Delete From " & LayerName & " where object_id_ = '" & object_id & "'"
        DeleteTXTVisual LayIdRef, object_id
        Conn.execute "Delete From texts" & LayIdRef & " where object_id = '" & object_id & "'"
        Conn.execute "Delete From lines" & LayIdRef & " where object_id = '" & object_id & "'"
    Else
        localErro = "5"
        Dim d1 As String
        d1 = "object_id"
        a = LayerName
        b = "DATA"
        c = LayIdRef
        e = "c"
        d = "OBJECT_ID_"
        f = "texts"
        g = "lines"
        CMD_Delete = "Delete From " + """" + a + b + """" + " where " + """" + "OBJECT_ID_" + """" + " = '" & object_id & "' "
        Conn.execute (CMD_Delete)
        Conn.execute "Delete From " + """" + a + """" + " where " + """" + d + """" + " = '" & object_id & "'"
        DeleteTXTVisual LayIdRef, object_id
        f = "texts"
        c = LayIdRef
        CMD_Delete = "Delete From " + """" + f + c + """" + " where " + """" + d1 + """" + " = '" & object_id & "' "
        Conn.execute (CMD_Delete)
        c = LayIdRef
        g = "lines"
        CMD_Delete = "Delete From " + """" + g + c + """" + " where " + """" + d1 + """" + " = '" & object_id & "' "
        Conn.execute (CMD_Delete)
    End If

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsGeoReference - Local Erro:" + localErro, "Public Function DeleteRede", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function
'###############################################################
' Autor     : Luis Claudio Rodrigues Domingues    Data: 20/12/05
' Nome      : SELECTRede
' Descrição : Colore a rede até encontrar válvula
'###############################################################
'a partir do object_id selecionado pelo usuário ele vai localizar todos os trechos de redes até encontrar registros. Válvulas (VRPs) e registros fixos (divisa de setor de abastecimento), não são considerados. Ele procura somente até encontrar um nó do tipo REGISTRO
'
'
Public Function SELECTRede(Object_id_ As String)
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim iContA As Integer, iContB As Integer
    Dim MyLocalLines() As String        'object_id s do vetor de trechos de redes que estão conectados aos nós inicial e final do último trecho de rede pesquisado
    Dim iLn As Integer, iPnt
    Dim MyComponents(1) As String
    Dim ACHOU As Boolean
    Dim rsType As ADODB.Recordset
    Dim fg As String
    Dim fh As String
    Dim fi As String
    Dim fj As String
    Dim fl As String
    Dim fm As String
    Dim fn As String
    Dim fo As String
    Dim ga As String
    Dim ga2 As String
    
    Set rsType = New ADODB.Recordset
    'Obtem todos os tipos de componentes de rede (nós da rede) que podem ser cadastrados
    If frmCanvas.TipoConexao <> 4 Then
        'Caso seja SQLServer ou Oracle
        rsType.Open "SELECT ID_TYPE,DESCRIPTION_,SPECIFICATION_ FROM WATERCOMPONENTSTYPES", Conn, adOpenDynamic, adLockReadOnly
    Else
        fg = "ID_TYPE"
        fh = "DESCRIPTION_"
        fi = "SPECIFICATION_"
        fj = "WATERCOMPONENTSTYPES"
        fl = "INITIALCOMPONENT"
        fm = "FINALCOMPONENT"
        fn = "WATERLINES"
        fo = "OBJECT_ID_"
        ga = "WATERLINES"
        ga2 = "WATERCOMPONENTS"
        'MsgBox "SELECT " + """" + fg + """" + "," + """" + fh + """" + "," + """" + fi + """" + " FROM " + """" + fj + ""
        rsType.Open "SELECT " + """" + fg + """" + "," + """" + fh + """" + "," + """" + fi + """" + " FROM " + """" + fj + """" + "", Conn, adOpenDynamic, adLockOptimistic
    End If
    'verifica se a tabela contém registros, caso contrário não é possível realizar esta funcionalidade, existe erro na base de dados
    If rsType.EOF = True Then
        MsgBox "Não encontrou os tipos de componentes de rede existentes na tabela WATERCOMPONENTSTYPES. Saindo da funcionalidade."
        Exit Function
    End If
    'obtem o número do nó inicial e o número do nó final do trecho de rede de água selecionado
    If frmCanvas.TipoConexao <> 4 Then
        'se SQLServer ou Oracle
        Set rs = Conn.execute("SELECT InitialComponent,FinalComponent from waterlines where object_id_='" & Object_id_ & "'")
    Else
        'se Postgres
        Set rs = Conn.execute("SELECT " + """" + fl + """" + "," + """" + fm + """" + " from " + """" + ga + """" + " where " + """" + fo + """" + "='" & Object_id_ & "'")
    End If
    If Not stopProcess Then
        DoEvents
        If Not rs.EOF Then
            MyComponents(0) = rs(0)         'número do nó inicial do trecho de rede
            MyComponents(1) = rs(1)         'número do nó final do trecho de rede
            rs.Close                        'fecha a conexão que obteve os nós do trecho de rede
            Set rs = Nothing
            'processa para os dois nós
            For iContA = 0 To 1
continuaFor:
                'obtem o tipo de nó a partir do nó inicial (primeira passada pelo For), e a partir do nó final (segunda passada pelo For), para saber se é um registro e deve parar a procura
                If frmCanvas.TipoConexao <> 4 Then
                    'se SQLServer ou Oracle
                    Set rs = Conn.execute("SELECT iD_type from watercomponents where object_id_='" & MyComponents(iContA) & "'")
                Else
                    'se Postgres
                    Set rs = Conn.execute("SELECT " + """" + fg + """" + " from " + """" + ga2 + """" + " where " + """" + fo + """" + "='" & MyComponents(iContA) & "'")
                End If
                'verifica em WATERCOMPONENTS se encontrou o nó inicial ou final (dependendo da passagem do For) que estava especificado na tabela WATERLINES
                If Not rs.EOF Then
                    'encontrou o nó em WATERCOMPONENTS que WATERLINES dizia existir (inicial ou final dependendo da passagem do For)
                    ACHOU = False
                    rsType.Filter = "ID_TYPE='" & rs!id_Type & "'"
                    If rsType.EOF = False Then
                        If Trim(rsType!Specification_) = "REGISTER" Then       'verifica se é um ponto de parada
                            ACHOU = True
                        End If
                    End If
                    'O PROBLEMA ESTÁ AQUI POIS QUANDO ELE ACHA UMA VALVULA(VALVE) ELE ABANDONA SE PROCURAR O OUTRO LADO
                    If ACHOU = True Then 'valor original = 1 mas não há nada cadastrado como tipo 1
                        If iContA = 0 Then
                            iContA = 1
                            GoTo continuaFor:
                        End If
                        'Exit For
                    Else
                        rs.Close
                        Set rs = Nothing
                        'procura por outros trechos de redes os quais estão conectados aos nós inicial e final
                        If frmCanvas.TipoConexao <> 4 Then
                            Set rs = Conn.execute("SELECT object_id_ from waterlines where InitialComponent=" & MyComponents(iContA) & " or FinalComponent= " & MyComponents(iContA))
                        Else
                            ' Set rs = Conn.execute("SELECT " + """" + fo + """" + " from " + """" + fn + """" + " where " + """" + fl + """" + "='" & MyComponents(iContA) & "' or " + """" + fm + """" + "= '" & MyComponents(iContA)) & "'"
                            Set rs = Conn.execute("SELECT " + """" + "OBJECT_ID_" + """" + " from " + """" + "WATERLINES" + """" + " where " + """" + "INITIALCOMPONENT" + """" + "=" & MyComponents(iContA) & " or " + """" + "FINALCOMPONENT" + """" + "= " & MyComponents(iContA))
                        End If
                        While Not rs.EOF 'caso tenha encontrado outros trechos de rede (novos trechos) conectados as extremidades inicial e final
                            ReDim Preserve MyLocalLines(iLn)
                            MyLocalLines(iLn) = rs!Object_id_
                            iLn = iLn + 1
                            rs.MoveNext
                        Wend
                        iLn = iLn - 1
                        rs.Close
                        Set rs = Nothing
                        For iPnt = 0 To iLn
                            SubSELECTRede MyLocalLines(iPnt)
                        Next
                    End If
                Else
                    'não encontrou o nó em WATERCOMPONENTS, existe uma inconsistência no banco de dados
                    MsgBox "Não foi encontrado o nó em WATERCOMPONETS de número " & MyComponents(iContA) & "O trecho de rede número " & Object_id_ & " com nó inicial " & MyComponents(0) & " e nó final " & MyComponents(1) & " foi detectada uma inconsistência na base de dados. Rodar o ValidaBase e analisar o arquivo de log gerado pelo mesmo."
                    rs.Close
                    Set rs = Nothing
                End If
            Next
        Else
            rs.Close
            Set rs = Nothing
        End If
    End If
    If Not (rs Is Nothing) Then
        If rs.State = adStateOpen Then
            rs.Close
        End If
    End If
    Set rs = Nothing
    Exit Function

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      PrintErro "clsGeoReference", "Public Function DeleteTextObjectWithInsertText", CStr(Err.Number), CStr(Err.Description), True
   End If
End Function

Private Sub SubSELECTRede(Object_id_ As String)
    Dim a As Integer, naoChamar As Boolean
    If Not p Then
        ReDim Preserve MyLines(i)
        p = True
    End If
    For a = 0 To i
        If MyLines(a) = Object_id_ Then
            naoChamar = True
        End If
    Next
    If Not naoChamar Then
        FrmProcess.txtRecord = i
        FrmProcess.Refresh
        MyLines(i) = Object_id_
        If Object_id_ <> "" Then
            If object_ids = "" Then
                object_ids = "'" & Object_id_ & "'"
            Else
                object_ids = object_ids & ",'" & Object_id_ & "'"
            End If
      End If
      i = i + 1
      ReDim Preserve MyLines(i)
      tcs.setDetachedLineStyle 3, 1, RGB(255, 255, 0), True
      tcs.addDetachedIds tpLINES, , Object_id_
      SELECTRede Object_id_
   End If
End Sub


'###############################################################
' Autor     : Luis Claudio Rodrigues Domingues    Data: 20/12/05
' Nome      : GetDeclivity
' Descrição : Colori a rede seguindo a declividade
'###############################################################
Public Function GetDeclivity(LayerName As String, RefLayer As String, Obj As String)
   Dim rs As ADODB.Recordset
   Dim iContA As Integer, iContB As Integer
   Dim MyLocalLines() As String, iLn As Integer, iPnt
   Dim MyComponents(1) As String
   Dim cota(1) As Double
   Dim NextNo As String


a = "INITIALCOMPONENT"
b = "FINALCOMPONENT"
c = "INITIALGROUNDHEIGHT"
d = "FINALGROUNDHEIGHT"
e = "OBJECT_ID_"
   If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT InitialComponent,FinalComponent,InitialGroundHeight,FinalGroundHeight from " & LayerName & " where object_id_='" & Obj & "'")
   Else
    Set rs = Conn.execute("SELECT " + """" + a + """" + "," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + " from " + """" + LayerName + """" + " where " + """" + e + """" + "='" & Obj & "'")
   End If
   If Not rs.EOF Then
      MyComponents(0) = rs(0)
      MyComponents(1) = rs(1)
      cota(0) = rs!INITIALGROUNDHEIGHT
      cota(1) = rs!FINALGROUNDHEIGHT
      rs.Close
      Set rs = Nothing
      If frmCanvas.TipoConexao <> 4 Then
     ' Set rs = Conn.execute("SELECT object_id_, InitialGroundHeight,FinalGroundHeight from " & LayerName & " " & _
              ' "where (InitialComponent in (" & MyComponents(0) & "," & MyComponents(1) & ") " & _
              ' "Or FinalComponent in (" & MyComponents(0) & "," & MyComponents(1) & ")) " & _
              ' "and FinalGroundHeight >= " & Replace(cota(1), ",", ".") & " and InitialGroundHeight >= " & Replace(cota(0), ",", "."))
               
               Set rs = Conn.execute("SELECT object_id_, InitialGroundHeight,FinalGroundHeight from " & LayerName & " " & _
               "where (InitialComponent in (" & MyComponents(0) & "," & MyComponents(1) & ") " & _
               "Or FinalComponent in (" & MyComponents(0) & "," & MyComponents(1) & ")) " & _
               "and FinalGroundHeight > " & Replace(cota(1), ",", ".") & " and InitialGroundHeight > " & Replace(cota(0), ",", "."))
               
            
            '    WritePrivateProfileString "A", "A", SQL, App.path & "\DEBUG.INI"
             
               
               Else
                 Set rs = Conn.execute("SELECT " + """" + e + """" + "," + """" + c + """" + "," + """" + d + """" + " from " + """" + LayerName + """" + " " & _
               "where (" + """" + a + """" + " in ('" & MyComponents(0) & "','" & MyComponents(1) & "') " & _
               "Or " + """" + b + """" + " in ('" & MyComponents(0) & "','" & MyComponents(1) & "')) " & _
               "and " + """" + d + """" + " > '" & Replace(cota(1), ",", ".") & "' and " + """" + c + """" + " > '" & Replace(cota(0), ",", ".") & "'")
               
               End If
               
         

               
      tcs.setDetachedLineStyle 3, 1, RGB(255, 255, 0), True
      tcs.addDetachedIds tpLINES, , Obj
      While Not rs.EOF
         ReDim Preserve MyLocalLines(iLn)
         MyLocalLines(iLn) = rs!Object_id_
         iLn = iLn + 1
         rs.MoveNext
      Wend
      iLn = iLn - 1
      rs.Close
      Set rs = Nothing
      For iPnt = 0 To iLn
         SubSELECTDeclivity LayerName, RefLayer, MyLocalLines(iPnt)
      Next
   End If
   If Not (rs Is Nothing) Then
      If rs.State = adStateOpen Then rs.Close
   End If
   Set rs = Nothing
End Function

Private Sub SubSELECTDeclivity(LayerName As String, RefLayer As String, Object_id_ As String)
   Dim a As Integer, naoChamar As Boolean
   If Not p Then
      ReDim Preserve MyLines(i)
      p = True
   End If
   For a = 0 To i
     If MyLines(a) = Object_id_ Then naoChamar = True
   Next
   If Not naoChamar Then
     FrmProcess.txtRecord = i
     FrmProcess.Refresh
     MyLines(i) = Object_id_
     i = i + 1
     ReDim Preserve MyLines(i)
     GetDeclivity LayerName, RefLayer, Object_id_
   End If
End Sub

Public Function GetLinesValues(Obj As String, RSNos As ADODB.Recordset, mtype As Long) As String
   Dim rs As ADODB.Recordset, RsN As ADODB.Recordset
   Dim iContA As Integer, iContB As Integer
   Dim MyLocalLines() As String, iLn As Integer, iPnt
   Dim MyComponents(1) As String
   Dim cota(1) As Double
   Dim NextNo As String
   
   'If Obj = "7705" Then Stop
   a = "INITIALCOMPONENT"
b = "FINALCOMPONENT"
c = "INITIALGROUNDHEIGHT"
d = "FINALGROUNDHEIGHT"
e = "OBJECT_ID_"
f = "WATERLINES"
g = "WATERCOMPONENTS"
h = "CALCULENODE"
   If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT InitialComponent,FinalComponent,InitialGroundHeight,FinalGroundHeight from waterlines where object_id_='" & Obj & "'")
   Else
      Set rs = Conn.execute("SELECT " + """" + a + """" + "," + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + " from " + """" + f + """" + " where " + """" + e + """" + "='" & Obj & "'")
   
   End If
   If Not rs.EOF Then
      MyComponents(0) = rs(0)
      MyComponents(1) = rs(1)
      rs.Close
      Set rs = Nothing
      
      
      'Verifica de os nós encontrados são de calculo,
       If frmCanvas.TipoConexao <> 4 Then
        Set RsN = Conn.execute("SELECT Object_id_ from watercomponents where object_id_ in(" & _
        MyComponents(0) & "," & MyComponents(1) & ") and calculenode<>0")
        Else
        Set RsN = Conn.execute("SELECT " + """" + e + """" + " from " + """" + g + """" + " where " + """" + e + """" + " in('" & MyComponents(0) & "','" & MyComponents(1) & "') and ""+h+""<>'0'")
        
        End If
      'Esta parte do código tem a pretençao de verificar qual o components final do trecho
      'está com erro
      
      'Rotina que define o trecho
      If RSNos.EOF Then
         RSNos.AddNew
      End If
      If Not RsN.EOF Then
         
         If IsNull(RSNos!cpi) Then
            
            RSNos!cpi = RsN(0)
            RSNos!lin = Obj
            RsN.MoveNext
         End If
      End If
      
      If Not RsN.EOF Then
         If Not IsNull(RSNos!cpi) And (Not RsN(0) = RSNos.Fields("cpI").value) Then
            RSNos!cpf = RsN(0)
            RsN.MoveNext
            
         End If
      Else
         'Resolve aqui
         'quais os trechos ligados ao no que náo é o de calculo
         RSNos!cpf = GetSecundComponent(IIf(RSNos!cpi = MyComponents(0), MyComponents(1), MyComponents(0)), RSNos, mtype)
         
         'Se não for encontrado um nó final de calculo, este trecho deve é ignorado
         If RSNos!cpf = "" Then
            Exit Function
         End If
         
         If RSNos!cpi = MyComponents(0) Then
            MyComponents(1) = RSNos!cpf
         Else
            MyComponents(0) = RSNos!cpf
         End If
      End If
      
      tcs.setDetachedLineStyle 3, 1, RGB(255, 255, 0), True
      tcs.addDetachedIds tpLINES, , Obj
      tcs.redraw
      
      'retorna todos os trechos ligados aos nós do trechos recibido "obj"
      
      '        n1-----l1------n2
      '        |            / |
      '        |         /    |
      '        l4     l5      l2
      '        |    /         |
      '        | /            |
      '        n4-----l3------n3
      
      'Trecho de Entrada = l1
      'Query abaixo retornor = l1,l2,l4,l5
        a = "INITIALCOMPONENT"
b = "FINALCOMPONENT"
c = "INITIALGROUNDHEIGHT"
d = "FINALGROUNDHEIGHT"
e = "OBJECT_ID_"
f = "WATERLINES"
g = "WATERCOMPONENTS"
h = "CALCULENODE"
i = "ID_TYPE"
   If frmCanvas.TipoConexao <> 4 Then
      Set rs = Conn.execute("SELECT object_id_, InitialGroundHeight,FinalGroundHeight from waterlines " & _
               "where (InitialComponent in (SELECT object_id_ from watercomponents where object_id_ in(" & MyComponents(0) & "," & MyComponents(1) & ")) " & _
               "Or FinalComponent in (SELECT object_id_ from watercomponents where object_id_ in(" & MyComponents(0) & "," & MyComponents(1) & "))) " & _
               " and id_type = " & mtype)
      
      Else
           Set rs = Conn.execute("SELECT " + """" + e + """" + "," + """" + c + """" + "," + """" + d + """" + " from " + """" + f + """" + _
               "where (" + """" + a + """" + " in (SELECT " + """" + e + """" + " from " + """" + g + """" + " where " + """" + e + """" + " in('" & MyComponents(0) & "','" & MyComponents(1) & "')) " & _
               "Or " + """" + b + """" + " in (SELECT " + """" + e + """" + " from " + """" + g + """" + " where " + """" + e + """" + " in('" & MyComponents(0) & "','" & MyComponents(1) & "'))) " & _
               " and " + """" + i + """" + " = '" & mtype & "'")
      
      End If
      
      
      While Not rs.EOF
         ReDim Preserve MyLocalLines(iLn)
         MyLocalLines(iLn) = rs!Object_id_
         iLn = iLn + 1
         rs.MoveNext
      Wend
      iLn = iLn - 1
      rs.Close
      Set rs = Nothing
      For iPnt = 0 To iLn
         GetSubPrimaria MyLocalLines(iPnt), RSNos, mtype
      Next
      DoEvents
   
   End If
   If Not (rs Is Nothing) Then
      If rs.State = adStateOpen Then rs.Close
   End If
   Set rs = Nothing
End Function


Public Sub GetSubPrimaria(Object_id_ As String, rs As ADODB.Recordset, mtype As Long)
   'Registra e chama quem sera o proximo
   Dim a As Integer, naoChamar As Boolean
   If Not p Then
      ReDim Preserve MyLines(i)
      p = True
   End If
   For a = 0 To i
     If MyLines(a) = Object_id_ Then naoChamar = True
   Next
   If Not naoChamar Then
     FrmProcess.txtRecord = i
     FrmProcess.Refresh
     MyLines(i) = Object_id_
     i = i + 1
     ReDim Preserve MyLines(i)
     rs.Filter = "lin=" & Object_id_
     GetLinesValues Object_id_, rs, mtype
     
   End If

End Sub


Private Function GetSecundComponent(Node_Object_id_ As String, RsTmp As ADODB.Recordset, mtype As Long) As String
   'Registra e chama quem sera o proximo
   Dim RsTrechos As New ADODB.Recordset, a As Integer, naoChamar As Boolean
   Dim rsNodes As New ADODB.Recordset
   
 a = "INITIALCOMPONENT"
b = "FINALCOMPONENT"
c = "INITIALGROUNDHEIGHT"
d = "FINALGROUNDHEIGHT"
e = "OBJECT_ID_"
f = "WATERLINES"
g = "WATERCOMPONENTS"
h = "CALCULENODE"
i = "iD_TYPE"
   If frmCanvas.TipoConexao <> 4 Then
   'retorna os trechos ligados ao no que não de calculo
   RsTrechos.Open "SELECT object_id_, InitialComponent,FinalComponent from waterlines where (FinalComponent=" & Node_Object_id_ & " or  InitialComponent= " & Node_Object_id_ & ") and id_type =" & mtype, Conn
   Else
    RsTrechos.Open "SELECT " + """" + e + """" + ", " + """" + a + """" + "," + """" + b + """" + " from " + """" + f + """" + " where (" + """" + b + """" + "='" & Node_Object_id_ & "' or  " + """" + a + """" + "= '" & Node_Object_id_ & "') and " + """" + i + """" + " ='" & mtype & "', Conn"
   End If
   While Not RsTrechos.EOF
   
      For a = 0 To i
        If MyLines(a) = RsTrechos!Object_id_ Then naoChamar = True
      Next
      If Not naoChamar Then 'vai chamar=false
        MyLines(i) = RsTrechos!Object_id_
        i = i + 1
        ReDim Preserve MyLines(i)
        
        
        If RsTrechos!INITIALCOMPONENT = Node_Object_id_ Then
        If frmCanvas.TipoConexao <> 4 Then
           rsNodes.Open "SELECT * from watercomponents where object_id_='" & RsTrechos!FinalComponent & "'", Conn
           Else
           rsNodes.Open "SELECT * from " + """" + g + """" + " where " + """" + e + """" + "='" & RsTrechos!FinalComponent & "'", Conn
           End If
           If rsNodes!CALCULENODE = 0 Then
              GetSecundComponent RsTrechos!FinalComponent, RsTmp, mtype
           Else
              GetSecundComponent = RsTrechos!FinalComponent
           End If
        Else
        If frmCanvas.TipoConexao <> 4 Then
           rsNodes.Open "SELECT * from watercomponents where object_id_='" & RsTrechos!INITIALCOMPONENT & "'", Conn
           Else
           rsNodes.Open "SELECT * from " + """" + g + """" + " where " + """" + e + """" + "='" & RsTrechos!INITIALCOMPONENT & "'", Conn
           End If
           If rsNodes!CALCULENODE = 0 Then
              GetSecundComponent RsTrechos!INITIALCOMPONENT, RsTmp, mtype
           Else
              GetSecundComponent = RsTrechos!INITIALCOMPONENT
           End If
        End If
      End If
      naoChamar = False
      RsTrechos.MoveNext
   Wend
   RsTrechos.Close
   If rsNodes.State = 1 Then
      rsNodes.Close
   End If
   Set RsTrechos = Nothing
   Set rsNodes = Nothing
End Function


Public Sub DeleteTXTVisual(layer_id As Integer, Object_id_ As String)
   On Error GoTo DeleteTXTVisual_Error
   Dim SQLComand As String
a = "texts"
b = layer_id
c = "b"
d = "_txvisual"
f = "geom_id"
g = layer_id
h = "g"
ii = "object_id"
   If frmCanvas.TipoConexao <> 4 Then
   SQLComand = "Delete From texts" & layer_id & "_txvisual where geom_id in (SELECT geom_id from texts" & layer_id & " where object_id = " & Object_id_ & ")"
   Else
   SQLComand = "Delete From " + """" + a + b + d + """" + " where " + """" + f + """" + " in (SELECT " + """" + f + """" + " from " + """" + a + g + """" + " where " + """" + ii + """" + " = '" & Object_id_ & "')"
   End If
   'MsgBox SQLComand
   'Conn.execute ("Delete From texts" & layer_id & "_txvisual where geom_id in (SELECT geom_id from texts" & layer_id & " where object_id = " & Object_id_)
   Conn.execute (SQLComand)
   Exit Sub
DeleteTXTVisual_Error:
   MsgBox Err.Number & Err.Description
   
   Resume Next
End Sub

Public Function IsValidLayerOperation(LayerName As String, T_Reference As LayerTypeRefence) As Boolean
   Dim rs As ADODB.Recordset, layer_id As Integer
   layer_id = GetLayerID(LayerName)
   g = "GS_LAYER_CONFIG"
h = "LAYER_ID"
ii = "TYPE_REFERENCE"


      If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT * from gs_layer_config where layer_id=" & layer_id & " and type_reference=" & T_Reference)
   Else
     Set rs = Conn.execute("SELECT * from " + """" + g + """" + " where " + """" + h + """" + "='" & layer_id & "' and " + """" + ii + """" + "='" & T_Reference & "'")
   
   End If
   If rs.EOF Then
      IsValidLayerOperation = False
   Else
      IsValidLayerOperation = True
   End If
   rs.Close
   Set rs = Nothing
End Function
' Retorna o tipo de layer de referência
'
' LayerName - nome do layer
'
Public Function GetLayerTypeReference(LayerName As String) As LayerTypeRefence
    On Error GoTo Trata_Erro
    
    Dim rs As ADODB.Recordset, layer_id As Integer
    
    layer_id = GetLayerID(LayerName)                            'obtem o id do layer
    Set rs = Conn.execute("SELECT * from gs_layer_config where layer_id=" & layer_id)
    If rs.EOF Then
        If LayerName = "NXGS_V_LIG_COMERCIAL_GPS" Then          'este if foi adicionado para considerar a leitura dos pontos GPS, no endSelect de um ponto, mas esta função endSelect deve ser revisada
            GetLayerTypeReference = CONSUMIDOR_GPS              'informa que o layer de referência é um ponto GPS
        Else
            GetLayerTypeReference = 0
        End If
    Else
       GetLayerTypeReference = rs.Fields("type_reference").value
    End If
    rs.Close
    Set rs = Nothing
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsGeoReference", "Public Function GetLayerTypeReference", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

Public Function GetLayerNameByTypeReference(Refence As LayerTypeRefence) As String
   Dim rs As ADODB.Recordset
  g = "GS_LAYER_CONFIG"
h = "LAYER_ID"
ii = "TYPE_REFERENCE"
   
If frmCanvas.TipoConexao <> 4 Then

   Set rs = Conn.execute("SELECT layer_id from gs_layer_config where type_reference=" & Refence)
   Else
   Set rs = Conn.execute("SELECT " + """" + h + """" + " from " + """" + g + """" + " where " + """" + ii + """" + "='" & Refence & "'")
   End If
   If Not rs.EOF Then
      GetLayerNameByTypeReference = GetLayerNameLayerID(rs.Fields(0).value)
   End If
   rs.Close
   Set rs = Nothing
End Function
' Esta função tem como objetivo obter o layer de referência de uma determinad operação. Ele está informado na tabela S_LAYER_CONFIG_LAYERS
' Esta função simplesmente vai nesta tabela e obtem o nome do layer de referência a partir do layer slecionado pelo usuário e o tipo de operação que o usuário deseja realizar
'
' LayerName - recebe o nome do layer que foi selecionado pelo usuário
' T_Operation - tipo de operação (ver tabela com os tipos de operações GS_LAYER_CONFIG_LAYERS)
'
Public Function GetLayerOperation(LayerName As String, T_Operation As Integer) As String
    Dim rs As ADODB.Recordset, layer_id As Integer
    Dim s As String
    Dim mensagemInformativa As String                   'para informar o que aconteceu simplesmente. Não foi um erro e sim o usuário não selecionou o plano correto.
    
    k = "GS_LAYER_CONFIG_LAYERS"
    h = "LAYER_ID"
    l = "TYPE_OPERATION"
    layer_id = GetLayerID(LayerName)                    'RETORNA O NUMERO ID DE UM DETERMINADO LAYER
    'layer_id = 3
    If frmCanvas.TipoConexao <> 4 Then
        s = "SELECT * from gs_layer_config_layers where layer_id='" & layer_id & "' and type_operation='" & T_Operation & "'"
    Else
        s = "SELECT * from " + """" + k + """" + " where " + """" + h + """" + "='" & layer_id & "' and " + """" + l + """" + "='" & T_Operation & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.Open s, Conn, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        GetLayerOperation = GetLayerNameLayerID(rs("LAYER_REFERENCE").value)
    Else
        'MsgBox "Não foi localizada referencia para para o layer " & LayerName & ", ID = " & layer_id & " na tabela gs_layer_config_layers com o tipo de operação = " & T_Operation, vbExclamation, "" ---- retirado e colocado como mensagem para a NEXUS
        mensagemInformativa = "Mensagem informativa. Usuário não selecionou o plano correto antes de desenhar a rede. Não foi localizada referencia para para o layer " & LayerName & ", ID = " & layer_id & " na tabela gs_layer_config_layers com o tipo de operação = " & T_Operation
        ErroUsuario.Registra "clsGeoReference", "GetLayerOperation", CStr(Err.Number), CStr(Err.Description), False, glo.enviaEmails, mensagemInformativa   'não apresenta nada para o usuário, pois o usuário apenas não selecionou o plano correto antes de desenhar
    End If
    rs.Close
    Set rs = Nothing
End Function
' Função para localizar o nome de um layer na tabela te_layer, tendo o número do mesmo
' Retorna nome do layer a partir do número (id) do layer
'
' layer_id - número (id) do layer a er procurado
'
Public Function GetLayerNameLayerID(layer_id As Integer) As String
    Dim rs As ADODB.Recordset
    Dim s As String
    Dim str As String
    
    Set rs = New ADODB.Recordset
    k = "te_layer"
    h = "layer_id"
    a = "name"
    If frmCanvas.TipoConexao <> 4 Then
        str = "SELECT name from Te_Layer where layer_id=" & layer_id
    Else
        str = "SELECT " + """" + a + """" + " from " + """" + k + """" + " where " + """" + h + """" + "='" & layer_id & "'"
    End If
    Set rs = Conn.execute(str)
    If rs.EOF = False Then
        GetLayerNameLayerID = rs(0).value
    Else
        MsgBox "Falha em: SELECT name from te_Layer where layer_id= " & layer_id & " - Function GetLayerNameLayerID. Tabela GS_LAYER_CONFIG_LAYERS não possui o LAYER_REFERENCE para encontrar os ramais em um lote da Prefeitura (LOTEGIS-não mais utilizada)"
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Sub AddAtributesLinesOut(LayerName As String)
   Dim rs As ADODB.Recordset, str As String, Point1 As String, Point2 As String, compr As Double, a As Integer, X As Double, Y As Double
   If frmCanvas.TipoConexao <> 4 Then
   str = "SELECT object_id from lines38 left join waterlines on object_id=object_id_ where object_id_ is null"
   Else
   a = "INITIALCOMPONENT"
b = "FINALCOMPONENT"
c = "INITIALGROUNDHEIGHT"
d = "FINALGROUNDHEIGHT"
e = "OBJECT_ID_"
f = "WATERLINES"
g = "WATERCOMPONENTS"
h = "lines38"
i = "object_id"
     str = "SELECT " + """" + i + """" + " from " + """" + h + """" + " left join " + """" + f + """" + " on " + """" + i + """" + "=" + """" + e + """" + " where " + """" + e + """" + " is null"
   End If
   Set rs = Conn.execute(str)
   While Not rs.EOF
      compr = 0
      tdb.getLengthOfLine 0, rs(0).value, compr
      For a = 0 To 1
         tdb.getPointOfLine 0, rs(0).value, a, X, Y
         If tdbcon.locateGeometry(X, Y, tpPOINTS, 1) Then
             If a = 0 Then
               Point1 = tdbcon.objectIds(0)
             Else
               Point2 = tdbcon.objectIds(0)
               
                   a = "OBJECT_ID_"
      b = "INITIALCOMPONENT"
      c = "FINALCOMPONENT"
      d = "LENGHTCALCULATED"
      


     If frmCanvas.TipoConexao <> 4 Then
         
     str = "insert into " & LayerName & _
                     " (object_id_,InitialComponent,FinalComponent, lenghtcalculated) values ('" & _
                     rs.Fields("object_id").value & "','" & Point1 & "','" & Point2 & "'," & Replace(compr, ",", ".") & ")"
     
     Else
     
      str = "insert into " + """" + LayerName + """" + " (" + """" + a + """" + "," + """" + b + """" + "," + """" + c + "," + """" + d + """" + ") values ('" & _
                     rs.Fields("object_id").value & "','" & Point1 & "','" & Point2 & "','" & Replace(compr, ",", ".") & "')"
     End If
               
               
               
            
               Conn.execute str
             End If
         End If
      Next
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   
a = "lines38"
b = "OBJECT_ID_"
c = "line_id"
d = "WATERLINES"
e = "object_id"
f = "LINE_ID"
    If frmCanvas.TipoConexao <> 4 Then
   Conn.execute "update lines38 set object_id = a.line_id " & _
                "from waterlines a inner join lines38 b on b.object_id = a.object_id_ " & _
                "Where a.line_id <> a.Object_id_"
   Conn.execute "update waterlines set object_id_ = line_id where line_id <> object_id_"
   Else
   Conn.execute "update " + """" + a + """" + " set " + """" + b + """" + " = " + """" + d + """" + "." + """" + f + """" + _
                "from " + """" + d + """" + " inner join " + """" + a + """" + "on " + """" + a + """" + "." + """" + e + """" + " = " + """" + d + """" + "." + """" + b + """" + _
                "Where " + """" + d + """" + "." + """" + f + """" + " <> " + """" + d + """" + "." + """" + b + """" + ""
   Conn.execute "update " + """" + d + """" + " set " + """" + b + """" + " = " + """" + c + """" + " where " + """" + c + """" + " <> " + """" + b + """" + ""
   
   End If
End Sub

'Public Sub Update_Demand_Node()
'   'On Error GoTo Update_Demand_Node_err '
'   Dim IniciouTransacao As Boolean
'   Dim rsLines As ADODB.Recordset, str As String
'   Dim rsNodes As ADODB.Recordset
'   Dim RsSpecfication As New ADODB.Recordset
'   Dim rsTrecho As New ADODB.Recordset, Demanda As Double
'   Dim NOI As String, UltimoNoDemand As String
'   Dim frm As New frmProgressBar, UltimoTrecho As String
'   Dim RSpARAMETER As New ADODB.Recordset
'   RSpARAMETER.Open "SELECT * from gs_parameters", Conn, adOpenKeyset, adLockOptimistic
'
'
'
'   frm.Show , FrmMain
'   Screen.MousePointer = vbHourglass
'
'   str = GetQueryProcess(22)
'   If RSpARAMETER!demand_update_line_id > 0 Then
'      str = Replace(str, "-- and w.line_id > @lines_id", " and w.line_id >" & RSpARAMETER!demand_update_line_id)
'   Else
'      str = Replace(str, "-- and w.line_id > @lines_id", "")
'      Conn.execute "UPDATE WATERCOMPONENTS SET DEMAND = 0"
'   End If
'
'   Set rsNodes = Conn.execute(str)
'   frm.ProgressBar1.Max = rsNodes(0).value
'   frm.ProgressBar1.Min = 0
'   rsNodes.Close
'   Set rsNodes = Nothing
'
'   str = GetQueryProcess(21)
'   If RSpARAMETER!demand_update_line_id > 0 Then
'      str = Replace(str, "-- and w.line_id > @lines_id", " and w.line_id >" & RSpARAMETER!demand_update_line_id)
'   Else
'      str = Replace(str, "-- and w.line_id > @lines_id", "")
'   End If
'   Set rsLines = ConnSec.execute(str)
'   Set rsNodes = New ADODB.Recordset
'   rsNodes.CursorType = adOpenDynamic
'   rsNodes.LockType = adLockOptimistic
'
'   str = "SELECT OBJECT_ID_, demand from watercomponents "
'   Set rsNodes = Conn.execute(str)
'
'   With rsLines
'
'      While Not .EOF
'         str = "object_id_='" & .Fields("initialcomponent").value & "' OR "
'         str = str & "object_id_='" & .Fields("finalcomponent").value & "'"
'         rsNodes.Requery
'         rsNodes.Filter = str
'         Conn.BeginTrans
'         IniciouTransacao = True
'         UltimoTrecho = .Fields("OBJECT_ID_").value
'         While Not rsNodes.EOF
'            str = "SELECT OBJECT_ID_ from watercomponents c inner join watercomponentsTypes w on w.id_type=c.id_type "
'            str = str & "where w.specification_ ='JUNCTION' AND C.OBJECT_ID_='" & rsNodes.Fields("OBJECT_ID_").value & "'"
'            RsSpecfication.Open str, Conn
'            If Not RsSpecfication.EOF Then
'               Demanda = Round(IIf(IsNull(.Fields("CONSUMO").value), 0, .Fields("CONSUMO").value) / 2, 2)
'               If Demanda > 0 Then
'                  Conn.execute "update watercomponents set demand= demand + " & Replace(Demanda, ",", ".") & _
'                               " where object_id_='" & rsNodes.Fields("object_id_").value & "'"
'               End If
'               'rsNodes.Fields("DEMAND").value = rsNodes.Fields("DEMAND").value + Round(.Fields("CONSUMO").value / 2, 2)
'               UltimoNoDemand = rsNodes.Fields("object_id_").value
'               rsNodes.MoveNext
'            Else 'O NO É UMA VALVULA OU BOMBA OU RESERVATORIO
'               rsTrecho.Open "SELECT * from waterlines " & _
'                  "where (initialcomponent = '" & rsNodes.Fields("OBJECT_ID_").value & "' or finalcomponent = '" & rsNodes.Fields("OBJECT_ID_").value & "')" & _
'                  " AND OBJECT_ID_ <> '" & UltimoTrecho & "'", Conn
'               If rsTrecho.EOF Then
'                  str = "object_id_='" & UltimoNoDemand & "'"
'                  rsNodes.Requery
'                  rsNodes.Filter = str
'               Else
'                  If CStr(rsTrecho.Fields("initialcomponent")) <> rsNodes.Fields("OBJECT_ID_").value Then
'                     NOI = rsTrecho.Fields("initialcomponent")
'                  ElseIf CStr(rsTrecho.Fields("finalcomponent")) <> rsNodes.Fields("OBJECT_ID_").value Then
'                     NOI = rsTrecho.Fields("finalcomponent")
'                  End If
'                  UltimoTrecho = rsTrecho.Fields("OBJECT_ID_").value
'                  str = "object_id_='" & NOI & "'"
'                  rsNodes.Requery
'                  rsNodes.Filter = str
'               End If
'               rsTrecho.Close
'            End If
'            RsSpecfication.Close
'         Wend
'         frm.ProgressBar1.Min = frm.ProgressBar1.Min + 1
'         frm.Caption = frm.ProgressBar1.Min & "/" & frm.ProgressBar1.Max
'         DoEvents
'         RSpARAMETER!demand_update_line_id = .Fields("object_id_").value
'         RSpARAMETER!DEMAND_LAST_UPDATE_LINE_ID = .Fields("object_id_").value
'         RSpARAMETER.Update
'         Conn.CommitTrans
'         IniciouTransacao = False
'         .MoveNext
'      Wend
'   End With
'   Unload frm
'   RSpARAMETER!demand_update_line_id = 0
'   RSpARAMETER.Update
'   Set frm = Nothing
'   Screen.MousePointer = vbNormal
'   Set RsSpecfication = Nothing
'   Set rsTrecho = Nothing
'   'rsNodes.Update
'   rsNodes.Close
'   Set rsNodes = Nothing
'   rsLines.Close
'   Set rsLines = Nothing
'   RSpARAMETER.Close
'   Set RSpARAMETER = Nothing
'   Exit Sub
'Update_Demand_Node_err:
'   If IniciouTransacao Then Conn.RollbackTrans
'   Screen.MousePointer = vbNormal
'   If Not RsSpecfication Is Nothing Then
'      If RsSpecfication.State = 1 Then RsSpecfication.Close
'      Set RsSpecfication = Nothing
'   End If
'   If Not rsTrecho Is Nothing Then
'      If rsTrecho.State = 1 Then rsTrecho.Close
'      Set rsTrecho = Nothing
'   End If
'   If Not rsNodes Is Nothing Then
'      If rsNodes.State = 1 Then rsNodes.Close
'      Set rsNodes = Nothing
'   End If
'   If Not rsLines Is Nothing Then
'      If rsLines.State = 1 Then rsLines.Close
'      Set rsLines = Nothing
'   End If
'   If Not RSpARAMETER Is Nothing Then
'      If RSpARAMETER.State = 1 Then RSpARAMETER.Close
'      Set RSpARAMETER = Nothing
'   End If
'
'   MsgBox Err.Description, vbExclamation
'
'End Sub

Public Sub Update_Demand_Node()
   On Error GoTo Update_Demand_Node_err '
   Dim IniciouTransacao As Boolean
   Dim rsLines As ADODB.Recordset, str As String
   Dim rsNodes As ADODB.Recordset
   Dim RsSpecfication As New ADODB.Recordset
   Dim rsTrecho As New ADODB.Recordset, Demanda As Double
   Dim NOI As String, UltimoNoDemand As String
   Dim frm As New frmProgressBar, UltimoTrecho As String
   Dim RSpARAMETER As New ADODB.Recordset
   
'   If frmCanvas.TipoConexao = 1 Then 'SQL
'      'MsgBox "Conexão SQL"
'      UPDATE_DEMANDA_BY_JONATHAS
'      Exit Sub
'   End If
   
   
   
      a = "gs_parameters"
      b = "INITIALCOMPONENT"
      c = "FINALCOMPONENT"
      d = "LENGHTCALCULATED"

   If frmCanvas.TipoConexao <> 4 Then
   RSpARAMETER.Open "SELECT * from gs_parameters", Conn, adOpenKeyset, adLockOptimistic
   Else
   RSpARAMETER.Open "SELECT * from " + """" + a + """", Conn, adOpenKeyset, adLockOptimistic
   End If
   
   str = ""
   str = GetQueryProcess(22)
   
   If Trim(str) = "" Then Exit Sub
   
   frm.Show , FrmMain
   Screen.MousePointer = vbHourglass
   
   If RSpARAMETER!demand_update_line_id > 0 Then
      str = Replace(str, "-- and w.line_id > @lines_id", " and w.line_id >" & RSpARAMETER!demand_update_line_id)
   Else
      str = Replace(str, "-- and w.line_id > @lines_id", "")
      
a = "WATERCOMPONENTS"
b = "DEMAND"
  If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "UPDATE WATERCOMPONENTS SET DEMAND = 0"
      Else
      Conn.execute "UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = '0'"
      End If
   End If
   
   Set rsNodes = Conn.execute(str)
   frm.ProgressBar1.Max = rsNodes(0).value
   frm.ProgressBar1.Min = 0
   rsNodes.Close
   Set rsNodes = Nothing
   
   str = GetQueryProcess(21)
   If RSpARAMETER!demand_update_line_id > 0 Then
      str = Replace(str, "-- and w.line_id > @lines_id", " and w.line_id >" & RSpARAMETER!demand_update_line_id)
   Else
      str = Replace(str, "-- and w.line_id > @lines_id", "")
   End If
   Set rsLines = ConnSec.execute(str)
   Set rsNodes = New ADODB.Recordset
   rsNodes.CursorType = adOpenDynamic
   rsNodes.CursorLocation = adUseClient
   
   
   With rsLines
      
      While Not .EOF
         Conn.BeginTrans
         IniciouTransacao = True
         UltimoTrecho = .Fields("OBJECT_ID_").value
          If frmCanvas.TipoConexao <> 4 Then
         str = "SELECT OBJECT_ID_, w.specification_ from watercomponents c left join watercomponentsTypes w on w.id_type=c.id_type "
         str = str & "where C.OBJECT_ID_='" & .Fields("initialcomponent").value & "' or  C.OBJECT_ID_='" & .Fields("finalcomponent").value & "'"
         Else
         b = "WATERCOMPONENTS"
c = "DEMAND"
d = "TIPO"
e = Replace(Demanda, ",", ".")
f = "ESPECIFICATION_"
g = "WATERCOMPONENTSTYPES"
h = "'g'"
'i = DEMAND
j = "iD_TYPE"
k = "OBJECT_ID_"

str = "SELECT " + """" + k + """" + ", " + """" + g + """" + "." + """" + e + """" + " from " + """" + b + """" + "  left join " + """" + g + """" + "  on " + """" + g + """" + "." + """" + j + """" + "=" + """" + b + """" + "." + """" + j + """" + ""
         str = str & "where " + """" + b + """" + "." + """" + k + """" + "='" & .Fields("INITIALCOMPONENT").value & "' or  " + """" + b + """" + " ." + """" + k + """" + " ='" & .Fields("FINALCOMPONENT").value & "'"
         End If
         rsNodes.Open str, Conn
         While Not rsNodes.EOF
            
            Demanda = Round(IIf(IsNull(.Fields("CONSUMO").value), 0, .Fields("CONSUMO").value) / 2, 2)
            
            If rsNodes.Fields("specification_").value = "JUNCTION" Then
               If Demanda > 0 Then
               a = "OBJECT_ID_"
b = "WATERCOMPONENTS"
c = "DEMAND"
d = "TIPO"
e = Replace(Demanda, ",", ".")
f = "e'"
g = rsNodes.Fields("object_id_").value
h = "'g'"
'i = DEMAND
j = "'i"
k = "OBEJECT_ID"


                  If frmCanvas.TipoConexao <> 4 Then
                  Conn.execute "update watercomponents set demand= demand + " & Replace(Demanda, ",", ".") & _
                               " where object_id_='" & rsNodes.Fields("object_id_").value & "'"
                Else
                 Conn.execute "update " + """" + b + """" + " set " + """" + c + """" + " = '" + c + e + "' where " + """" + k + """" + "='" + g + "'"
                End If
               End If
               UltimoNoDemand = rsNodes.Fields("object_id_").value
               
            Else 'O NO É UMA VALVULA OU BOMBA OU RESERVATORIO
               Sub_Update_Demand_Next_Node rsNodes.Fields("OBJECT_ID_").value, UltimoNoDemand, UltimoTrecho, Demanda
            End If
            rsNodes.MoveNext
         Wend
         rsNodes.Close
         frm.ProgressBar1.value = frm.ProgressBar1.value + 1
         frm.Caption = frm.ProgressBar1.Min & "/" & frm.ProgressBar1.Max
         DoEvents
         RSpARAMETER!demand_update_line_id = .Fields("object_id_").value
         RSpARAMETER!DEMAND_LAST_UPDATE_LINE_ID = .Fields("object_id_").value
         RSpARAMETER.Update
         Conn.CommitTrans
         IniciouTransacao = False
         .MoveNext
      Wend
   End With
   Unload frm
   RSpARAMETER!demand_update_line_id = 0
   RSpARAMETER.Update
   Set frm = Nothing
   Screen.MousePointer = vbNormal
   Set RsSpecfication = Nothing
   Set rsTrecho = Nothing
   Set rsNodes = Nothing
   rsLines.Close
   Set rsLines = Nothing
   RSpARAMETER.Close
   Set RSpARAMETER = Nothing
   Exit Sub
Update_Demand_Node_err:
   If IniciouTransacao Then Conn.RollbackTrans
   Screen.MousePointer = vbNormal
   If Not RsSpecfication Is Nothing Then
      If RsSpecfication.State = 1 Then RsSpecfication.Close
      Set RsSpecfication = Nothing
   End If
   If Not rsTrecho Is Nothing Then
      If rsTrecho.State = 1 Then rsTrecho.Close
      Set rsTrecho = Nothing
   End If
   If Not rsNodes Is Nothing Then
      If rsNodes.State = 1 Then rsNodes.Close
      Set rsNodes = Nothing
   End If
   If Not rsLines Is Nothing Then
      If rsLines.State = 1 Then rsLines.Close
      Set rsLines = Nothing
   End If
   If Not RSpARAMETER Is Nothing Then
      If RSpARAMETER.State = 1 Then RSpARAMETER.Close
      Set RSpARAMETER = Nothing
   End If
    MsgBox Err.Description, vbExclamation
    If mid(Err.Description, 1, 9) = "ORA-03114" Then
        Conn.RollbackTrans
        MsgBox "Não há conexão ativa com o banco de dados", vbExclamation, "Conexão desativada"
        End
    End If
   'Resume

End Sub

Sub Sub_Update_Demand_Next_Node(No As String, UltimoNo As String, Trecho As String, Demanda As Double)
   Dim rsTrecho As New ADODB.Recordset
   Dim RsSpecfication As ADODB.Recordset
   Dim strsql  As String
   If frmCanvas.TipoConexao <> 4 Then
   rsTrecho.Open "SELECT * from waterlines " & _
      "where (initialcomponent = '" & No & "' or finalcomponent = '" & No & "')" & _
      " AND OBJECT_ID_ <> '" & Trecho & "'", Conn
      Else
      a = "WATERCOMPONENTS"
b = "DEMAND"
c = Replace(Demanda, ",", ".")
d = "WATERLINES"
e = "FINALCOMPONENT"
f = "INITIALCOMPONENT"
g = "OBJECT_ID_"

      rsTrecho.Open "SELECT * from " + """" + d + """" + " " & _
      "where (" + """" + f + """" + " = '" & No & "' or " + """" + e + """" + " = '" & No & "')" & _
      " AND " + """" + g + """" + " <> '" & Trecho & "'", Conn
      End If
      
      
   If rsTrecho.EOF Then

 If frmCanvas.TipoConexao <> 4 Then
 
      Conn.execute "update watercomponents set demand= demand + " & Replace(Demanda, ",", ".") & _
             " where object_id_='" & UltimoNo & "'"
             Else
             a = "WATERCOMPONENTS"
b = "DEMAND"
c = Replace(Demanda, ",", ".")
d = "WATERLINES"
'e = DEMAND
f = "'e'"
g = "OBJECT_ID_"
             Conn.execute "update " + """" + a + """" + "set " + """" + b + """" + "='b + " & Replace(Demanda, ",", ".") ' & "where " +""""+ g +""""+ "='" & UltimoNo & "'"
             End If
             
             
      rsTrecho.Close
      Set rsTrecho = Nothing
   Else
      If CStr(rsTrecho.Fields("INITIALCOMPONENT")) <> No Then
         No = rsTrecho.Fields("INITIALCOMPONENT")
      ElseIf CStr(rsTrecho.Fields("FINALCOMPONENT")) <> No Then
         No = rsTrecho.Fields("FINALCOMPONENT")
      End If
a = "WATERCOMPONENTS"
b = "DEMAND"
c = Replace(Demanda, ",", ".")
d = "'c'"
'e = DEMAND
f = "'e'"
g = "OBJECT_ID_"
h = "WATERCOMPONENTSTYPES"
i = "ID_TYPE"

k = Replace(Demanda, ",", ".")
l = "'k'"
Dim gu As String
gu = "SPECIFICATION_"
      If frmCanvas.TipoConexao <> 4 Then
      strsql = "SELECT OBJECT_ID_ from watercomponents c inner join watercomponentsTypes w on w.id_type=c.id_type "
      strsql = strsql & "where w.specification_ ='JUNCTION' AND C.OBJECT_ID_='" & No & "'"
      Else
       strsql = "SELECT " + """" + g + """" + " from " + """" + a + """" + " inner join " + """" + h + """" + " on " + """" + h + """" + "." + """" + i + """" + "=" + """" + a + """" + "." + """" + i + """" + ""
      strsql = strsql + "where " + """" + h + """" + "." + """" + gu + """" + " ='JUNCTION' AND " + """" + a + """" + "." + """" + g + """" + "='" & No & "'"
      End If
      Set RsSpecfication = New ADODB.Recordset
      RsSpecfication.Open strsql, Conn
      If Not RsSpecfication.EOF Then
         If Demanda > 0 Then
         If frmCanvas.TipoConexao <> 4 Then
            Conn.execute "update watercomponents set demand= demand + " & Replace(Demanda, ",", ".") & _
                         " where object_id_='" & No & "'"
                         Else
                           Conn.execute "update " + """" + a + """" + " set " + """" + e + """" + "= " + """" + "demand" + Replace(Demanda, ",", ".") + """" + _
                         " where " + """" + g + """" + "='" & No & "'"
                         End If
         End If
         rsTrecho.Close
         Set rsTrecho = Nothing
         RsSpecfication.Close
         Set RsSpecfication = Nothing
      
      Else 'O NO É UMA VALVULA OU BOMBA OU RESERVATORIO
         Trecho = rsTrecho.Fields("object_id_").value
         rsTrecho.Close
         Set rsTrecho = Nothing
         RsSpecfication.Close
         Set RsSpecfication = Nothing
         
         Sub_Update_Demand_Next_Node No, UltimoNo, Trecho, Demanda
      End If
   End If
End Sub


