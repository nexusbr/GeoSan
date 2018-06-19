VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEPACurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarID As Long
Private mvarDESCRICAO As String
Private mvarCOORDENADA_X As String
Private mvarCOORDENADA_Y As String
Private mvarTIPO As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let DESCRICAO(ByVal vData As String)
    mvarDESCRICAO = vData
End Property

Public Property Get DESCRICAO() As String
    DESCRICAO = mvarDESCRICAO
End Property

Public Property Let COORDENADA_X(ByVal vData As String)
    mvarCOORDENADA_X = vData
End Property

Public Property Get COORDENADA_X() As String
    COORDENADA_X = mvarCOORDENADA_X
End Property

Public Property Let COORDENADA_Y(ByVal vData As String)
    mvarCOORDENADA_Y = vData
End Property

Public Property Get COORDENADA_Y() As String
    COORDENADA_Y = mvarCOORDENADA_Y
End Property

Public Property Let tipo(ByVal vData As String)
    mvarTIPO = vData
End Property

Public Property Get tipo() As String
    tipo = mvarTIPO
End Property


Public Function Inserir_Curva() As Boolean
   Inserir_Curva = False
   On Error GoTo InsereDado_err
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
  ' rs.Open "x_curves", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
   rs.Open "x_curves", Conn, adOpenDynamic, adLockOptimistic
   rs.AddNew
   rs.Fields("ID").value = ID
   rs.Fields("DESCRICAO").value = IIf(DESCRICAO = "", Null, DESCRICAO)
   rs.Fields("COORDENADA_X").value = IIf(COORDENADA_X = "", Null, COORDENADA_X)
   rs.Fields("COORDENADA_Y").value = IIf(COORDENADA_Y = "", Null, COORDENADA_Y)
   rs.Fields("TIPO").value = IIf(tipo = "", Null, tipo)
   rs.Update
   Inserir_Curva = True
   rs.Close
   Set rs = Nothing
   Exit Function
InsereDado_err:
   If Err.Number <> 0 Then
      MsgBox "Número da curva já existe", vbExclamation
   End If
End Function

'/// Atualiza a Curva selecionada da instancia para o Banco
'/// ou do banco para instancia.
'/// Por Padrão atualiza/carrega a instancia com o id
'/// fornecido.
Public Function Atualizar_Curva(Curva_id As Long, Optional Banco As Boolean = False) As Boolean
   Atualizar_Curva = False
   On Error GoTo Atualizar_Curva_err
   Dim rs As New ADODB.Recordset
     If frmCanvas.TipoConexao <> 4 Then
  ' rs.Open "SELECT * from x_curves where id = " & Curva_id, Conn, adOpenKeyset, adLockOptimistic
   rs.Open "SELECT * from x_curves where id = " & Curva_id, Conn, adOpenDynamic, adLockOptimistic
   Else
   a = "x_curves"
   b = "id"
      rs.Open "SELECT * from " + """" + a + """" + " where " + """" + b + """" + " = '" & Curva_id & "'", Conn, adOpenDynamic, adLockOptimistic
   
   End If
   If rs.EOF Then
         MsgBox "Número da curva não cadastrada", vbExclamation
   Else
      If Banco Then
         rs.Fields("DESCRICAO").value = IIf(DESCRICAO = "", Null, DESCRICAO)
         rs.Fields("COORDENADA_X").value = IIf(COORDENADA_X = "", Null, COORDENADA_X)
         rs.Fields("COORDENADA_Y").value = IIf(COORDENADA_Y = "", Null, COORDENADA_Y)
         rs.Fields("TIPO").value = IIf(tipo = "", Null, tipo)
         rs.Update
         
      Else
         ID = rs.Fields("id").value
         DESCRICAO = rs.Fields("DESCRICAO").value
         COORDENADA_X = rs.Fields("COORDENADA_X").value
         COORDENADA_Y = rs.Fields("COORDENADA_Y").value
         tipo = rs.Fields("TIPO").value
      End If
      Atualizar_Curva = True
   End If
   rs.Close
   Set rs = Nothing
   Exit Function
Atualizar_Curva_err:
   If Err.Number <> 0 Then
      MsgBox "Número da curva já existe", vbExclamation
   End If
End Function

'/// Retorna um cursor com todas as curvas cadastradas
Public Function Retorna_Curvas(rsCurvas As ADODB.Recordset) As Boolean
   Retorna_Curvas = False
   On Error GoTo Retorna_Curvas_err
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   a = "X_CURVES"
   b = "ID"
   
     If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from x_curves order by id", Conn
   Else
   rs.Open "SELECT * from " + """" + a + """" + " order by " + """" + b + """", Conn, adOpenDynamic, adLockOptimistic
   End If
   
   If rs.EOF Then
         rs.Close
   Else
      Set rsCurvas = rs
      Retorna_Curvas = True
   End If
   Set rs = Nothing
   Exit Function
Retorna_Curvas_err:
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & "Classe:Curvas - Método: Retorna_Curvas", vbExclamation
   End If
End Function

'/// Exclui um curva existente no database
Public Function Excluir_Curva(Curva_id As Long) As Boolean
   Excluir_Curva = False
   On Error GoTo Delete_Curva_err
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
    a = "x_CURVES"
   b = "ID"
   
     If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from x_curves where id = " & Curva_id, Conn, adOpenKeyset, adLockOptimistic
   Else
    rs.Open "SELECT * from " + """" + a + """" + " where " + """" + b + """" + " = '" & Curva_id & "'", Conn, adOpenDynamic, adLockOptimistic
   End If
   
   
   If rs.EOF Then
         MsgBox "Número da curva não cadastrada", vbExclamation
   Else
      rs.Delete
      rs.Update
      rs.Close
      Excluir_Curva = True
   End If
   Set rs = Nothing
   Exit Function
Delete_Curva_err:
   If Err.Number <> 0 Then
      MsgBox "Número da curva não encontrado", vbExclamation
   End If
End Function

