VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEPAPatterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarID As Long
Private mvarDESCRICAO As String
Private mvarPADRAO As String

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

Public Property Let PADRAO(ByVal vData As String)
    mvarPADRAO = vData
End Property

Public Property Get PADRAO() As String
    PADRAO = mvarPADRAO
End Property

Public Function Inserir_Padrao() As Boolean
   Inserir_Padrao = False
   On Error GoTo InsereDado_err
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   
   'rs.Open "x_patterns", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
   rs.Open "x_patterns", Conn, adOpenDynamic, adLockOptimistic
   rs.AddNew
   rs.Fields("ID").value = ID
   rs.Fields("DESCRICAO").value = IIf(DESCRICAO = "", Null, DESCRICAO)
   rs.Fields("PADRAO").value = IIf(PADRAO = "", Null, PADRAO)
   rs.Update
   Inserir_Padrao = True
   rs.Close
   Set rs = Nothing
   Exit Function
InsereDado_err:
   If Err.Number <> 0 Then
      MsgBox "Número da curva já existe", vbExclamation
   End If
End Function

'/// Atualiza a padrao  da instancia para o Banco
'/// ou do banco para instancia.
'/// Atualiza/carrega a instancia com o id
'/// fornecido.
Public Function Atualizar_Padrao(Padrao_id As Long, Optional Banco As Boolean = False) As Boolean
   Atualizar_Padrao = False
   On Error GoTo Atualizar_Padrao_err
   Dim rs As New ADODB.Recordset
If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from x_patterns where id = " & Padrao_id, Conn, adOpenKeyset, adLockOptimistic
   Else
   Dim p As String
    Dim k As String
    p = "X_PATTERNS"
    k = "ID"
    rs.Open "SELECT * from " + """" + p + """" + " where " + """" + k + """" + " = '" & Padrao_id & "'", Conn, adOpenDynamic, adLockOptimistic
   End If
  
   If rs.EOF Then
         MsgBox "Número da curva não cadastrada", vbExclamation
   Else
      If Banco Then
         rs.Fields("DESCRICAO").value = IIf(DESCRICAO = "", Null, DESCRICAO)
         rs.Fields("PADRAO").value = IIf(PADRAO = "", Null, PADRAO)
         rs.Update
         
      Else
         ID = rs.Fields("id").value
         DESCRICAO = rs.Fields("DESCRICAO").value
         PADRAO = rs.Fields("PADRAO").value
      End If
      Atualizar_Padrao = True
   End If
   rs.Close
   Set rs = Nothing
   Exit Function
Atualizar_Padrao_err:
   If Err.Number <> 0 Then
      MsgBox "Número da curva já existe", vbExclamation
   End If
End Function

'/// Retorna um cursor com todas as padroes cadastrados
Public Function Retorna_Padroes(rsCurvas As ADODB.Recordset) As Boolean
   Retorna_Padroes = False
   On Error GoTo Retorna_Padroes_err
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim a As String
   Dim b As String
   a = "X_PATTERNS"
   b = "ID"
   If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from x_patterns order by id", Conn
   Else
    rs.Open "SELECT * from " + """" + a + """" + " order by" + """" + b + """", Conn, adOpenDynamic, adLockOptimistic
   
   End If
   
   If rs.EOF Then
         rs.Close
   Else
      Set rsCurvas = rs
      Retorna_Padroes = True
   End If
   Set rs = Nothing
   Exit Function
Retorna_Padroes_err:
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & "Classe:Padrões EPA - Método: Retorna_Padroes", vbExclamation
   End If
End Function

'/// Exclui um curva existente no database
Public Function Excluir_Padrao(Padrao_id As Long) As Boolean
   Excluir_Padrao = False
   On Error GoTo Delete_Curva_err
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   If frmCanvas.TipoConexao <> 4 Then
   rs.Open "SELECT * from x_patterns where id = " & Padrao_id, Conn, adOpenKeyset, adLockOptimistic
  Else
  Dim p As String
    Dim k As String
    p = "X_PATTERNS"
    k = "ID"
    rs.Open "SELECT * from " + """" + p + """" + " where " + """" + k + """" + " = '" & Padrao_id & "', Conn, adOpenDynamic, adLockOptimistic"
   End If
   
   If rs.EOF Then
         MsgBox "Padrão não cadastrada", vbExclamation
   Else
      rs.Delete
      rs.Update
      rs.Close
      Excluir_Padrao = True
   End If
   Set rs = Nothing
   Exit Function
Delete_Curva_err:
   If Err.Number <> 0 Then
      MsgBox "Padrão não encontrado", vbExclamation
   End If
End Function

