VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLocalizaTrechos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para localizar os trechos de rede que estão conectados a um determinado nó

' Recebe o objID do nó de rede e o ponteiro para a lista de ObjIds dos trechos de rede conectados ao nó enviado
'
' objIdNoSelecionado - objID do nó de rede (ponto), ao qual serão verificados os trechos conectados ao mesmo
' objIDs - retorno da lista de objIDs dos trechos de rede que estão ligados ao nó de rede fornecido
'
Public Sub obtemTrechos(objIdNoSelecionado, ByRef objIDs() As String)
    On Error GoTo Trata_Erro:
    Dim rsTrechosConectadosAoNo As ADODB.Recordset
    Dim i As Integer
    
    i = 0
    Set rsTrechosConectadosAoNo = New ADODB.Recordset
    rsTrechosConectadosAoNo.Open "Select object_id_ From WaterLines WHERE InitialComponent = '" & objIdNoSelecionado & "' OR FinalComponent = '" & objIdNoSelecionado & "'", Conn, adOpenDynamic, adLockOptimistic
    Do While rsTrechosConectadosAoNo.EOF = False
        ReDim Preserve objIDs(i) As String
        objIDs(i) = rsTrechosConectadosAoNo("object_id_").value
        i = i + 1
        rsTrechosConectadosAoNo.MoveNext
    Loop
    rsTrechosConectadosAoNo.Close
    Exit Sub
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CLocalizaTrechos", "obtemTrechos", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Sub
