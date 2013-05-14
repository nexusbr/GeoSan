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
' Classe para localizar os trechos de rede que est�o conectados a um determinado n�

' Recebe o objID do n� de rede e o ponteiro para a lista de ObjIds dos trechos de rede conectados ao n� enviado
'
' objIdNoSelecionado - objID do n� de rede (ponto), ao qual ser�o verificados os trechos conectados ao mesmo
' objIDs - retorno da lista de objIDs dos trechos de rede que est�o ligados ao n� de rede fornecido
'
Public Sub obtemTrechos(objIdNoSelecionado, ByRef objIDs() As String)
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
End Sub