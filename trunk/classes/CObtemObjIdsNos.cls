VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObtemObjIdsNos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Obtem os object_ids dos nós de um determinado trehco de rede
'
' retorna verdadeira se a operação foi realizada com sucesso.
' objIdTrecho - do trecho de rede em que ele vai procurar os nós inicial e final
' objIdNo1 - retorna o obj_id do nó inicial
' objIdNo2 - retorna o obj_id do nó final
'
Public Function getObjIdsNosRede(objIdTrecho As Long, ByRef objIdNo1 As String, ByRef objIdNo2 As String) As Boolean
    On Error GoTo Trata_Erro:
    Dim rsRedeLocNos As ADODB.Recordset
    Dim sqlString As String
    
    Set rsRedeLocNos = New ADODB.Recordset
    sqlString = "Select InitialComponent, FinalComponent from WATERLINES where object_id_ = '" & CStr(objIdTrecho) & "'"
    rsRedeLocNos.Open sqlString, Conn, adOpenDynamic, adLockOptimistic
    If rsRedeLocNos.EOF = False Then
        objIdNo1 = CStr(rsRedeLocNos("InitialComponent").value)
        objIdNo2 = CStr(rsRedeLocNos("FinalComponent").value)
    Else
        ErroUsuario.Registra "CObtemObjIdsNos", "getObjIdsNosRede-Não econtrado o trecho de rede: " & objIdTrecho, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
    rsRedeLocNos.Close
    getObjIdsNosRede = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CObtemObjIdsNos", "getObjIdsNosRede. Trecho rede: " & objIdTecho & " Nó inicial " & objIdNo1 & ", nó final " & objIdNo2, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
        getObjIdsNosRede = False
    End If
End Function
