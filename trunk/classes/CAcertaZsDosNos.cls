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
    

    Set RSNos = New ADODB.Recordset
    RSNos.Open "Select x, y, object_id from Points2", Conn, adOpenKeyset, adLockOptimistic
    Open "d:\temp.txt" For Output As #10                                                                            'Abre o arquivo novo
    Close #10
    retorno = cGeoDatabase.geoDatabase.setCurrentLayer("mdt")                                                       'ativa o layer do MDT
    Do While RSNos.EOF = False
        xNo = RSNos("x").value
        yNo = RSNos("y").value
        objIdNo = RSNos("object_id").value
        zNo = getZNo.ObtemZ(xNo, yNo)
        RSNos.MoveNext
        Set rsAtributoNos = New ADODB.Recordset
        rsAtributoNos.Open "SELECT GROUNDHEIGHT FROM WATERCOMPONENTS WHERE OBJECT_ID_ = " & objIdNo, Conn, adOpenKeyset, adLockOptimistic
        Open "d:\temp.txt" For Append As #10
        Print #10, objIdNo & " , " & xNo & " , " & yNo & " , " & rsAtributoNos.Fields("GROUNDHEIGHT").value & " , " & CStr(zNo)
        Close #10
        rsAtributoNos.Close
        'Conn.execute "UPDATE WATERCOMPONENTS SET GROUNDHEIGHT = " & CStr(zNo) & " WHERE OBJECT_ID_ = " & objIdNo
    Loop
    Close #10
    Exit Sub
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CAcertaZsDosNos", "AtribuiZs", CStr(Err.Number), CStr(Err.Description), True, False
   End If
End Sub


