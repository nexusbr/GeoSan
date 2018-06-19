VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObtemObjIDsRamais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Obtem os objIds de todos os ramais ligados a um determinado trecho de rede
'
' Recebe o objID do trecho de rede, a conexão com o banco de dados e o ponteiro para a lista de ObjIds dos ramais
'
' objIdTrecho - objID do trecho de rede (linha)
' teDatabaseConn - conexão com o banco de dados
' objIDs - retorno da lista de objIDs dos ramais que estão ligados ao trecho de rede fornecido
'
Public Sub getObjIDs(objIdTrecho As String, ByRef objIDs() As String, ByRef geomIDs() As String)
'ByVal teDatabaseConn As TeDatabase,
    On Error GoTo Trata_Erro:
    Dim rsRamais As ADODB.Recordset
    Dim rsGeomRamais As ADODB.Recordset
    Dim i As Integer
    
    i = 0
    Set rsRamais = New ADODB.Recordset
    Set rsGeomRamais = New ADODB.Recordset
    rsRamais.Open "Select OBJECT_ID_ From RAMAIS_AGUA WHERE OBJECT_ID_TRECHO = '" & objIdTrecho & "'", Conn, adOpenDynamic, adLockOptimistic
    If rsRamais.EOF = False Then                         'se existem ramais conectados ao trecho de rede
        Do While rsRamais.EOF = False
            rsGeomRamais.Open "Select geom_id from Lines7 where object_id = '" & rsRamais("OBJECT_ID_").value & "'", Conn, adOpenDynamic, adLockOptimistic
            ReDim Preserve geomIDs(i) As String
            If rsGeomRamais.EOF = False Then
                geomIDs(i) = rsGeomRamais("geom_id").value
            Else
                MsgBox ("O ramal de object_id " & rsRamais("OBJECT_ID_").value & " não possui geometria associada em Lines7")
            End If
            rsGeomRamais.Close
            ReDim Preserve objIDs(i) As String
            objIDs(i) = rsRamais("OBJECT_ID_").value
            i = i + 1
            rsRamais.MoveNext
        Loop
    Else                                                'se não existem ramais conectados ao trecho de rede
        ReDim Preserve objIDs(i) As String
        objIDs(0) = ""
    End If
    rsRamais.Close
    Exit Sub
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CObtemObjIDsRamais", "getObjIDs", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Sub
