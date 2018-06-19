VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDatabaseNexus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para manipular e consultar dados das tabelas do Terralib
'
'
'
' Obtem o nome da coluna que contem o object_id na tabela de atributos
'
' Retorna o nome da coluna que contem o object_id na tabela de atributos
' mdt - nome do layer
'
Public Function GetNomeColunaObjectId(mdt As String) As String
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    
    rs.Open ("SELECT unique_id FROM TE_LAYER_TABLE WHERE ATTR_TABLE = '" & mdt & "'"), Conn, adOpenKeyset, adLockOptimistic
    If rs.EOF = False Then
        GetNomeColunaObjectId = rs.Fields("unique_id").value
    Else
        MsgBox ("Não encontra o nome da coluna object_id no layer MDT.")
    End If
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "CDatabaseNexus", "GetNomeColunaObjectId", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function

