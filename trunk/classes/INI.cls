VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim pro As String


Public Static Function Pro1() As String


pro = mid(ReadINI("CONEXAO", "PROVEDOR", App.path & "\GEOSAN.ini"), 1, 1) ' PEGA SÓ O NÚMERO DA CONEXÃO
Pro1 = pro

End Function

