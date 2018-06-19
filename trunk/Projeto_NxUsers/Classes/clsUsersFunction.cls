VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsersFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarFunId As Long 'local copy
Private mvarFunNom As String 'local copy
Public Property Let FunNom(ByVal vData As String)
    mvarFunNom = vData
End Property

Public Property Get FunNom() As String
    FunNom = mvarFunNom
End Property

Public Property Let FunId(ByVal vData As Long)
    mvarFunId = vData
End Property

Public Property Get FunId() As Long
    FunId = mvarFunId
End Property

Public Function UpdateData(Conn As ADODB.Connection) As Boolean

End Function

Public Function SelectData(Conn As ADODB.Connection) As Boolean

End Function

Public Function InsertData(Conn As ADODB.Connection) As Boolean

End Function


