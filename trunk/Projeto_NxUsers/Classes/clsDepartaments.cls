VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDepartaments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarDepId As Long 'local copy
Private mvarDepNom As String 'local copy
Public Property Let DepNom(ByVal vData As String)
    mvarDepNom = vData
End Property

Public Property Get DepNom() As String
    DepNom = mvarDepNom
End Property

Public Property Let DepId(ByVal vData As Long)
    mvarDepId = vData
End Property

Public Property Get DepId() As Long
    DepId = mvarDepId
End Property

Public Function UpdateData(Conn As ADODB.Connection) As Boolean

End Function

Public Function SelectData(Conn As ADODB.Connection) As Boolean

End Function

Public Function InsertData(Conn As ADODB.Connection) As Boolean

End Function


