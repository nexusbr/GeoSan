VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppMode"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarMyUsers As NexusUsers.clsUsers
'Private mvarMyUsers As Object

'Public Property Let MyUsers(ByVal vData As Object)
'    mvarMyUsers = vData
'End Property
'
'Public Property Get MyUsers() As Object
'  If mvarMyUsers Is Nothing Then
'     Set mvarMyUsers = CreateObject("NexusUsers.clsUsers")
'  End If
'  Set MyUsers = mvarMyUsers
'End Property

Public Property Let MyUsers(ByVal vData As NexusUsers.clsUsers)
    mvarMyUsers = vData
End Property

Public Property Get MyUsers() As NexusUsers.clsUsers
  If mvarMyUsers Is Nothing Then
     Set mvarMyUsers = New NexusUsers.clsUsers ' CreateObject("NexusUsers.clsUsers")
  End If
  Set MyUsers = mvarMyUsers
End Property

Public Function OpenGroups(Conn As ADODB.Connection) As Boolean
   OpenGroups = FrmGroups.Init(Conn)
End Function

Public Function OpenUserChangePwd(Conn As ADODB.Connection, UserID As Long) As Boolean
   OpenUserChangePwd = FrmUserChangePwd.Init(Conn, UserID)
End Function

Public Function OpeNexusUsers(Conn As ADODB.Connection) As Boolean
   OpeNexusUsers = FrmUsers.Init(Conn)
End Function

Public Function OpenLogin(Conn As ADODB.Connection) As String
   OpenLogin = FrmLog.Init(Conn)
End Function
Public Function OpenUsers(Conn As ADODB.Connection) As Boolean
   OpenUsers = FrmUsers.Init(Conn)
End Function

