VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarGrpId As Long 'local copy
Private mvarGrpNom As String 'local copy

Public Property Let GrpNom(ByVal vData As String)
    mvarGrpNom = vData
End Property

Public Property Get GrpNom() As String
    GrpNom = mvarGrpNom
End Property

Public Property Let GrpId(ByVal vData As Long)
    mvarGrpId = vData
End Property

Public Property Get GrpId() As Long
    GrpId = mvarGrpId
End Property

Public Function UpdateData(Conn As ADODB.Connection) As Boolean

End Function

Public Function SelectData(Conn As ADODB.Connection, Grp As Integer) As Boolean
Dim rs As ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
Dim d As String
a = "SYSTEMGROUPS"
b = "GRPID"
c = Grp
d = c
   If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select * from SystemGroups Where GrpId = " & Grp)
   Else
   Set rs = Conn.Execute("Select * from " + """" + a + """" + " Where " + """" + b + """" + " = '" & Grp & "'")
   End If
   If Not rs Then
      GrpId = rs.Fields("Grpid")
      GrpNom = rs.Fields("GrpNom")
      SelectData = True
   End If
   CloseObjectRs rs
   Set rs = Nothing
End Function

Public Function SelectGroups(Conn As ADODB.Connection) As ADODB.Recordset
   Dim rs As ADODB.Recordset
   Dim a As String
Dim b As String
a = "SYSTEMGROUPS"
b = "GRPNOM"
 If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select * from SystemGroups Order By GrpNom")
   Else
    Set rs = Conn.Execute("Select * from " + """" + a + """" + " Order By " + """" + b + """" + "")
   End If
   Set SelectGroups = rs
   Set rs = Nothing
End Function

Public Function InsertData(Conn As ADODB.Connection) As Boolean
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.Open "SystemGroups", Conn, adOpenDynamic, adLockOptimistic
   rs.AddNew
   rs.Fields("GrpNom").Value = GrpNom
   rs.Update
   CloseObjectRs rs
   InsertData = True
   Set rs = Nothing
End Function

Public Function SelectUserGroups(Conn As ADODB.Connection, UserId As Long) As ADODB.Recordset
   Dim rs As ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String


a = "SYSTEMGROUPS"
b = "GRPID"
c = "SYSTEMUSERSGROUPS"
d = "SYSTEMUSERS"
e = "USRID"
f = "GRPNOM"
g = UserId
h = "'g'"

     If Form1.TipoConexao <> 4 Then

   Set rs = Conn.Execute("Select A.GrpID,A.GrpNom from (SystemGroups A " & _
                           "inner join SystemUsersGroups B " & _
                           "on A.GrpId = B.GrpID) " & _
                           "inner join SystemUsers C " & _
                           "on B.UsrId = C.UsrID " & _
                           "Where C.UsrId = " & UserId)
                           Else
                           Set rs = Conn.Execute("Select " + """" + a + """" + "." + """" + b + """" + "," + """" + a + """" + "." + """" + f + """" + " from (" + """" + a + """" + _
                           "inner join " + """" + c + """" + _
                           "on " + """" + a + """" + "." + """" + b + """" + " = " + """" + c + """" + "." + """" + b + """" + _
                           "inner join " + """" + d + """" + _
                           "on " + """" + c + """" + "." + """" + e + """" + " = " + """" + d + """" + "." + """" + e + """" + _
                           "Where " + """" + d + """" + "." + """" + e + """" + " =  '" & UserId & "'")
                           End If
   Set SelectUserGroups = rs
   Set rs = Nothing
End Function

Public Function InsertUserGroups(Conn As ADODB.Connection, User As Long, Group As Long) As Boolean
   On Error GoTo SelectUserGroups_error
   Dim MyRs As ADODB.Recordset
   Set MyRs = New ADODB.Recordset
   MyRs.Open "SystemUsersGroups", Conn, adOpenDynamic, adLockOptimistic
   MyRs.AddNew
   MyRs.Fields("UsrID").Value = User
   MyRs.Fields("GrpID").Value = Group
   MyRs.Update
   CloseObjectRs MyRs
   Set MyRs = Nothing
   InsertUserGroups = True
SelectUserGroups_error:
   CloseObjectRs MyRs
   Set MyRs = Nothing
End Function

Private Sub CloseObjectRs(ObjRS As ADODB.Recordset)
   If Not (ObjRS Is Nothing) Then
      If ObjRS.State = adStateOpen Then ObjRS.Close
   End If
End Sub

