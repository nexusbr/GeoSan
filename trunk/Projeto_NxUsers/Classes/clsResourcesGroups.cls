VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResourcesGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarRcsId As Integer 'local copy
Private mvarGrpId As Integer 'local copy
Private mvarPMSId As Integer 'local copy
Private mvarRcsSel As Boolean 'local copy
Private mvarRcsIns As Boolean 'local copy
Private mvarRcsUpd As Boolean 'local copy
Private mvarRcsDel As Boolean 'local copy
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String

Public Property Let GrpId(ByVal vData As Integer)
    mvarGrpId = vData
End Property

Public Property Get GrpId() As Integer
    GrpId = mvarGrpId
End Property

Public Property Let RcsID(ByVal vData As Integer)
    mvarRcsId = vData
End Property

Public Property Get RcsID() As Integer
    RcsID = mvarRcsId
End Property

Public Property Let RcsDel(ByVal vData As Boolean)
    mvarRcsDel = vData
End Property

Public Property Get RcsDel() As Boolean
    RcsDel = mvarRcsDel
End Property

Public Property Let RcsUpd(ByVal vData As Boolean)
    mvarRcsUpd = vData
End Property

Public Property Get RcsUpd() As Boolean
    RcsUpd = mvarRcsUpd
End Property

Public Property Let RcsIns(ByVal vData As Boolean)
    mvarRcsIns = vData
End Property

Public Property Get RcsIns() As Boolean
    RcsIns = mvarRcsIns
End Property

Public Property Let RcsSel(ByVal vData As Boolean)
    mvarRcsSel = vData
End Property

Public Property Get RcsSel() As Boolean
    RcsSel = mvarRcsSel
End Property



Public Function UpdateData(Conn As ADODB.Connection) As Boolean
   On Error GoTo UpdateData_error
   Dim MyRs As ADODB.Recordset 'local copy
   Set MyRs = New ADODB.Recordset

a = "SYSTEMUSERS"
b = "USRID"
c = "RCSID"
d = "GRPID"


     If Form1.TipoConexao <> 4 Then
   MyRs.Open "Select * from  SystemUsers " & _
             "Where UsrId = " & RcsID & " and GrpId=" & GrpId, Conn, _
               adOpenKeyset, adLockOptimistic, adCmdText
               Else
                MyRs.Open "Select * from  " + """" + a + """" + _
             "Where " + """" + b + """" + " = '" & RcsID & "' and " + """" + d + """" + "='" & GrpId & "'", Conn, _
             adOpenDynamic, adLockOptimistic
               End If
   If Not MyRs.EOF Then
      MyRs.Fields("RcsSel").Value = RcsSel
      MyRs.Fields("RcsIns").Value = RcsIns
      MyRs.Fields("RcsUpd").Value = RcsUpd
      MyRs.Fields("RcsSel").Value = RcsDel
      MyRs.Update
   End If
   MyRs.Close
   UpdateData = True
   Set MyRs = Nothing
UpdateData_error:
   If Not (MyRs Is Nothing) Then
      If MyRs.State = adStateOpen Then MyRs.Close
   End If
   Set MyRs = Nothing
End Function

Public Function SelectResource(Conn As ADODB.Connection, Grp As Integer, Rcs As Integer) As Boolean
   On Error GoTo SelectResource_error
   Dim MyRs As ADODB.Recordset
   Set MyRs = New ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String



If Form1.TipoConexao <> 4 Then
MyRs.Open "Select * from SYSTEMRESOURCESGROUPS Where GRPID='" & Grp & "' and RCSID = '" & Rcs & "' ", Conn
Else
a = "SYSTEMRESOURCESGROUPS"
b = "GRPID"
c = "RCSID"
End If
   MyRs.Open "Select * from " + """" + a + """" + " Where " + """" + b + """" + "='" & Grp & "' and " + """" + c + """" + " = '" & Rcs & "' ", Conn, adOpenDynamic, adLockOptimistic
   If Not MyRs.EOF Then
      GrpId = MyRs.Fields("GrpId").Value
      RcsID = MyRs.Fields("RcsId").Value
      RcsSel = MyRs.Fields("RcsSel").Value
      RcsIns = MyRs.Fields("RcsIns").Value
      RcsUpd = MyRs.Fields("RcsUpd").Value
      RcsDel = MyRs.Fields("RcsDel").Value
      SelectResource = True
   Else
      MsgBox "Não foi encontrado nenhum registro", vbExclamation
   End If
   MyRs.Close
   Set MyRs = Nothing
SelectResource_error:
   If Not (MyRs Is Nothing) Then
      If MyRs.State = adStateOpen Then MyRs.Close
   End If
   Set MyRs = Nothing
End Function


Public Function SelectResourcesByGroup(Conn As ADODB.Connection, GrpId As Integer) As ADODB.Recordset
   Dim rs As ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String

   If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select B.* from SystemResourcesGroups A inner Join  SystemResources B on A.RcsId=B.RcsID Where GrpId=" & GrpId)
   Else
a = "SYSTEMRESOURCESGROUPS"
b = "SYSTEMRESOURCES"
c = "RCSID"
d = "GRPID"

   Set rs = Conn.Execute("Select " + """" + b + """" + "." + """" + "*" + """" + " from " + """" + a + """" + " inner Join  " + """" + b + """" + " on " + """" + a + """" + "." + """" + c + """" + "=" + """" + b + """" + "." + """" + c + """" + " Where " + """" + d + """" + "='" & GrpId & "'")
   End If
   
   Set SelectResourcesByGroup = rs
   Set rs = Nothing
End Function

Public Function InsertData(Conn As ADODB.Connection) As Boolean
   On Error GoTo InsertData_error
   Dim MyRs As ADODB.Recordset
   Set MyRs = New ADODB.Recordset
   MyRs.Open "SystemResourcesGroups", Conn, adOpenDynamic, adLockOptimistic
   MyRs.AddNew
   MyRs.Fields("RcsId").Value = RcsID
   MyRs.Fields("GrpId").Value = GrpId
   MyRs.Fields("RcsSel").Value = RcsSel
   MyRs.Fields("RcsIns").Value = RcsIns
   MyRs.Fields("RcsUpd").Value = RcsUpd
   MyRs.Fields("RcsDel").Value = RcsDel
   MyRs.Update
   MyRs.Close
   Set MyRs = Nothing
   InsertData = True
InsertData_error:
   If Not (MyRs Is Nothing) Then
      If MyRs.State = adStateOpen Then MyRs.Close
   End If
  Set MyRs = Nothing
End Function

Public Function DeleteGroupsResources(Conn As ADODB.Connection, Grp As Integer) As Boolean
   Conn.Execute "Delete From SystemResourcesGroups Where GrpID=" & Grp
   DeleteGroupsResources = True
End Function

Public Function SelectResourcesByGroupII(Conn As ADODB.Connection, mGroupsID As String, RcsID As Integer, mUpd As Boolean, mView As Boolean, mDel As Boolean, mIns As Boolean) As Boolean
   On Error GoTo SelectResourcesByGroupII_error
   Dim rs As ADODB.Recordset
   
   If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select A.* from SystemResourcesGroups A inner Join  SystemResources B on A.RcsId=B.RcsID Where GrpId in(" & mGroupsID & ") and A.RcsID = " & RcsID)
   Else
   a = "SYSTEMRESOURCESGROUPS"
b = "SYSTEMRESOURCES"
c = "RCSID"
d = "GRPID"
    Set rs = Conn.Execute("Select " + """" + a + """" + "." + """" + " *" + """" + " from " + """" + a + """" + " inner Join  " + """" + b + """" + " on " + """" + a + """" + "." + """" + c + """" + "=" + """" + b + """" + "." + """" + c + """" + " Where " + """" + d + """" + "in('" & mGroupsID & "') and " + """" + a + """" + "." + """" + c + """" + " = '" & RcsID & "'")
   End If
   While Not rs.EOF
      If rs.Fields("rcsins").Value Then mIns = True
      If rs.Fields("rcsdel").Value Then mDel = True
      If rs.Fields("rcssel").Value Then mView = True
      If rs.Fields("rcsupd").Value Then mUpd = True
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   SelectResourcesByGroupII = True
   Exit Function
SelectResourcesByGroupII_error:
   If Not (rs Is Nothing) Then
      If rs.State = adStateOpen Then rs.Close
   End If
  Set rs = Nothing
  MsgBox "SelectResourcesByGroupII" & Err.Description
End Function
