VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsersGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRscId As Long
Private mvarGrpId As Long

Public Property Let GrpId(ByVal vData As Long)
    mvarGrpId = vData
End Property

Public Property Get GrpId() As Long
    GrpId = mvarGrpId
End Property

Public Property Let RscId(ByVal vData As Long)
    mvarRscId = vData
End Property

Public Property Get RscId() As Long
    RscId = mvarRscId
End Property

Public Function UpdateData(Conn As ADODB.Connection) As Boolean

End Function

Public Function SelectData(Conn As ADODB.Connection) As Boolean

End Function

Public Function InsertData(Conn As ADODB.Connection) As Boolean

End Function

Public Function SelectGroupByUsers(Conn As ADODB.Connection, Usrid As Long, mGroups As String) As Boolean
   On Error GoTo SelectResource_error
   Dim MyRs As ADODB.Recordset
   
   If Form1.TipoConexao <> 4 Then
   Set MyRs = Conn.Execute("Select GrpId from SystemUsersGroups Where UsrId=" & Usrid)
   Else
   Dim oa As String
   Dim oe As String
   Dim oi As String
   oa = "GrpId"
   oe = "SystemUsersGroups"
   oi = "UsrId"
   Set MyRs = Conn.Execute("Select " + oa + " from  " + oe + " Where  " + oi + "='" & Usrid & "'")
   End If
   While Not MyRs.EOF
      If mGroups = "" Then
         mGroups = MyRs.Fields("GrpId").Value
      Else
         mGroups = mGroups & "," & MyRs.Fields("GrpId").Value
      End If
      SelectGroupByUsers = True
      MyRs.MoveNext
   Wend
   MyRs.Close
   Set MyRs = Nothing
   Exit Function
SelectResource_error:
   If Not (MyRs Is Nothing) Then
      If MyRs.State = adStateOpen Then MyRs.Close
   End If
   Set MyRs = Nothing
   MsgBox "SelectGroupByUsers" & Err.Description
End Function


