VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRcsId As Long 'local copy
Private mvarRcsNom As String 'local copy

Public Property Let RcsNom(ByVal vData As String)
    mvarRcsNom = vData
End Property

Public Property Get RcsNom() As String
    RcsNom = mvarRcsNom
End Property

Public Property Let RcsID(ByVal vData As Long)
    mvarRcsId = vData
End Property

Public Property Get RcsID() As Long
    RcsID = mvarRcsId
End Property

Public Function SelectResources(Conn As ADODB.Connection) As ADODB.Recordset
   Dim rs As ADODB.Recordset
    If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select * From SystemResources")
   Else
    Set rs = Conn.Execute("Select * From " + """" + "SYSTEMRESOURCES" + """" + "")
   
   End If
   Set SelectResources = rs
   Set rs = Nothing
End Function


Public Function SelectResourcesUser(Conn As ADODB.Connection, MyUserId As Long, MyRcsId As Integer) As String
   Dim rs As ADODB.Recordset, R(3) As Integer
    If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select * From SystemResourcesGroups where RcsID = " & MyRcsId & " and  GrpID in ( " & _
                         "Select grpid from SystemUsersGroups where Usrid = " & MyUserId & ")")
     Else
     Dim dd As String
     Dim de As String
     Dim di As String
     Dim du As String
     Dim df As String
     dd = "SYSTEMRESOURCESGROUPS"
     de = "RcsID"
     di = "SYSTEMUSERSGROUPS"
     du = "USRID"
     df = "GRPID"
    Set rs = Conn.Execute("Select * From " + """" + dd + """" + " where " + """" + de + """" + " = '" & MyRcsId & "' and  GrpID in ( " & _
                         "Select " + """" + df + """" + " from " + """" + di + """" + " where " + """" + du + """" + " = '" & MyUserId & "')")
                         
    End If
   While Not rs.EOF
     R(0) = IIf(rs(0) = 1, rs(0), IIf(rs.Fields("rcssel").Value, 1, 0))
     R(1) = IIf(rs(1) = 1, rs(1), IIf(rs.Fields("rcsupd").Value, 1, 0))
     R(2) = IIf(rs(2) = 1, rs(2), IIf(rs.Fields("rcsins").Value, 1, 0))
     R(3) = IIf(rs(3) = 1, rs(3), IIf(rs.Fields("rcsdel").Value, 1, 0))
     rs.MoveNext
   Wend
   SelectResourcesUser = R(0) & R(1) & R(2) & R(3)
   rs.Close
   Set rs = Nothing
End Function

Public Function SelectResourcesByName(Conn As ADODB.Connection, RcsNom As String, RcsID As Integer) As Boolean
   On Error GoTo SelectResourcesByName_err
Dim rs As ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
a = "RCSID"
b = "SYSTEMRESOURCES"
c = "RCSNOM"
        If Form1.TipoConexao <> 4 Then
   Set rs = Conn.Execute("Select RcsID From SystemResources where RcsNom = '" & RcsNom & "'")
   Else
      Set rs = Conn.Execute("Select " + """" + a + """" + " From " + """" + b + """" + " where " + """" + c + """" + " = '" & RcsNom & "'")
   End If
   If Not rs.EOF Then
      RcsID = rs.Fields("RcsId").Value
      SelectResourcesByName = True
   End If
   rs.Close
   Set rs = Nothing
   Exit Function
SelectResourcesByName_err:
   MsgBox "SelectResourcesByName" & Err.Description
End Function
