VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGroups 
   Caption         =   "Gerenciador de Grupos, Recursos e Permissão"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   Icon            =   "FrmGroups.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         Height          =   435
         Left            =   6090
         TabIndex        =   16
         Top             =   4320
         Width           =   945
      End
      Begin VB.Frame Frame3 
         Caption         =   "Recursos"
         Height          =   3075
         Left            =   3270
         TabIndex        =   8
         Top             =   990
         Width           =   3795
         Begin VB.CheckBox ChkDelete 
            Caption         =   "Excluir"
            Height          =   285
            Left            =   2730
            TabIndex        =   15
            Top             =   1920
            Width           =   945
         End
         Begin VB.CheckBox ChkInsert 
            Caption         =   "Inserir"
            Height          =   285
            Left            =   2730
            TabIndex        =   14
            Top             =   1410
            Width           =   945
         End
         Begin VB.CheckBox ChkUpdate 
            Caption         =   "Atualizar"
            Height          =   285
            Left            =   2730
            TabIndex        =   13
            Top             =   900
            Width           =   945
         End
         Begin VB.CheckBox chkRead 
            Caption         =   "Ler"
            Height          =   285
            Left            =   2730
            TabIndex        =   12
            Top             =   390
            Width           =   945
         End
         Begin VB.CommandButton cmdDeleteResource 
            Caption         =   "Remover"
            Height          =   345
            Left            =   1590
            TabIndex        =   11
            Top             =   2550
            Width           =   945
         End
         Begin VB.CommandButton cmdAddResource 
            Caption         =   "Adicionar"
            Height          =   345
            Left            =   270
            TabIndex        =   10
            Top             =   2550
            Width           =   945
         End
         Begin MSComctlLib.ListView LvResources 
            Height          =   2115
            Left            =   270
            TabIndex        =   9
            Top             =   300
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   3731
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Novo"
         Height          =   405
         Left            =   5550
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtGrpNom 
         Height          =   345
         Left            =   1020
         TabIndex        =   4
         Top             =   330
         Width           =   3915
      End
      Begin VB.Frame Frame2 
         Caption         =   "Grupos"
         Height          =   3075
         Left            =   180
         TabIndex        =   2
         Top             =   990
         Width           =   2925
         Begin VB.CommandButton cmdDeleteGroup 
            Caption         =   "Remover"
            Height          =   345
            Left            =   1560
            TabIndex        =   7
            Top             =   2520
            Width           =   945
         End
         Begin MSComctlLib.ListView LvGroups 
            Height          =   2115
            Left            =   270
            TabIndex        =   3
            Top             =   300
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   3731
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   180
         TabIndex        =   1
         Top             =   4320
         Width           =   945
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2730
         Top             =   4200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGroups.frx":0320
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGroups.frx":099B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type MyAcs
   acsSelect As Boolean
   acsUpdate As Boolean
   acsInsert As Boolean
   acsDelete As Boolean
   acsGrpId As Integer
   acsRcsId As Integer
End Type

Private Conn As ADODB.Connection
Private MyRs As ADODB.Recordset
Private tipoDeConexao As String

'Private MyUsers As Object
Private MyUsers As New NexusUsers.clsUsers
Private Itmx As ListItem
Private MyArray() As MyAcs
Private Icount As Integer
Private TInit As Boolean
Private RsR As ADODB.Recordset



Public Static Function TipoConexao(tipo As String) As String
tipoDeConexao = tipo



End Function
Public Static Function TipoConexao2() As String

TipoConexao2 = tipoDeConexao


End Function

Public Function Init(MyConn As ADODB.Connection) As Boolean

   ReDim MyArray(0)
   Dim Rs As ADODB.Recordset
   TipoConexao (MyConn.Provider)

   TInit = True
   Icount = 0
   
   Set Conn = MyConn
   
   Set RsR = New ADODB.Recordset
   RsR.CursorLocation = adUseClient
   RsR.CursorType = adOpenDynamic
   Set RsR = MyUsers.Resources.SelectResources(Conn)
   
   Set Rs = New ADODB.Recordset
   Set MyRs = MyUsers.Users.Groups.SelectGroups(Conn)
   
   While Not MyRs.EOF
      Set Itmx = LvGroups.ListItems.Add(, , MyRs.Fields("GrpNom").Value, , 1)
         Itmx.Tag = MyRs.Fields("GrpID").Value
         Set Rs = MyUsers.ResourcesGroups.SelectResourcesByGroup(Conn, Itmx.Tag)
         While Not Rs.EOF
            With MyUsers.ResourcesGroups
               If .SelectResource(Conn, CInt(Itmx.Tag), CInt(Rs.Fields("RcsId"))) Then
                  ReDim Preserve MyArray(Icount)
                  
                  MyArray(Icount).acsRcsId = .RcsID
                  MyArray(Icount).acsGrpId = .GrpId
                  MyArray(Icount).acsDelete = .RcsDel
                  MyArray(Icount).acsInsert = .RcsIns
                  MyArray(Icount).acsSelect = .RcsSel
                  MyArray(Icount).acsUpdate = .RcsUpd
                  
                  Icount = Icount + 1
               End If
            End With
            Rs.MoveNext
         Wend
         Rs.Close
         Set Rs = Nothing
      MyRs.MoveNext
   Wend
   
   If LvGroups.ListItems.Count > 0 Then
      LvGroups_ItemClick LvGroups.ListItems.Item(1)
   End If
   
   MyRs.Close
   
   Set MyRs = Nothing
   
   LvGroups.ColumnHeaders.Item(1).Width = LvGroups.Width - 100
   LvResources.ColumnHeaders.Item(1).Width = LvGroups.Width - 100
   
   TInit = False
   
   Me.Show vbModal
   
   Init = TInit
   
End Function

Private Function LoadResources(GrpId As Integer) As Boolean

    Dim A As Integer
    
    LvResources.ListItems.Clear
    For A = 0 To Icount - 1
        
        If GrpId = MyArray(A).acsGrpId Then
            
            RsR.Filter = "RcsID=" & MyArray(A).acsRcsId
            
            If Not RsR.EOF Then
                
                Set Itmx = LvResources.ListItems.Add(, , RsR.Fields("RcsNom").Value, , 2)
                Itmx.Tag = RsR.Fields("RcsID").Value
            
            End If
        
        End If
    
    Next
   
    If LvResources.ListItems.Count > 0 Then
        LvResources_ItemClick LvResources.ListItems.Item(1)
    Else
        ClearChecks
    End If
   
End Function

Private Sub cmdAddResource_Click()

   Dim Itmx As ListItems, lt As ListItem
   Dim Frm As FrmResources
   Dim A As Integer

   If Not (LvGroups.SelectedItem Is Nothing) Then
      Set Frm = New FrmResources
      If Frm.Init(Conn, LvGroups.SelectedItem.Tag, Itmx) Then
         
         For A = 1 To Itmx.Count
            If Not Itmx.Item(A).Tag = "" Then
               ReDim Preserve MyArray(Icount)
               MyArray(Icount).acsRcsId = Itmx.Item(A).Tag
               MyArray(Icount).acsGrpId = LvGroups.SelectedItem.Tag
               
               Set lt = LvResources.ListItems.Add(, , Itmx.Item(A).Text, , 2)
                  lt.Tag = Itmx.Item(A).Tag
                  If A = 1 Then
                     lt.Selected = True
                     lt.EnsureVisible
                  End If
               Icount = Icount + 1
            End If
         Next
         
         ClearChecks
      End If
   Else
      MsgBox "Selecion um o grupo para atribui um recurso", vbExclamation
   End If
   
   Set Frm = Nothing
   Set Itmx = Nothing
   
End Sub

Private Sub cmdCancel_Click()

   Unload Me
   
End Sub

Private Sub cmdDeleteGroup_Click()

   Dim A As Integer
   If Not LvGroups.SelectedItem Is Nothing Then
      If MsgBox("Deseja realmente excluir este grupo", 36) = vbYes Then
         LvGroups.ListItems.Remove LvGroups.SelectedItem.Index
         For A = 0 To Icount - 1
            If LvGroups.SelectedItem.Tag = MyArray(A).acsGrpId Then
               MyArray(A).acsGrpId = 0
               MyArray(A).acsRcsId = 0
               
               Exit For
            End If
         Next
         
      End If
   End If
   
End Sub

Private Sub cmdDeleteResource_Click()

   Dim A As Integer
   If Not LvResources.SelectedItem Is Nothing Then
      For A = 0 To Icount - 1
         If LvResources.SelectedItem.Tag = MyArray(A).acsRcsId And LvGroups.SelectedItem.Tag = MyArray(A).acsGrpId Then
            MyArray(A).acsGrpId = 0
            MyArray(A).acsRcsId = 0
            LvResources.ListItems.Remove LvResources.SelectedItem.Index
            Exit For
         End If
      Next
   End If
   
End Sub

Private Sub cmdNew_Click()

   Dim A As Integer, B As Integer
   If Len(txtGrpNom.Text) > 4 Then
      Set Itmx = LvGroups.ListItems.Add(, , txtGrpNom.Text, , 1)
      For A = 0 To Icount - 1
         If MyArray(A).acsGrpId > B Then
            B = MyArray(A).acsGrpId
         End If
      Next
      Itmx.Tag = 0 'B + 1
   End If
   
End Sub

Private Sub cmdOK_Click()
On Error GoTo cmd_ok_error
   
   RsR.Close
   Set RsR = Nothing
   Dim A As Integer, MyGroups As String
   For A = 1 To LvGroups.ListItems.Count
      With MyUsers.Users.Groups
         If Not .SelectData(Conn, LvGroups.ListItems.Item(A).Tag) Then
            .GrpNom = LvGroups.ListItems.Item(A).Text
            .InsertData Conn
         End If
         If A = 1 Then
            MyGroups = LvGroups.ListItems.Item(A).Tag
         Else
            MyGroups = MyGroups & "," & LvGroups.ListItems.Item(A).Tag
         End If
      End With
   Next
   'Conn.BeginTrans
   If FrmGroups.TipoConexao2 <> 4 Then
   Conn.Execute "delete from SystemResourcesGroups"
   Conn.Execute "delete from SystemUsersGroups where GrpId not In(" & MyGroups & ")"
   Conn.Execute "delete from SystemGroups where GrpId not In(" & MyGroups & ")"
   Else
   Dim da As String
    Dim de As String
    Dim di As String
    da = "SystemResourcesGroups"
    de = "GrpId"
    di = "SystemGroups"
   Conn.Execute "delete from " + """" + da + """" + ""
   Conn.Execute "delete from " + """" + da + """" + " where " + """" + de + """" + " not In('" & MyGroups & "')"
   Conn.Execute "delete from " + """" + di + """" + " where " + """" + de + """" + " not In('" & MyGroups & "')"
   End If
   For A = 0 To Icount - 1
      With MyArray(A)
         If .acsGrpId > 0 And .acsRcsId > 0 Then
            MyUsers.Users.ResourcesGroups.RcsDel = .acsDelete
            MyUsers.Users.ResourcesGroups.RcsIns = .acsInsert
            MyUsers.Users.ResourcesGroups.RcsSel = .acsSelect
            MyUsers.Users.ResourcesGroups.RcsUpd = .acsUpdate
            MyUsers.Users.ResourcesGroups.RcsID = .acsRcsId
            MyUsers.Users.ResourcesGroups.GrpId = .acsGrpId
            MyUsers.Users.ResourcesGroups.InsertData Conn
         End If
      End With
   Next
   'Conn.CommitTrans
   Set MyUsers = Nothing
   Unload Me
   
   Exit Sub
   
cmd_ok_error:
   'Conn.RollbackTrans
   MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If Not (MyRs Is Nothing) Then
      If MyRs.State = adStateOpen Then MyRs.Close
   End If
   Set MyRs = Nothing
   If Not (RsR Is Nothing) Then
      If RsR.State = adStateOpen Then RsR.Close
   End If
   
   Set RsR = Nothing
   Set Conn = Nothing
   Set MyUsers = Nothing
   Set Itmx = Nothing
   ReDim MyArray(0)
   
End Sub

Private Sub LvGroups_ItemClick(ByVal Item As MSComctlLib.ListItem)

   LoadResources Item.Tag
   
End Sub

Private Sub LvResources_ItemClick(ByVal Item As MSComctlLib.ListItem)

   Dim A As Integer
   
   TInit = True
   For A = 0 To Icount - 1
      If Item.Tag = MyArray(A).acsRcsId And LvGroups.SelectedItem.Tag = MyArray(A).acsGrpId Then
         ChkDelete.Value = IIf(MyArray(A).acsDelete, 1, 0)
         ChkInsert.Value = IIf(MyArray(A).acsInsert, 1, 0)
         chkRead.Value = IIf(MyArray(A).acsSelect, 1, 0)
         ChkUpdate.Value = IIf(MyArray(A).acsUpdate, 1, 0)
         Exit For
      End If
   Next
   
   TInit = False
   
End Sub

Private Sub ClearChecks()

   TInit = True
   ChkDelete.Value = 0
   ChkInsert.Value = 0
   chkRead.Value = 0
   ChkUpdate.Value = 0
   TInit = False
   
End Sub

Private Sub chkRead_Click()

   UpdatePermision
   
End Sub

Private Sub ChkDelete_Click()

   UpdatePermision
   
End Sub

Private Sub ChkInsert_Click()

   UpdatePermision
   
End Sub

Private Sub ChkUpdate_Click()

   UpdatePermision
   
End Sub

Private Sub UpdatePermision()

    Dim A As Integer
    
    If Not LvResources.SelectedItem Is Nothing And Not TInit Then
        For A = 0 To Icount - 1
            If MyArray(A).acsRcsId = LvResources.SelectedItem.Tag And _
                MyArray(A).acsGrpId = LvGroups.SelectedItem.Tag Then
                MyArray(A).acsDelete = IIf(ChkDelete.Value = 0, False, True)
                MyArray(A).acsInsert = IIf(ChkInsert.Value = 0, False, True)
                MyArray(A).acsSelect = IIf(chkRead.Value = 0, False, True)
                MyArray(A).acsUpdate = IIf(ChkUpdate.Value = 0, False, True)
                
                Exit For
            End If
        Next
    End If
   
End Sub
