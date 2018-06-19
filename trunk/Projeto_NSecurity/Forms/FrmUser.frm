VERSION 5.00
Begin VB.Form FrmUser 
   Caption         =   "Usuário"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3105
   Icon            =   "FrmUser.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4035
      Left            =   120
      TabIndex        =   8
      Top             =   30
      Width           =   2865
      Begin VB.Frame Frame2 
         Caption         =   "Perfil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   255
         TabIndex        =   12
         Top             =   1695
         Width           =   2370
         Begin VB.OptionButton Option5 
            Caption         =   "Visitante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   270
            TabIndex        =   5
            Top             =   1080
            Width           =   1290
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Usuário"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   270
            TabIndex        =   4
            Top             =   705
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Administrador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   270
            TabIndex        =   3
            Top             =   345
            Width           =   1590
         End
      End
      Begin VB.TextBox txtUsrPwd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1155
         Width           =   1755
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1515
         TabIndex        =   7
         Top             =   3420
         Width           =   1080
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   255
         TabIndex        =   6
         Top             =   3420
         Width           =   1080
      End
      Begin VB.TextBox txtUsrLog 
         Height          =   330
         Left            =   840
         TabIndex        =   1
         Top             =   735
         Width           =   1755
      End
      Begin VB.TextBox txtUsrNom 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Senha"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Login"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyConn As ADODB.Connection
Private MyRs As ADODB.Recordset
Private MyUsers As New NexusUsers.clsUsers
Private Itmx As ListItem
Private ChangePwd As Boolean

Public Function Init(Conn As ADODB.Connection, Optional UsrId As Long) As Boolean
   Dim ItmxFind As ListItem
   Set MyConn = Conn
   If UsrId > 0 Then
      Me.Caption = Me.Caption & " - Alteração" 'ALTERAÇÃO DE USUÁRIO
      If MyUsers.Users.SelectData(Conn, UsrId) Then
         txtUsrLog.Text = MyUsers.Users.UsrLog
         txtUsrPwd.Text = MyUsers.Users.UsrPwd
         txtUsrNom.Text = MyUsers.Users.UsrNom
'         chkUserBrk.Value = IIf(MyUsers.Users.UsrBrk, 1, 0)
'         chkUserExp.Value = IIf(MyUsers.Users.UsrExp, 1, 0)
'         Set MyRs = MyUsers.Users.Groups.SelectUserGroups(Conn, UsrId)
'         If MyRs.State = adStateOpen Then
'            While Not MyRs.EOF
'               Set Itmx = LvUserGroups.ListItems.Add(, , MyRs.Fields("GrpNom").Value, , 1)
'                  Itmx.Tag = MyRs.Fields("GrpID").Value
'               MyRs.MoveNext
'            Wend
'            MyRs.Close
'         End If
'         Set MyRs = Nothing
      End If
   Else
      Me.Caption = Me.Caption & " - Inclusão" 'INCLUSÃO DE NOVO USUÁRIO
   End If
'   Set MyRs = MyUsers.Groups.SelectGroups(Conn)
'   If MyRs.State = adStateOpen Then
'      While Not MyRs.EOF
'         Set ItmxFind = LvUserGroups.FindItem(MyRs.Fields("GrpNom").Value, lvwText, , lvwPartial)
'         If ItmxFind Is Nothing Then
'            Set Itmx = LvGroups.ListItems.Add(, , MyRs.Fields("GrpNom").Value, , 2)
'               Itmx.Tag = MyRs.Fields("GrpID").Value
'         End If
'         MyRs.MoveNext
'      Wend
'      MyRs.Close
'   End If
'   Set MyRs = Nothing
'   LvGroups.ColumnHeaders.Item(1).Width = LvGroups.Width - 100
'   LvUserGroups.ColumnHeaders.Item(1).Width = LvGroups.Width - 100
   ChangePwd = False
   Me.Show vbModal
   Set MyUsers = Nothing
End Function

Private Sub cmdCancel_Click()
   Unload Me
End Sub
'
'Private Sub cmdDeleteGroup_Click()
'   If Not LvUserGroups.SelectedItem Is Nothing Then
'      Set Itmx = LvGroups.ListItems.Add(, , LvUserGroups.SelectedItem.Text, , 1)
'         Itmx.Tag = LvUserGroups.SelectedItem.Tag
'         LvUserGroups.ListItems.Remove LvUserGroups.SelectedItem.Index
'   End If
'End Sub

'Private Sub cmdInsertGroup_Click()
'   If Not LvGroups.SelectedItem Is Nothing Then
'      Set Itmx = LvUserGroups.ListItems.Add(, , LvGroups.SelectedItem.Text, , 1)
'         Itmx.Tag = LvGroups.SelectedItem.Tag
'         LvGroups.ListItems.Remove LvGroups.SelectedItem.Index
'   End If
'End Sub

Private Sub cmdOk_Click()
   On Error GoTo cmdOK_Error
   Dim Cont As Integer
   With MyUsers.Users
'      If ChangePwd Then
'         If Not FrmUsersPwdConfirm.Init(txtUsrPwd.Text) Then
'            MsgBox "Senha não confirmada", vbExclamation
'            Exit Sub
'         End If
'      End If
      MyConn.BeginTrans
      .UsrLog = txtUsrLog.Text
      .UsrPwd = txtUsrPwd.Text
      .UsrNom = txtUsrNom.Text
'      .UsrBrk = IIf(chkUserBrk.Value = 1, True, False)
      .UsrBrk = True
'      .UsrExp = IIf(chkUserExp.Value = 1, True, False)
      .UsrExp = True
      .UsrFun = 1
      .UsrDep = 1

      If .UsrId > 0 Then
         If Not .UpdateData(MyConn) Or Not .DeleteUserGroups(MyConn) Then
            MyConn.RollbackTrans
            Exit Sub
         End If
      Else
         .UsrId = .InsertData(MyConn)
         If .UsrId = 0 Then
            MyConn.RollbackTrans
            Exit Sub
         End If
      End If
'      For Cont = 1 To LvUserGroups.ListItems.Count
'         If Not MyUsers.Groups.InsertUserGroups(MyConn, .UsrId, LvUserGroups.ListItems.Item(Cont).Tag) Then
'            MyConn.RollbackTrans
'            Exit Sub
'         End If
'      Next
      MyConn.CommitTrans
   End With
   Unload Me
   Exit Sub
cmdOK_Error:
   MyConn.RollbackTrans
   MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not (MyRs Is Nothing) Then
      If MyRs.State = adStateOpen Then MyRs.Close
   End If
   Set MyRs = Nothing
   Set MyConn = Nothing
   Set MyUsers = Nothing
   Set Itmx = Nothing
End Sub

Private Sub txtUsrPwd_Change()
   ChangePwd = True
End Sub


