VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsers 
   Caption         =   "Usuários"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   Icon            =   "FrmsUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   3375
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2340
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmsUsers.frx":0320
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   375
         Left            =   1170
         TabIndex        =   2
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Novo"
         Height          =   375
         Left            =   210
         TabIndex        =   1
         Top             =   2880
         Width           =   855
      End
      Begin MSComctlLib.ListView LvUsers 
         Height          =   2535
         Left            =   210
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
End
Attribute VB_Name = "FrmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Conn As ADODB.Connection

Private Sub cmdEdit_Click()
   If Not (LvUsers.SelectedItem Is Nothing) Then
      FrmUser.Init Conn, LvUsers.SelectedItem.Tag
   Else
      MsgBox "Selecione um usuário", vbExclamation
   End If
End Sub

Public Function Init(MyConn As ADODB.Connection) As Boolean
   On Error GoTo Users_Init_Error
   Set Conn = MyConn
   UpdateForm
   LvUsers.ColumnHeaders.Item(1).Width = LvUsers.Width - 350
   Me.Show vbModal
   Init = True
   Exit Function
Users_Init_Error:
   MsgBox "Users_Init_Error" & " " & Err.Description
End Function

Private Sub cmdNew_Click()
   FrmUser.Init Conn, 0
   UpdateForm
End Sub

Private Sub UpdateForm()
   Dim Rs As ADODB.Recordset
   Dim MyUsers As New NexusUsers.clsUsers
   Set Rs = MyUsers.Users.SelectAllUsers(Conn)
   LvUsers.ListItems.Clear
   Dim Itmx As ListItem
   While Not Rs.EOF
      Set Itmx = LvUsers.ListItems.Add(, , Rs.Fields("UsrNom").Value, , 1)
         Itmx.Tag = Rs.Fields("UsrId").Value
      Rs.MoveNext
   Wend
   Rs.Close
   Set Rs = Nothing
   Set MyUsers = Nothing
End Sub

