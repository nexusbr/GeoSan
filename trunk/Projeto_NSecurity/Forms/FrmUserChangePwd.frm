VERSION 5.00
Begin VB.Form FrmUserChangePwd 
   Caption         =   "Alteração de Senha"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   Icon            =   "FrmUserChangePwd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3045
      Left            =   0
      TabIndex        =   5
      Top             =   -60
      Width           =   3825
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   2490
         TabIndex        =   10
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox txtUsrPwd 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "txtUsrPwd"
         Top             =   810
         Width           =   1755
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   345
         Left            =   270
         TabIndex        =   4
         Top             =   2430
         Width           =   975
      End
      Begin VB.TextBox txtUsrPwdNewCon 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "txtUsrPwd"
         Top             =   1860
         Width           =   1755
      End
      Begin VB.TextBox txtUsrLog 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1680
         TabIndex        =   0
         Text            =   "txtUsrLog"
         Top             =   300
         Width           =   1755
      End
      Begin VB.TextBox txtUsrPwdNew 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "txtUsrPwd"
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Senha Atual"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Confirme a Senha"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Usuário"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Nova Senha"
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1380
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmUserChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyConn As ADODB.Connection
'Private MyUsers As Object
Private MyUsers As New NexusUsers.clsUsers
Private ChangePwd As Boolean

Public Function Init(Conn As ADODB.Connection, UserID As Long) As Boolean
   Set MyConn = Conn
   ' Set MyUsers = CreateObject("NexusUsers.clsUsers")
   With MyUsers.Users
      If .SelectData(Conn, UserID) Then
         txtUsrLog.Text = .UsrLog
         txtUsrPwd.Text = ""
         txtUsrPwdNew.Text = ""
         txtUsrPwdNewCon.Text = ""
      End If
   End With
   Me.Show vbModal
   Init = ChangePwd
   Set MyUsers = Nothing
End Function

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOk_Click()
   With MyUsers.Users
      If txtUsrPwd.Text = .UsrPwd Then
         If txtUsrPwdNew.Text = txtUsrPwdNewCon.Text Then
            .UsrPwd = txtUsrPwdNew.Text
            .UpdateData MyConn
            ChangePwd = True
            Unload Me
         Else
            MsgBox "Campos nova senha e confirmação não são iguais", vbExclamation
         End If
      Else
         MsgBox "Senha atual inválida", vbExclamation
         txtUsrPwd.Text = ""
      End If
   End With
End Sub
