VERSION 5.00
Begin VB.Form FrmUsersPwdConfirm 
   Caption         =   "Confirme a senha"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   2370
      TabIndex        =   1
      Top             =   270
      Width           =   705
   End
   Begin VB.TextBox txtUsrPwd 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   300
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   270
      Width           =   1755
   End
End
Attribute VB_Name = "FrmUsersPwdConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrPwd As String


Public Function Init(Pwd As String) As Boolean

   Me.Show vbModal
   If StrPwd = Pwd Then
      Init = True
   End If
   
End Function

Private Sub cmdOK_Click()
   StrPwd = txtUsrPwd.Text
   Unload Me
End Sub

Private Sub txtUsrPwd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOK_Click
   End If
End Sub
