VERSION 5.00
Begin VB.Form frmAutoLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto-Login"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmAutoLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkManterLogin 
      Caption         =   "Manter login automático"
      Height          =   285
      Left            =   195
      TabIndex        =   1
      Top             =   1125
      Value           =   1  'Checked
      Width           =   2160
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   420
      Left            =   4995
      TabIndex        =   0
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   315
      Picture         =   "frmAutoLogin.frx":0442
      Top             =   315
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "Logando automaticamente com usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   990
      TabIndex        =   2
      Top             =   495
      Width           =   5220
   End
End
Attribute VB_Name = "frmAutoLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
   If Me.chkManterLogin.value = 0 Then
      If MsgBox("Você selecionou eliminar o auto-login." & Chr(13) & "O próximo login requisitará uma senha.    " & Chr(13) & Chr(13) & "Confirma esta opção?", vbDefaultButton2 + vbQuestion + vbYesNo, "") = vbYes Then
         Kill App.path & "\controles\autologin.txt"
      End If
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   Me.lblMsg.Caption = "Logando automaticamente com usuário: " & strUser
End Sub

