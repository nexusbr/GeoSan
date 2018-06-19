VERSION 5.00
Begin VB.Form FrmLog 
   BackColor       =   &H80000005&
   Caption         =   "GeoSan - Acesso"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   ControlBox      =   0   'False
   Icon            =   "FrmLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsrLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1215
      TabIndex        =   0
      Top             =   1185
      Width           =   2085
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1215
      TabIndex        =   2
      Top             =   2175
      Width           =   975
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1695
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2310
      TabIndex        =   3
      Top             =   2175
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   705
      Picture         =   "FrmLog.frx":0320
      Top             =   120
      Width           =   2040
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1215
      Width           =   900
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   1725
      Width           =   810
   End
End
Attribute VB_Name = "FrmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Logou As Boolean
Private UName As Long
Private MyConn As ADODB.Connection

Private Usr As New NexusUsers.clsUsers

Public Function Init(Conn As ADODB.Connection) As Long
   Set MyConn = Conn
   
   Me.Show vbModal
   Init = UName
   Set Usr = Nothing
End Function

Private Sub cmdCancel_Click()
   Set Usr = Nothing
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If Not Usr.FindUser(MyConn, txtUsrLog) Then
      MsgBox "Usuario não encontrado", vbExclamation
      'MsgBox "gustavo, vbExclamation"
   Else
      If Usr.UsrBrk = True Then
         MsgBox "Usuário bloqueado, contate o Administrador.", vbInformation, "Acesso Negado"

      ElseIf Usr.UsrPwd <> txtUsrPwd Then
         MsgBox "Senha inválida, tente novamente", vbInformation, "Acesso Negado"
         txtUsrPwd.SelLength = Len(txtUsrPwd.Text)
         
         
      Else
         UName = Usr.UsrId
         Logou = True
         Unload Me
      End If
   End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      If txtUsrLog.Text = "" Then
         txtUsrLog.SetFocus
      ElseIf txtUsrPwd.Text = "" Then
         txtUsrPwd.SetFocus
      Else
         cmdOK_Click
      End If
   End If
End Sub

