VERSION 5.00
Begin VB.Form frmNetWorkLegth 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comprimento da Rede"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "aplicar"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmNetWorkLegth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tcs As TeCanvas
Private mLen As Double
Public Function init(mtcs As TeCanvas, frmOnner As Object)
   On Error GoTo init_err
   Set tcs = mtcs
   Me.Show , frmOnner
   Me.Top = frmOnner.Top + 1000
   Me.Left = frmOnner.Left + frmOnner.Width - frmOnner.pctSfondo.Width
init_err:
End Function

Private Sub cmdApply_Click()
   If Val(txtLength.Text) < 2 Then
      txtLength.Text = 2
      tcs.setLengthOfLastSegmentOfLine 5
   Else
      tcs.setLengthOfLastSegmentOfLine CDbl(txtLength.Text)
   End If
End Sub

Private Sub Form_Load()
   'LoozeXP1.InitIDESubClassing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'LoozeXP1.EndWinXPCSubClassing
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 44 Or KeyAscii = 46) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then 'gsdecimal
       KeyAscii = 0
       'MsgBox "Digite somente números, ponto ou virgula", vbInformation
   Else
      If KeyAscii = vbKeyReturn Then cmdApply_Click
   End If
End Sub

