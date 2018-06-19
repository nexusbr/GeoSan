VERSION 5.00
Begin VB.Form FrmMsg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entre com o novo texto"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   5250
   Icon            =   "FrmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTexto 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3915
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "FrmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'numero inteiro, decimal ou texto
Enum gsDataType
    gsString = 1
    gsInteger = 2
    gsDouble = 3
End Enum
Private tipo As gsDataType
Private Texto As String
Public Function init(Optional MsgInformation As String, Optional DataType As gsDataType = gsString) As String
   'LoozeXP1.InitSubClassing
   If MsgInformation <> "" Then Me.Caption = MsgInformation
   Texto = ""
   tipo = DataType
   Me.Show vbModal
   init = Texto
   'LoozeXP1.EndWinXPCSubClassing
End Function

Private Sub cmd_Ok_Click()
   Texto = txtTexto
   Unload Me
End Sub

Private Sub txtTexto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmd_Ok_Click
        Exit Sub
    End If
    Select Case tipo
      
      Case gsInteger
          
          If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = vbKeyBack) Then   'gsinteger
          
              KeyAscii = 0
              MsgBox "Digite somente númenros", vbInformation
          End If
          
      Case gsDouble
          
          If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 44 Or KeyAscii = 46) Or KeyAscii = vbKeyBack) Then 'gsdecimal
              KeyAscii = 0
              MsgBox "Digite somente númenros e ponto ou virgula", vbInformation
          End If
    End Select
End Sub



