VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TECOMD~1.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEPAPatterns 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Padrões"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   345
      Left            =   4140
      TabIndex        =   7
      Top             =   2100
      Width           =   945
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5160
      TabIndex        =   6
      Top             =   2100
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   6015
      Begin VB.TextBox txtDescricao 
         Height          =   345
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   2
         Top             =   420
         Width           =   4215
      End
      Begin VB.TextBox txtID 
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   420
         Width           =   1515
      End
      Begin MSFlexGridLib.MSFlexGrid GrdPatterns 
         Height          =   885
         Left            =   90
         TabIndex        =   3
         Top             =   900
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   1561
         _Version        =   393216
         Cols            =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "DESCRIÇÃO"
         Height          =   255
         Left            =   1770
         TabIndex        =   5
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   1455
      End
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabasePoligono 
      Left            =   1080
      OleObjectBlob   =   "frmEPAPatterns.frx":0000
      Top             =   2160
   End
End
Attribute VB_Name = "frmEPAPatterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Iniciou_Entrada As Boolean
Private Pattern As New clsEPAPatterns
Private NewID As Boolean
Public Function init(ID As Long) As Boolean
   'LoozeXP1.InitIDESubClassing
   Dim Patterns() As String, a As Integer
   GrdPatterns.TextMatrix(0, 0) = "Período"
   GrdPatterns.TextMatrix(0, 0) = "Fator Multiplicativo"
   GrdPatterns.ColWidth(0) = GrdPatterns.ColWidth(0) * 2
   For a = 1 To 24
      GrdPatterns.TextMatrix(0, a) = a
   Next
  
   If ID = 0 Then
      NewID = True
      txtID.Enabled = True
   Else
      If Pattern.Atualizar_Padrao(ID) Then
         txtID = Pattern.ID
         txtDescricao = Pattern.DESCRICAO
         Patterns = Split(Pattern.PADRAO, ";", 25)
         For a = 0 To 23
            GrdPatterns.TextMatrix(1, a + 1) = Patterns(a)
         Next
      End If
      NewID = False
      txtID.Enabled = False
   End If
   Me.Show vbModal
   'LoozeXP1.EndWinXPCSubClassing
End Function

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdConfirmar_Click()
   Dim Patterns As String, a As Integer
   For a = 1 To GrdPatterns.Cols - 1
      Patterns = Patterns & GrdPatterns.TextMatrix(1, a) & ";"
   Next
   With Pattern
      .ID = txtID
      .DESCRICAO = txtDescricao
      .PADRAO = Patterns
      If NewID Then
         If .Inserir_Padrao Then Sair_Form
      Else
         If .Atualizar_Padrao(txtID.Text, True) Then Sair_Form
      End If
   End With
End Sub

Private Sub GrdPatterns_KeyPress(KeyAscii As Integer)
   With GrdPatterns
      If Iniciou_Entrada Then
         Select Case KeyAscii
           Case vbKeyDelete, vbKeyBack
               .TextMatrix(.Row, .Col) = ""
           Case vbKeyReturn
               If .Col < .Cols Then .Col = .Col + 1
           Case Else
               .TextMatrix(.Row, .Col) = ""
               If IsNumeric(.TextMatrix(.Row, .Col) & Chr(KeyAscii)) Then
                  .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) & Chr(KeyAscii)
               End If
         End Select
         Iniciou_Entrada = True
      Else
         Select Case KeyAscii
           Case vbKeyBack
               If Len(.TextMatrix(.Row, .Col)) > 0 Then
                  .TextMatrix(.Row, .Col) = Left(.TextMatrix(.Row, .Col), Len(.TextMatrix(.Row, .Col)) - 1)
               End If
           Case vbKeyReturn
               If .Row < .Rows Then .Col = .Col + 1
               
           Case Else
               If IsNumeric(.TextMatrix(.Row, .Col) & Chr(KeyAscii)) Then
                  .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) & Chr(KeyAscii)
               End If
         End Select
      End If
   End With
End Sub

Private Sub GrdPatterns_RowColChange()
   Iniciou_Entrada = False
End Sub

Private Sub Sair_Form()
   Set Pattern = Nothing
   Unload Me
End Sub


