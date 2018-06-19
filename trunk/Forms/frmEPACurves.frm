VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEPACurves 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Curva"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   11
      Top             =   5070
      Width           =   945
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   345
      Left            =   4260
      TabIndex        =   10
      Top             =   5070
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   5025
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   6165
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2985
         Left            =   2700
         TabIndex        =   12
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5265
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   0   'False
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmEPACurves.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtID 
         Height          =   345
         Left            =   90
         TabIndex        =   0
         Top             =   420
         Width           =   1515
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtDescricao 
         Height          =   345
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   1
         Top             =   420
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1170
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid GrdCoordenadas 
         Height          =   3045
         Left            =   120
         TabIndex        =   4
         Top             =   1770
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   5371
         _Version        =   393216
         Rows            =   51
         FixedCols       =   0
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "TIPO"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "DESCRIÇÃO"
         Height          =   255
         Left            =   1770
         TabIndex        =   7
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "EQUAÇÃO"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1770
         TabIndex        =   6
         Top             =   930
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmEPACurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Iniciou_Entrada As Boolean
Private Curves As New clsEPACurves
Private NewID As Boolean
Public Function init(ID As Long) As Boolean
   'LoozeXP1.InitIDESubClassing
   Dim coordenadas_x() As String, coordenadas_y() As String, a As Integer
   GrdCoordenadas.TextMatrix(0, 0) = "Vazão"
   GrdCoordenadas.TextMatrix(0, 1) = "Carga"
   cboTipo.AddItem "Bomba"
   cboTipo.AddItem "Redimento"
   cboTipo.AddItem "Volume"
   cboTipo.AddItem "Perda de Carga"
   
   If ID = 0 Then
      NewID = True
      txtID.Enabled = True
   Else
      If Curves.Atualizar_Curva(ID) Then
         txtID = Curves.ID
         txtDescricao = Curves.DESCRICAO
         cboTipo.Text = Curves.tipo
         coordenadas_x = Split(Curves.COORDENADA_X, ";", 50)
         coordenadas_y = Split(Curves.COORDENADA_Y, ";", 50)
         For a = 0 To 49
            GrdCoordenadas.TextMatrix(a + 1, 0) = coordenadas_x(a)
            GrdCoordenadas.TextMatrix(a + 1, 1) = coordenadas_y(a)
         Next
      End If
      NewID = False
      txtID.Enabled = False
   End If
   Me.Show vbModal
   'LoozeXP1.EndWinXPCSubClassing
End Function

Private Sub cboTipo_Click()
   Select Case cboTipo.Text
      Case "Bomba"
         GrdCoordenadas.TextMatrix(0, 0) = "Vazão"
         GrdCoordenadas.TextMatrix(0, 1) = "Carga"
         RichTextBox1.Text = vbCrLf & "Informe os dados da curva" & vbCrLf & "Vazão (LPS) e" & vbCrLf & "Carga (m)"
      Case "Redimento"
         GrdCoordenadas.TextMatrix(0, 0) = "Vazão"
         GrdCoordenadas.TextMatrix(0, 1) = "Rendimento"
         RichTextBox1.Text = vbCrLf & "Informe os dados da curva" & vbCrLf & "Vazão (LPS) e" & vbCrLf & "Redimento (%)"
      Case "Volume"
         GrdCoordenadas.TextMatrix(0, 0) = "Vazão"
         GrdCoordenadas.TextMatrix(0, 1) = "Volume"
         RichTextBox1.Text = vbCrLf & "Informe os dados da curva" & vbCrLf & "Vazão (LPS) e" & vbCrLf & "Volume (m3)"
      Case "Perda de Carga"
         GrdCoordenadas.TextMatrix(0, 0) = "Vazão"
         GrdCoordenadas.TextMatrix(0, 1) = "Perda de Carga"
         RichTextBox1.Text = vbCrLf & "Informe os dados da curva" & vbCrLf & "Vazão (LPS) e" & vbCrLf & "Perda de Carga (m)"
   End Select
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdConfirmar_Click()
   Dim coordenadas_x As String, coordenadas_y As String, a As Integer
   For a = 1 To GrdCoordenadas.Rows - 1
      coordenadas_x = coordenadas_x & GrdCoordenadas.TextMatrix(a, 0) & ";"
      coordenadas_y = coordenadas_y & GrdCoordenadas.TextMatrix(a, 1) & ";"
   Next
   With Curves
      .ID = txtID
      .DESCRICAO = txtDescricao
      .tipo = cboTipo.Text
      .COORDENADA_X = coordenadas_x
      .COORDENADA_Y = coordenadas_y
      If NewID Then
         If .Inserir_Curva Then Sair_Form
      Else
         If .Atualizar_Curva(txtID.Text, True) Then Sair_Form
      End If
   
   End With
End Sub

Private Sub GrdCoordenadas_KeyPress(KeyAscii As Integer)
   With GrdCoordenadas
      If Iniciou_Entrada Then
         Select Case KeyAscii
           Case vbKeyDelete, vbKeyBack
               .TextMatrix(.Row, .Col) = ""
           Case vbKeyReturn
               If .Row < .Rows Then .Row = .Row + 1
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
               If .Row < .Rows Then .Row = .Row + 1
               
           Case Else
               If IsNumeric(.TextMatrix(.Row, .Col) & Chr(KeyAscii)) Then
                  .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) & Chr(KeyAscii)
               End If
         End Select
      End If
   End With
End Sub

Private Sub GrdCoordenadas_RowColChange()
   Iniciou_Entrada = False
End Sub

Private Sub Sair_Form()
   Set Curves = Nothing
   Unload Me
End Sub

