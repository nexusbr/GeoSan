VERSION 5.00
Object = "{9F81AD38-0759-4BFA-992A-04EA87455873}#1.0#0"; "LoozeXP.ocx"
Begin VB.Form frmPadroesCurvasSelecao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Navegador"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Looze.LoozeXP LoozeXP1 
      Left            =   2190
      Top             =   3570
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   405
      Left            =   1230
      Picture         =   "frmEPANetNavegador.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3270
      Width           =   435
   End
   Begin VB.CommandButton cmdEditar 
      Height          =   405
      Left            =   690
      Picture         =   "frmEPANetNavegador.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3270
      Width           =   435
   End
   Begin VB.CommandButton cmdAdcionar 
      Height          =   405
      Left            =   150
      Picture         =   "frmEPANetNavegador.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3270
      Width           =   435
   End
   Begin VB.ComboBox cboSelecao 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   180
      Width           =   2265
   End
   Begin VB.ListBox lstID 
      Height          =   2595
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   2235
   End
End
Attribute VB_Name = "frmPadroesCurvasSelecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Init() As Boolean
  LoozeXP1.InitIDESubClassing
  
  cboSelecao.AddItem "Curvas"
  cboSelecao.AddItem "Padrões"
  
  Me.Show vbModal
  LoozeXP1.EndWinXPCSubClassing
End Function

Private Sub cboSelecao_Click()
   Dim rs As ADODB.Recordset
   Dim curvas As New clsCurves
   lstID.Clear
   Select Case cboSelecao.Text
      Case "Curvas"
         If Not curvas.Retorna_Curvas(rs) Then
            Set rs = Nothing
            Exit Sub
         End If
      Case "Padrões"
   End Select
   While Not rs.EOF
      lstID.AddItem rs("id").value
      rs.MoveNext
   Wend
   lstID.ListIndex = 0
   
   rs.Close
   Set rs = Nothing
   Set curvas = Nothing
End Sub
