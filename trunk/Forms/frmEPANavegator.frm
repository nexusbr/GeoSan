VERSION 5.00
Begin VB.Form frmEPANavegator 
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
   Begin VB.CommandButton cmdExcluir 
      Height          =   405
      Left            =   1230
      Picture         =   "frmEPANavegator.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3270
      Width           =   435
   End
   Begin VB.CommandButton cmdEditar 
      Height          =   405
      Left            =   690
      Picture         =   "frmEPANavegator.frx":0410
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3270
      Width           =   435
   End
   Begin VB.CommandButton cmdAdcionar 
      Height          =   405
      Left            =   150
      Picture         =   "frmEPANavegator.frx":0C72
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3270
      Width           =   435
   End
   Begin VB.ComboBox cboSelecao 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   1
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
Attribute VB_Name = "frmEPANavegator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function init() As Boolean
  'LoozeXP1.InitIDESubClassing
  cboSelecao.AddItem "Curvas"
  cboSelecao.AddItem "Padrões"
  Me.Show vbModal
  'LoozeXP1.EndWinXPCSubClassing
End Function

Private Sub cboSelecao_Click()
   Dim rs As ADODB.Recordset
   Dim curvas As New clsEPACurves
   Dim padroes As New clsEPAPatterns
   lstID.Clear
   Select Case cboSelecao.Text
      Case "Curvas"
         If Not curvas.Retorna_Curvas(rs) Then
            Set rs = Nothing
            Exit Sub
         End If
      Case "Padrões"
         If Not padroes.Retorna_Padroes(rs) Then
            Set rs = Nothing
            Exit Sub
         End If
   End Select
   If Not rs Is Nothing Then
      While Not rs.EOF
         lstID.AddItem rs("id").value
         rs.MoveNext
      Wend
      lstID.ListIndex = 0
      
      rs.Close
   End If
   Set rs = Nothing
   Set curvas = Nothing
End Sub

Private Sub cmdAdcionar_Click()

   Select Case cboSelecao.Text
      Case "Curvas"
         Dim fCurves As New frmEPACurves
         fCurves.init 0
      Case "Padrões"
         Dim fPatterns As New frmEPAPatterns
         fPatterns.init 0
   End Select
   Set fCurves = Nothing
   Set fPatterns = Nothing
   cboSelecao_Click
End Sub

Private Sub cmdEditar_Click()
   Select Case cboSelecao.Text
      Case "Curvas"
         Dim fCurves As New frmEPACurves
         If lstID.Text <> "" Then fCurves.init lstID.Text
      Case "Padrões"
         Dim fPatterns As New frmEPAPatterns
         If lstID.Text <> "" Then fPatterns.init lstID.Text
   End Select
   Set fCurves = Nothing
   Set fPatterns = Nothing
   cboSelecao_Click
End Sub

Private Sub cmdExcluir_Click()
   Select Case cboSelecao.Text
      Case "Curvas"
         Dim Curves As New clsEPACurves
         If lstID.Text <> "" Then Curves.Excluir_Curva lstID
      Case "Padrões"
         Dim Patterns As New clsEPAPatterns
         If lstID.Text <> "" Then Patterns.Excluir_Padrao lstID
   End Select
   Set Curves = Nothing
   Set Patterns = Nothing
   cboSelecao_Click
End Sub

