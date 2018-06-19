VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadastroRamalFiltro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pesquisa"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7830
      TabIndex        =   9
      Top             =   3900
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   6630
      TabIndex        =   8
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione para a Pesquisa"
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   8895
      Begin VB.TextBox txtTextoFiltro 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   480
         Width           =   5415
      End
      Begin VB.ComboBox cboTipoFiltro 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "..."
         Height          =   285
         Left            =   8400
         TabIndex        =   3
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Filtro"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Digite a Informação"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Selecione os consumidores associados ao ramal"
      Height          =   2745
      Left            =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   8925
      Begin MSComctlLib.ListView lvLigacoes 
         Height          =   2415
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   4260
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NRO_LIGACAO"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CLASSIF. FISCAL"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ENDEREÇO"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CONSUMIDOR"
            Object.Width           =   5292
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCadastroRamalFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private frm As FrmCadastroRamal
Private tcs As TeCanvas
Private object_id_ramal As String

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdConfirmar_Click()
On Error GoTo Trata_Erro
   Dim a As Integer, itmx As ListItem
   For a = 1 To lvLigacoes.ListItems.count
      If lvLigacoes.ListItems(a).Checked Then
         With frm.lvLigacoes
            Set itmx = .ListItems.Add(, , lvLigacoes.ListItems(a).Text)
            itmx.SubItems(1) = lvLigacoes.ListItems(a).SubItems(1)
            itmx.SubItems(2) = lvLigacoes.ListItems(a).SubItems(2)
            itmx.SubItems(3) = lvLigacoes.ListItems(a).SubItems(3)
            itmx.Checked = True
            itmx.Tag = lvLigacoes.ListItems(a).Tag
         End With
      End If
   Next
   Unload Me
   'LoozeXP1.EndWinXPCSubClassing
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
    
       PrintErro CStr(Me.Name), "cmdConfirmar_Click()", CStr(Err.Number), CStr(Err.Description), True
    
    End If
End Sub

Private Sub cmdPesquisa_Click()
On Error GoTo Trata_Erro
   Dim rs As ADODB.Recordset, itmx As ListItem
   If txtTextoFiltro.Text = "" Then
      MsgBox "É necessario informar dados para pesquisa", vbExclamation
      Exit Sub
   End If
   lvLigacoes.ListItems.Clear
   Set rs = ConnSec.execute(SelecionarPesquisa)
      While Not rs.EOF
         
         With lvLigacoes
            Set itmx = .ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
            itmx.SubItems(1) = IIf(IsNull(rs.Fields("CLASSIFICACAO_FISCAL").value), "", rs.Fields("CLASSIFICACAO_FISCAL").value)
            itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
            itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)
            itmx.Tag = rs.Fields("codlograd").value
         End With
      rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Sub cmdPesquisa", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Sub

Sub init(m_frm As FrmCadastroRamal, m_tcs As TeCanvas, m_object_id_ramal)
On Error GoTo Trata_Erro
   Set frm = m_frm
   'LoozeXP1.InitIDESubClassing
   Set tcs = m_tcs
   object_id_ramal = m_object_id_ramal
   'CARREGA OPÇÕES PARA O FILTRO
   cboTipoFiltro.Clear
   With cboTipoFiltro
      .AddItem "Consumidor"
      .AddItem "Nº da Ligação"
      .AddItem "Logradouro"
      .AddItem "Classificação Fiscal"
      .ListIndex = 0
   End With
   Me.Show vbModal
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Sub Init", CStr(Err.Number), CStr(Err.Description), True
      
    End If
End Sub

Private Function SelecionarPesquisa()
On Error GoTo Trata_Erro
   Dim str As String
   Select Case cboTipoFiltro.Text
      Case "Logradouro"
         str = GetQueryProcess(6)
         str = Replace(str, "@LOGRADOURO", UCase(txtTextoFiltro.Text))
      Case "Consumidor"
         str = GetQueryProcess(7)
'MsgBox str
         str = Replace(str, "@USUARIO", UCase(txtTextoFiltro.Text))
'MsgBox str
      Case "Nº da Ligação"
         str = GetQueryProcess(8)
         str = Replace(str, "@NRO_LIGACAO", txtTextoFiltro.Text)
      Case "Classificação Fiscal"
         str = GetQueryProcess(9)
         str = Replace(str, "@CLASSIFICACAO_FISCAL", txtTextoFiltro.Text)
   End Select
   SelecionarPesquisa = str
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Function SelecionarPesquisa", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Function


Private Function Verifica_Ligacao(index As Integer) As Boolean
On Error GoTo Trata_Erro
   Dim a As Integer, UltimoEndereco As String, rs As ADODB.Recordset, str As String, b As Integer
   For a = 1 To lvLigacoes.ListItems.count
      If lvLigacoes.ListItems(a).Checked Then
         If UltimoEndereco <> "" Then
            'VERIFICA SE TODOS OS LOGRADOUROS SAO O MESMO
            If lvLigacoes.ListItems(a).Tag <> UltimoEndereco Then
               MsgBox "Não é possível vincular em um mesmo ramal ligações de logradouros diferentes", vbExclamation
               Exit Function
            End If
            
         End If
         UltimoEndereco = lvLigacoes.ListItems(a).Tag
      End If
   Next
   
   For a = 1 To lvLigacoes.ListItems.count
      If lvLigacoes.ListItems(a).Checked Then
         For b = 1 To frm.lvLigacoes.ListItems.count
            If frm.lvLigacoes.ListItems(b).Checked And frm.lvLigacoes.ListItems(b).Tag <> lvLigacoes.ListItems(a).Tag Then
               'VERIFICA SE TODOS OS LOGRADOUROS SAO O MESMO
               MsgBox "Não é possível vincular em um mesmo ramal ligações de logradouros diferentes", vbExclamation
               Exit Function
            End If
         Next
      End If
   Next
   
   'VEVIFICA SE A LIGAÇÃO JÁ ESTÁ VINCULADA EM OUTRO RAMAL
   str = GetQueryProcess(10)
   str = Replace(str, "@LAYER", tcs.getCurrentLayer)
   str = Replace(str, "@OBJECT_ID_RAMAL", object_id_ramal)
   str = Replace(str, "@NRO_LIGACAO", lvLigacoes.ListItems(index).Text)
   Set rs = Conn.execute(str)
   If Not rs.EOF Then
      MsgBox "Esta ligação, embora esteja vinculada a este lote, já está vinculada a outro ramal:" & rs(0).value, vbExclamation
      Exit Function
   End If
   rs.Close
   Set rs = Nothing
   Verifica_Ligacao = True

Trata_Erro:
    
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       PrintErro CStr(Me.Name), "Private Function Verifica_Ligacao", CStr(Err.Number), CStr(Err.Description), True
    End If

End Function

Private Sub lvLigacoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Trata_Erro
   
   If Item.Checked Then
      If Not Verifica_Ligacao(Item.index) Then
         Item.Checked = False
      End If
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      
      PrintErro CStr(Me.Name), "lvLigacoes_ItemCheck", CStr(Err.Number), CStr(Err.Description), True
    
   End If
End Sub

