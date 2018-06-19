VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{91488A85-7250-4842-8681-87818334B791}#1.0#0"; "NxViewManager2.ocx"
Begin VB.Form FrmPesquisaCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pesquisa"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin NxViewManager.ViewManager ViewManager1 
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione para a Pesquisa"
      Height          =   975
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   8895
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "..."
         Height          =   285
         Left            =   8400
         TabIndex        =   5
         Top             =   480
         Width           =   345
      End
      Begin VB.ComboBox cboTipoFiltro 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   3525
      End
      Begin VB.TextBox txtTextoFiltro 
         Height          =   285
         Left            =   3780
         TabIndex        =   3
         Top             =   480
         Width           =   4515
      End
      Begin VB.Label Label2 
         Caption         =   "Digite a Informação"
         Height          =   255
         Left            =   3780
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Filtro"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   345
      Left            =   7770
      TabIndex        =   1
      Top             =   4560
      Width           =   1155
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   1110
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmPesquisaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tcs As TeCanvas
Private cgeo As New clsGeoReference
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String
Dim m As String


Public Function init(mtcs As TeCanvas) As Boolean
   
   'LoozeXP1.InitIDESubClassing
   
   
   With cboTipoFiltro
      .AddItem "USUÁRIO - LOTES"
      .AddItem "LOGRADOURO - LOTES"
      .AddItem "CLASSIF. FISCAL - LOTES"
      .AddItem "LIGACÕES PENDENTES - LOTES"
      .AddItem "RAMAL ÁGUA"
      .AddItem "LIGAÇÕES - RAMAL ÁGUA"
      .AddItem "RAMAL ESGOTO"
      .AddItem "LIGAÇÕES - RAMAL ESGOTO"
      .AddItem "LOGRADOURO - NÓS DA REDE DE ÁGUA"
      .ListIndex = 1
   End With
   txtTextoFiltro.Text = ""
   
   Set tcs = mtcs
   
   Me.Show vbModal
   
   'LoozeXP1.EndWinXPCSubClassing
   
End Function


Private Sub cmdPesquisa_Click()
   On Error GoTo cmdPesquisa_Click
   Dim rs As ADODB.Recordset, str As String, rsA As ADODB.Recordset, cgeo As New clsGeoReference, LayerName As String
   Dim a As Integer, itmx As ListItem
   
   Screen.MousePointer = vbHourglass
   With FrmMain
   Select Case cboTipoFiltro.Text
      Case "USUÁRIO - LOTES"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES)) Then
            str = GetQueryProcess(7)
            str = Replace(str, "@USUARIO", UCase(txtTextoFiltro.Text))
            Set rs = ConnSec.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES), vbExclamation
            Exit Sub
         End If
      Case "LOGRADOURO - LOTES"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES)) Then
         
            str = GetQueryProcess(6)
            str = Replace(str, "@LOGRADOURO", UCase(txtTextoFiltro.Text))
            Set rs = ConnSec.execute(str)
         Else
            Screen.MousePointer = vbNormal
            
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES), vbExclamation
            Exit Sub
         End If
      Case "LOGRADOURO - NÓS DA REDE DE ÁGUA"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.Componente_Rede_Agua)) Then
            str = GetQueryProcess(6)
            str = Replace(str, "@LOGRADOURO", UCase(txtTextoFiltro.Text))
            Set rs = ConnSec.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.Componente_Rede_Agua), vbExclamation
            Exit Sub
         End If
      Case "CLASSIF. FISCAL - LOTES"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES)) Then
            str = GetQueryProcess(9)
            str = Replace(str, "@CLASSIFICACAO_FISCAL", txtTextoFiltro.Text)
            Set rs = ConnSec.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES), vbExclamation
            Exit Sub
         End If
      Case "LIGACÕES PENDENTES - LOTES"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES)) Then
            Set rsA = New ADODB.Recordset
            rsA.CursorType = adOpenDynamic
            
a = "NRO_LIGACAO"
b = cgeo.GetLayerOperation(tcs.getCurrentLayer, IIf(cgeo.GetLayerTypeReference(tcs.getCurrentLayer) = RAMAIS_AGUA, 1, 2))
c = b
d = "TIPO"
e = "HIDROMETRADO"
f = "ECONOMIAS"
g = "CONSUMO_LPS"
h = "TB_LIGACOES"
i = "HIDROMETRADO"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"
m = "_LIGACAO"


     If frmCanvas.TipoConexao <> 4 Then
            Set rsA = Conn.execute("SELECT NRO_LIGACAO FROM " & cgeo.GetLayerOperation(tcs.getCurrentLayer, IIf(cgeo.GetLayerTypeReference(tcs.getCurrentLayer) = RAMAIS_AGUA, 1, 2)) & "_LIGACAO")
            Else
             Set rsA = Conn.execute("SELECT " + """" + a + """" + " FROM  + """" + c + m+ """" +")
            End If
            Set rs = ConnSec.execute(GetQueryProcess(17))
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.LOTES), vbExclamation
            Exit Sub
         End If
      Case "RAMAL ÁGUA"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_AGUA)) Then
            str = GetQueryProcess(12)
            str = Replace(str, "@OBJECT_ID_", txtTextoFiltro.Text)
            str = Replace(str, "@LAYER", tcs.getCurrentLayer)
            Set rs = Conn.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_AGUA), vbExclamation
            Exit Sub
         End If
      Case "RAMAL ESGOTO"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_ESGOTO)) Then
            str = GetQueryProcess(12)
            str = Replace(str, "@OBJECT_ID_", txtTextoFiltro.Text)
            str = Replace(str, "@LAYER", tcs.getCurrentLayer)
            Set rs = Conn.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_ESGOTO), vbExclamation
            Exit Sub
         End If
      Case "LIGAÇÕES - RAMAL ÁGUA"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_AGUA)) Then
            str = GetQueryProcess(13)
            str = Replace(str, "@NRO_LIGACAO", txtTextoFiltro.Text)
            str = Replace(str, "@LAYER", tcs.getCurrentLayer)
            Set rs = Conn.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_AGUA), vbExclamation
            Exit Sub
         End If
      Case "LIGAÇÕES - RAMAL ESGOTO"
         If .ViewManager1.TvSetCurrentLayer(cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_ESGOTO)) Then
         str = GetQueryProcess(13)
         str = Replace(str, "@NRO_LIGACAO", txtTextoFiltro.Text)
         str = Replace(str, "@LAYER", tcs.getCurrentLayer)
         Set rs = Conn.execute(str)
         Else
            Screen.MousePointer = vbNormal
            MsgBox "Não existe um tema relacionado ao plano: " & cgeo.GetLayerNameByTypeReference(LayerTypeRefence.RAMAIS_ESGOTO), vbExclamation
            Exit Sub
         End If
   End Select
   End With
   
   lv.ListItems.Clear
   lv.ColumnHeaders.Clear
   
   For a = 0 To rs.Fields.count - 1
      lv.ColumnHeaders.Add , , UCase(rs.Fields(a).Name)
   Next
   If rsA Is Nothing Then Set rsA = New ADODB.Recordset
   
   While Not rs.EOF
      If rsA.State = 1 Then
         rsA.Filter = "NRO_LIGACAO='" & rs.Fields("NRO_LIGACAO").value & "'"
         If rsA.EOF Then
            Set itmx = lv.ListItems.Add(, , IIf(IsNull(rs(0).value), "", UCase(rs(0).value)))
            For a = 1 To rs.Fields.count - 1
               itmx.SubItems(a) = IIf(IsNull(rs(a).value), "", UCase(rs(a).value))
            Next
            itmx.Tag = Left(rs.Fields(rs.Fields.count - 1).value, 11)
         End If
      Else
         Set itmx = lv.ListItems.Add(, , IIf(IsNull(rs(0).value), "", UCase(rs(0).value)))
         For a = 1 To rs.Fields.count - 1
            itmx.SubItems(a) = IIf(IsNull(rs(a).value), "", UCase(rs(a).value))
         Next
         itmx.Tag = Left(rs.Fields(rs.Fields.count - 1).value, 11)
      End If
      rs.MoveNext
   Wend
   If rsA.State = 1 Then rs.Close
   Set rsA = Nothing
   If rs.State = 1 Then rs.Close
   Set rs = Nothing
   Screen.MousePointer = vbNormal
   Exit Sub
cmdPesquisa_Click:
   MsgBox "Ocorreu um erro no sistema, é possível que você não esteja no plano correto" & vbCrLf & Err.Description
   Screen.MousePointer = vbNormal
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
   On Error GoTo lv_ItemClick_err
   Dim rs As ADODB.Recordset, object_id As String, xmin As Double, ymin As Double, xmax As Double, ymax As Double
   Dim str As String
   Select Case cboTipoFiltro
      Case "USUÁRIO - LOTES", "LOGRADOURO - LOTES", "CLASSIF. FISCAL - LOTES", "LIGACÕES PENDENTES - LOTES"
         str = GetQueryProcess(14)
         str = Replace(str, "@CLASSIFICACAO_FISCAL", Item.SubItems(1))
      Case "RAMAL ÁGUA", "LIGAÇÕES - RAMAL ÁGUA"
         If cgeo.GetLayerTypeReference(tcs.getCurrentLayer) <> RAMAIS_AGUA Then
            MsgBox "Selecione o plano/tema referente ao Ramal de Água", vbExclamation
         End If
         str = GetQueryProcess(16)
         str = Replace(str, "@OBJECT_ID_", Item.Text)
         str = Replace(str, "@LAYER", tcs.getCurrentLayer)
      Case "RAMAL ESGOTO", "LIGAÇÕES - RAMAL ESGOTO"
         If cgeo.GetLayerTypeReference(tcs.getCurrentLayer) <> RAMAIS_ESGOTO Then
            MsgBox "Selecione o plano/tema referente ao Ramal de Esgoto", vbExclamation
            Exit Sub
         End If
         str = GetQueryProcess(16)
         str = Replace(str, "@CLASSIFICACAO_FISCAL", Item.Text)
         str = Replace(str, "@LAYER", tcs.getCurrentLayer)
      Case "LOGRADOURO - NÓS DA REDE DE ÁGUA"
         str = GetQueryProcess(20)
         str = Replace(str, "@CLASSIFICACAO_FISCAL", Item.SubItems(1))
      
         
   End Select
      
   With tcs
      .Normal
      Set rs = Conn.execute(str)
      If Not rs.EOF Then
         While Not rs.EOF
            object_id = IIf(IsNull(rs!Object_id_), "", rs!Object_id_)
            'If object_id <> "" Then
               
               If .addSelectObjectIds(object_id) = 1 Then
                  .getSelectBox xmin, ymin, xmax, ymax
                  .setWorld xmin - 1000, ymin - 1000, xmax + 1000, ymax + 1000
                  .Select
                  .setScale 1000
               Else
                  MsgBox "Não foi encontrado a geometria referente ao atributo selecionado", vbExclamation
               End If
            'Else
            '   MsgBox "Objecto não encontrado", vbExclamation
            'End If
            rs.MoveNext
         Wend
      Else
         MsgBox "Número da inscrição não encontrado", vbExclamation
      
      End If
      rs.Close
      Set rs = Nothing
   End With
   Exit Sub
lv_ItemClick_err:
   MsgBox "Ocorreu um erro no sistema, é possível que você não esteja no plano correto" & vbCrLf & Err.Description
End Sub


'            StrSql = "SELECT "
'            StrSql = StrSql & " TL.tlogradnome AS TIPOVIA,"
'            StrSql = StrSql & " L.logradnome AS LOGRADOURO,"
'            StrSql = StrSql & " B.bairnome AS BAIRRO,"
'            StrSql = StrSql & " I.imoburbcomplemento AS COMPLEMENTO,"
'            StrSql = StrSql & " S.sannumerohidrometro AS NRO_HIDROMETRO,"
'            StrSql = StrSql & " I.imoburbinscricao AS NRO_INSCRICAO"
'            StrSql = StrSql & " FROM imobiliario_urbano I"
'            StrSql = StrSql & " LEFT JOIN saneamento_imobiliario_urbano S ON S.imoburbcod=I.imoburbcod"
'            StrSql = StrSql & " LEFT JOIN bairro B ON B.baircod=I.baircod"
'            StrSql = StrSql & " LEFT JOIN logradouro L ON L.logradcod = I.logradcod"
'            StrSql = StrSql & " LEFT JOIN tipo_logradouro TL ON TL.tlogradabreviatura = L.tlogradabreviatura"
'            StrSql = StrSql & " WHERE  L.logradnome LIKE '" & UCase(txtPesquisa.Text) & "%'"

'            StrSql = "SELECT CO.contribnome AS CONTRIBUINTE,"
'            StrSql = StrSql & " TL.tlogradnome AS TIPOVIA,"
'            StrSql = StrSql & " L.logradnome AS LOGRADOURO,"
'            StrSql = StrSql & " i.imoburbnumero as NUMERO,"
'            StrSql = StrSql & " B.bairnome AS BAIRRO,"
'            StrSql = StrSql & " I.imoburbcomplemento AS COMPLEMENTO,"
'            StrSql = StrSql & " S.sannumerohidrometro AS NRO_HIDROMETRO,"
'            StrSql = StrSql & " I.imoburbinscricao AS NRO_INSCRICAO"
'            StrSql = StrSql & " FROM contribuinte CO"
'            StrSql = StrSql & " INNER JOIN  socio_imobiliario_urbano SO ON SO.contribcodsocio=CO.contribcod"
'            StrSql = StrSql & " LEFT JOIN imobiliario_urbano I  ON I.imoburbcod=SO.imoburbcod"
'            StrSql = StrSql & " LEFT JOIN saneamento_imobiliario_urbano S ON S.imoburbcod=I.imoburbcod"
'            StrSql = StrSql & " LEFT JOIN bairro B ON B.baircod=I.baircod"
'            StrSql = StrSql & " LEFT JOIN logradouro L ON L.logradcod = I.logradcod"
'            StrSql = StrSql & " LEFT JOIN tipo_logradouro TL ON TL.tlogradabreviatura = L.tlogradabreviatura"
'            StrSql = StrSql & " WHERE  CO.contribnome LIKE '" & UCase(txtPesquisa.Text) & "%'"
'            StrSql = StrSql & " ORDER BY CO.contribnome,L.logradnome,i.imoburbnumero"


'            StrSql = "SELECT "
'            StrSql = StrSql & " TL.tlogradnome AS TIPOVIA,"
'            StrSql = StrSql & " L.logradnome AS LOGRADOURO,"
'            StrSql = StrSql & " B.bairnome AS BAIRRO,"
'            StrSql = StrSql & " I.imoburbcomplemento AS COMPLEMENTO,"
'            StrSql = StrSql & " S.sannumerohidrometro AS NRO_HIDROMETRO,"
'            StrSql = StrSql & " I.imoburbinscricao AS NRO_INSCRICAO"
'            StrSql = StrSql & " FROM imobiliario_urbano I"
'            StrSql = StrSql & " LEFT JOIN saneamento_imobiliario_urbano S ON S.imoburbcod=I.imoburbcod"
'            StrSql = StrSql & " LEFT JOIN bairro B ON B.baircod=I.baircod"
'            StrSql = StrSql & " LEFT JOIN logradouro L ON L.logradcod = I.logradcod"
'            StrSql = StrSql & " LEFT JOIN tipo_logradouro TL ON TL.tlogradabreviatura = L.tlogradabreviatura"
'            StrSql = StrSql & " WHERE  I.imoburbinscricao LIKE '" & UCase(txtPesquisa.Text) & "%'"

Private Sub txtTextoFiltro_Change()

End Sub

