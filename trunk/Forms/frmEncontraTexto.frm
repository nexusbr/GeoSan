VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEncontraTexto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar Textos"
   ClientHeight    =   4260
   ClientLeft      =   9645
   ClientTop       =   5790
   ClientWidth     =   5475
   Icon            =   "frmEncontraTexto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5475
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   180
      Width           =   4545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parte do texto"
      Height          =   720
      Left            =   90
      TabIndex        =   3
      Top             =   1170
      Width           =   3750
      Begin VB.OptionButton optQQRParte 
         Caption         =   "Qualquer parte"
         Height          =   300
         Left            =   2055
         TabIndex        =   6
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optFim 
         Caption         =   "Fim"
         Height          =   315
         Left            =   1185
         TabIndex        =   5
         Top             =   300
         Width           =   840
      End
      Begin VB.OptionButton optInicio 
         Caption         =   "Início"
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1995
      Left            =   90
      TabIndex        =   2
      Top             =   1950
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Localizado"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Eixo X"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Eixo Y"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Localizar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   4065
      TabIndex        =   1
      Top             =   1395
      Width           =   1005
   End
   Begin VB.TextBox TXTSTRING 
      Height          =   330
      Left            =   780
      TabIndex        =   0
      Top             =   660
      Width           =   4515
   End
   Begin VB.Label Label3 
      Height          =   225
      Left            =   180
      TabIndex        =   10
      Top             =   3990
      Width           =   3405
   End
   Begin VB.Label Label2 
      Caption         =   "Texto"
      Height          =   270
      Left            =   165
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Tema"
      Height          =   300
      Left            =   165
      TabIndex        =   8
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmEncontraTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim str As String
Dim rs As ADODB.Recordset
Dim strLayerID As String

Dim gu1 As String
Dim gu2 As String
Dim gu3 As String
Dim gu4 As String
Dim gu5 As String
Dim gu6 As String
Dim gu7 As String
Dim gu8 As String
Dim gu9 As String
 Dim str3, str2, str4 As String
Private Sub Combo1_Click()
    On Error GoTo Trata_Erro
    Dim rs As New ADODB.Recordset
    Dim Vetor As Variant
    Dim intTema As Integer
    
    Me.Lista.ListItems.Clear

      intTema = 0

    Open glo.diretorioGeoSan & "\CONTROLES\FTema.txt" For Input As #3   'LÊ O ARQUIVO LOG QUE FOI CRIADO NO MOMENTO DE ABERTURA DO MAPA
    Do While Not EOF(3)
        Line Input #3, str4
        Vetor = Split(str4, ";")
        If Vetor(1) = Combo1.Text Then
            intTema = Vetor(0)
            Exit Do
        End If
        'MsgBox vetor(0) & " É O NÚMERO THEME_ID QUE IDENTIFICA O LAYER E É FEITO O SELECT"
        'MsgBox vetor(1) & " É O NOME DO LAYER"
        ' vetor(2) 'É O COMANDO DO FILTRO
    Loop
    Close #3
    If frmCanvas.TipoConexao <> 4 Then
    str3 = "SELECT THEME_ID, LAYER_ID FROM TE_THEME WHERE THEME_ID =" & intTema & ""
    Else
    gu1 = "theme_id"
    gu2 = "layer_id"
    gu3 = "te_theme"
    gu4 = "geom_id"
    gu5 = "text_value"
    gu6 = "Texts"
    str3 = "SELECT " + """" + gu1 + """" + ", " + """" + gu2 + """" + " FROM " + """" + gu3 + """" + " WHERE " + """" + gu1 + """" + " ='" & intTema & "'"
    End If
     Set rs = Conn.execute(str3)
    ' DE ABERTURA DO MAPA
  
    
   
   
    If rs.EOF = False Then
        strLayerID = rs!layer_id
    End If
    rs.Close
 
    If frmCanvas.TipoConexao <> 4 Then
     str2 = "SELECT GEOM_ID,TEXT_VALUE FROM TEXTS" & strLayerID & " WHERE GEOM_ID = 0"
     Else
      gu1 = "theme_id"
    gu2 = "layer_id"
    gu3 = "te_theme"
    gu4 = "geom_id"
    gu5 = "text_value"
    gu6 = "Texts"
    gu7 = strLayerID
    gu8 = gu6 + gu7
     str2 = "SELECT " + """" + gu4 + """" + "," + """" + gu5 + """" + " FROM " + """" + "texts" + strLayerID + """" + " WHERE " + """" + gu4 + """" + " = '0'"
     End If
    
    
  
     Set rs = Conn.execute(str2)

     Me.cmdPesquisar.Enabled = True

     rs.Close

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
Else

    'MsgBox Err.Number & " " & Err.Description
    Err.Clear
    MsgBox "Não há texto na vista selecionada.", vbInformation
    Me.cmdPesquisar.Enabled = False
End If
End Sub
' Carrega os temas que estão ativos para o usuário
'
'
'
Private Sub Form_Load()
    On Error GoTo Trata_Erro
    Dim Vetor As Variant
    Dim str As String
    
    Close #3
    Open glo.diretorioGeoSan & "\CONTROLES\FTema.txt" For Input As #3 'LÊ O ARQUIVO LOG QUE FOI CRIADO NO MOMENTO DE ABERTURA DO MAPA
    Do While Not EOF(3)
        Line Input #3, str
        Vetor = Split(str, ";")
        Combo1.AddItem Vetor(1)
    Loop
    Close #3

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmEncontraTexto", "Form_Load", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Localiza os textos no mapa para poder fazer zoom
'
'
'
Private Sub cmdPesquisar_Click()
    On Error GoTo Trata_Erro
    Dim j As Long
    Dim itmx As ListItem
    Dim rs As New ADODB.Recordset

    Lista.ListItems.Clear
    gu1 = "theme_id"
    gu2 = "layer_id"
    gu3 = "te_theme"
    gu4 = "geom_id"
    gu5 = "text_value"
    gu6 = "texts"
    gu7 = "x"
    gu8 = "y"
    If Me.optInicio.value = True Then
        If frmCanvas.TipoConexao <> 4 Then
            str = "SELECT GEOM_ID,TEXT_VALUE,X,Y FROM TEXTS" & strLayerID & " WHERE TEXT_VALUE LIKE '" & TXTSTRING.Text & "%'"
        Else
            str = "SELECT " + """" + gu4 + """" + "," + """" + gu5 + """" + "," + """" + gu7 + """" + "," + """" + gu8 + """" + " FROM " + """" + gu6 + strLayerID + """" + " WHERE " + """" + gu5 + """" + " LIKE '" & TXTSTRING.Text & "%'"
        End If
    ElseIf Me.optFim.value = True Then
        If frmCanvas.TipoConexao <> 4 Then
            str = "SELECT GEOM_ID,TEXT_VALUE,X,Y FROM TEXTS" & strLayerID & " WHERE TEXT_VALUE LIKE '%" & TXTSTRING.Text & "'"
        Else
            str = "SELECT " + """" + gu4 + """" + "," + """" + gu5 + """" + "," + """" + gu7 + """" + "," + """" + gu8 + """" + " FROM " + """" + gu6 + strLayerID + """" + " WHERE " + """" + gu5 + """" + " LIKE '" & TXTSTRING.Text & "%'"
        End If
    ElseIf Me.optQQRParte.value = True Then
        If frmCanvas.TipoConexao <> 4 Then
            str = "SELECT GEOM_ID,TEXT_VALUE,X,Y FROM TEXTS" & strLayerID & " WHERE TEXT_VALUE LIKE '%" & TXTSTRING.Text & "%'"
        Else
            str = "SELECT " + """" + gu4 + """" + "," + """" + gu5 + """" + "," + """" + gu7 + """" + "," + """" + gu8 + """" + " FROM " + """" + gu6 + strLayerID + """" + " WHERE " + """" + gu5 + """" + " LIKE '" & TXTSTRING.Text & "%'"
        End If
    End If
    'FAZ SELECT COM BASE NOS CAMPOS CRIADOS
    j = 0
    If str <> "" Then
        Set rs = Conn.execute(str)
        If rs.EOF = False Then
            'CARREGA NO FORM TODAS AS LIGAÇÕES DISPONIVEIS COM BASE NO PRÉ FILTRO
            Do While Not rs.EOF
                'DoEvents
                Set itmx = Lista.ListItems.Add(, , rs.Fields("TEXT_VALUE").value)
                itmx.SubItems(1) = IIf(IsNull(rs.Fields("X").value), "", rs.Fields("X").value)
                itmx.SubItems(2) = IIf(IsNull(rs.Fields("Y").value), "", rs.Fields("Y").value)
                itmx.Tag = rs.Fields("GEOM_ID").value
                j = j + 1
                rs.MoveNext
            Loop
        End If
        rs.Close
        'Set Rs = Nothing
    End If
    Label3.Caption = "Localizadas " & j & " referências."
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmEncontraTexto", "cmdPesquisar_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Usuário clicou duas vezes no texto que deseja visualizar no mapa
'
'
'
Private Sub Lista_DblClick()
    On Error GoTo Trata_Erro
    Dim i As Long
    Dim X As Double, Y As Double
    Dim rs As New ADODB.Recordset

    gu1 = "theme_id"
    gu2 = "layer_id"
    gu3 = "te_theme"
    gu4 = "geom_id"
    gu5 = "text_value"
    gu6 = "texts"
    gu7 = "x"
    gu8 = "y"
    If strLayerID <> "" And Me.cmdPesquisar.Enabled = True Then
        If Lista.ListItems.count <= 0 Then
            Exit Sub
        End If
        i = Lista.SelectedItem.Tag
        If frmCanvas.TipoConexao <> 4 Then
            str = "SELECT GEOM_ID,TEXT_VALUE,X,Y FROM TEXTS" & strLayerID & " WHERE GEOM_ID =" & i & ""
        Else
            str = "SELECT " + """" + gu4 + """" + "," + """" + gu5 + """" + "," + """" + gu7 + """" + "," + """" + gu8 + """" + " FROM " + """" + gu6 + strLayerID + """" + " WHERE " + """" + gu4 + """" + "='" & i & "'"
        End If
        Set rs = Conn.execute(str)
        If rs.EOF = False Then
            xWorld = CLng(rs!X) 'carrega as variáveis públicas com valores do banco
            yWorld = CLng(rs!Y) 'carrega as variáveis públicas com valores do banco
        End If
        rs.Close
    End If
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "frmEncontraTexto", "Lista_DblClick", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

