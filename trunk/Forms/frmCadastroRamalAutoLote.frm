VERSION 5.00
Begin VB.Form frmCadastroRamalAutoLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Campos de Pesquisa por Lotes"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTabelaAtributos 
      Height          =   315
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1455
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   3705
      TabIndex        =   7
      Top             =   3900
      Width           =   975
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   405
      Left            =   4725
      TabIndex        =   6
      Top             =   3900
      Width           =   975
   End
   Begin VB.ComboBox cboCampoNroLigacao 
      Height          =   315
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   2325
   End
   Begin VB.ComboBox cboCampoIPTU 
      Height          =   315
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2250
      Width           =   2325
   End
   Begin VB.ComboBox cboTemaLotes 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   2325
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Selecione a tabela de Atributos do plano de Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   9
      Top             =   1140
      Width           =   4470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selecione a coluna referente a informação Número Ligação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   2805
      Width           =   5325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Selecione a coluna referente a informação IPTU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   4
      Top             =   1920
      Width           =   4290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o tema referente ao plano de Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   315
      Width           =   4080
   End
End
Attribute VB_Name = "frmCadastroRamalAutoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim contVetor As Integer
Dim rs As ADODB.Recordset
Dim blnIPTU As Boolean
Dim TbAtributos As String
Dim ve As String
         Dim vi As String
         Dim vo As String
         Dim vu As String
         Dim vc As String
          Dim vd As String
          Dim vm As String
          Dim vf As String




Private Sub cmdCancelar_Click()
   Unload Me
 
End Sub

Private Sub cmdSalvar_Click()

   'GRAVA NO ARQUIVO INI AS CONFIGURAÇÕES DE PESQUISA
   
   Call WriteINI("RAMAISFILTROLOTES", "TABELA_PLANO", Me.cboTemaLotes.Text, App.path & "\CONTROLES\GEOSAN.INI")
   Call WriteINI("RAMAISFILTROLOTES", "TABELA_ATRIB", Me.cboTabelaAtributos.Text, App.path & "\CONTROLES\GEOSAN.INI")
   Call WriteINI("RAMAISFILTROLOTES", "REF_IPTU", Me.cboCampoIPTU.Text, App.path & "\CONTROLES\GEOSAN.INI")
   Call WriteINI("RAMAISFILTROLOTES", "REF_NROLIGACAO", Me.cboCampoNroLigacao.Text, App.path & "\CONTROLES\GEOSAN.INI")


   Unload Me

End Sub

Private Sub Form_Load()
    
   Dim str As String
   
   Close #3
   
   Open glo.diretorioGeoSan + "\CONTROLES\FTema.txt" For Input As #3    'LÊ O ARQUIVO LOG QUE FOI CRIADO NO MOMENTO DE ABERTURA DO MAPA
   Do While Not EOF(3)
      Line Input #3, str
      Vetor = Split(str, ";")
      Me.cboTemaLotes.AddItem Vetor(1)
   Loop
   Close #3
    
   Set rs = New ADODB.Recordset
   va = "TABLES"
         ve = "OBJECT_ID_"
         vi = "NRO_LIGACAO"
         vo = "INSCRICAO_LOTE"
         vu = "TIPO"
         vc = "HIDROMETRADO"
         vd = "NXGS_V_LIG_COMERCIAL"
         ve = "CONSUMO_LPS"
         vf = "ECONOMIAS"
         
             If frmCanvas.TipoConexao = 1 Then
   str = "SELECT NAME AS TABELA FROM SYS.TABLES"
    
   ElseIf frmCanvas.TipoConexao = 2 Then
      str = "SELECT DISTINCT TABLE_NAME AS " + """" + "TABELA" + """" + " FROM ALL_TAB_COLS"
   
  
   ElseIf frmCanvas.TipoConexao = 4 Then
    Dim ad As String
   Dim ae As String
   ad = "NAME"
   ae = "pg_tables"
     str = "SELECT" + """" + "tablename" + """" + "As" + """" + "TABELA" + """" + " FROM " + """" + ae + """"
     End If
     
      rs.Open str, Conn, adOpenDynamic, adLockOptimistic
   
   Do While Not rs.EOF
      Me.cboTabelaAtributos.AddItem rs!TABELA
      rs.MoveNext
   Loop
   rs.Close
    
   Me.cboTemaLotes.Text = ReadINI("RAMAISFILTROLOTES", "TABELA_PLANO", App.path & "\CONTROLES\GEOSAN.INI")
   Me.cboTabelaAtributos.Text = ReadINI("RAMAISFILTROLOTES", "TABELA_ATRIB", App.path & "\CONTROLES\GEOSAN.INI")
   Me.cboCampoIPTU.Text = ReadINI("RAMAISFILTROLOTES", "REF_IPTU", App.path & "\CONTROLES\GEOSAN.INI")
   Me.cboCampoNroLigacao.Text = ReadINI("RAMAISFILTROLOTES", "REF_NROLIGACAO", App.path & "\CONTROLES\GEOSAN.INI")

    
End Sub
Private Sub ValidaTema()
   'CAPTURAR O NOME DE TODAS AS COLUNAS DA TABELA
   
   Dim strsql As String
   Dim codTema As Integer
   Dim str As String
   
   Me.cboCampoIPTU.Clear
   Me.cboCampoNroLigacao.Clear

   Set rs = New ADODB.Recordset
   
   'PROCURAR NO VETOR O ID DO TEMA SELECIONADO
   Close #3
   Open glo.diretorioGeoSan + "\GEOSAN\CONTROLES\FTema.txt" For Input As #3     'LÊ O ARQUIVO LOG QUE FOI CRIADO NO MOMENTO DE ABERTURA DO MAPA
   Do While Not EOF(3)
       Line Input #3, str
       Vetor = Split(str, ";")
       If CStr(Vetor(1)) = CStr(Me.cboTemaLotes.Text) Then
           codTema = Vetor(0)
           Exit Do
       End If
   Loop
   Close #3
   
   Dim LayNome As String
   If frmCanvas.TipoConexao <> 4 Then
   
   'VERIFICA SE O TEMA SELECIONADO POSSUI A GEOMETRIA DE POLIGONOS
   strsql = "SELECT * FROM TE_REPRESENTATION WHERE GEOM_TYPE = 1 AND LAYER_ID = (SELECT LAYER_ID FROM TE_THEME WHERE THEME_ID = " & codTema & ")"
   Set rs = New ADODB.Recordset
   rs.Open strsql, Conn, adOpenForwardOnly, adLockReadOnly
   
   If rs.EOF = True Then
      'NÃO FOI LOCALIZADO A GEOMETRIA DE POLÍGONOS NO PLANO SELECIONADO
      MsgBox "O tema selecionado não possui polígonos.", vbInformation, ""
      rs.Close
      'Exit Sub
   End If
   Else
   Dim ff As String
   Dim fd As String
   Dim fc As String
   Dim fb As String
   Dim fg As String
   ff = "te_representation"
   fd = "geom_type"
   fc = "layer_id"
   fb = "te_theme"
   fg = "theme_id"
   
   
   strsql = "SELECT * FROM " + """" + ff + """" + " WHERE " + """" + fd + """" + " = '1' AND " + """" + fc + """" + " = (SELECT " + """" + fc + """" + " FROM " + """" + fb + """" + " WHERE " + """" + fg + """" + " = '" & codTema & "')"
   End If
   Set rs = New ADODB.Recordset
    rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   
   If rs.EOF = True Then
      'NÃO FOI LOCALIZADO A GEOMETRIA DE POLÍGONOS NO PLANO SELECIONADO
      MsgBox "O tema selecionado não possui polígonos.", vbInformation, ""
      rs.Close
      'Exit Sub
   End If
   
End Sub


Private Sub CARREGA_COMBOS()


   
'   'LOCALIZA A TABELA DE ATRIBUTOS DO PLANO
'   strsql = "SELECT ATTR_TABLE FROM TE_LAYER_TABLE WHERE LAYER_ID = (SELECT LAYER_ID FROM TE_THEME WHERE THEME_ID = " & codTema & ")"
'   Set RS = New ADODB.Recordset
'   RS.Open strsql, Conn, adOpenForwardOnly, adLockReadOnly
'
'   If RS.EOF = False Then
'      TbAtributos = RS!ATTR_TABLE
'   Else
'      MsgBox "Não foi localizada a tabela de atributos do tema. (TE_LAYER_TABLE) " & Chr(13) & Chr(13) & "Será utilizado os campos da própria tabela.", vbInformation, ""
'      TbAtributos = Me.cboTemaLotes.Text
'      RS.Close
'      'Exit Sub
'   End If

   
   'CARREGA NOS COMBOS IPTU E NRO_LIGACAO OS NOMES DAS COLUNAS DA TABELA DE ATRIBUTOS
   va = "RAMAIS_AGUA_LIGACAO"
         ve = "OBJECT_ID_"
         vi = "NRO_LIGACAO"
         vo = "INSCRICAO_LOTE"
         vu = "TIPO"
         vc = "HIDROMETRADO"
         vd = "NXGS_V_LIG_COMERCIAL"
         vm = "CONSUMO_LPS"
         vf = "ECONOMIAS"
         
             If frmCanvas.TipoConexao <> 4 Then
   strsql = "SELECT * FROM " & Me.cboTabelaAtributos.Text
   Else
      strsql = "SELECT * FROM " + """" + Me.cboTabelaAtributos.Text + """"
   
   End If
   
   
   Set rs = New ADODB.Recordset
     rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   
   Me.cboCampoIPTU.Clear
   Me.cboCampoNroLigacao.Clear

   
  ' If rs.EOF = False Then
     
      For i = 0 To rs.Fields.count - 1
         Me.cboCampoIPTU.AddItem rs.Fields(i).Name 'NOME DA COLUNA
         Me.cboCampoNroLigacao.AddItem rs.Fields(i).Name 'NOME DA COLUNA
      Next
   
  ' End If
   
   rs.Close

End Sub

Private Sub cboTemaLotes_Click()
   If Me.cboTemaLotes.Text <> "" Then
      ValidaTema
   End If
End Sub

'Private Sub cboTemaLotes_LostFocus()
'   If Me.cboTemaLotes.Text <> "" Then
'      ValidaTema
'   End If
'End Sub
'
'Private Sub cboTabelaAtributos_LostFocus()
'   If Me.cboTabelaAtributos.Text <> "" Then
'      CARREGA_COMBOS
'   End If
'End Sub

Private Sub cboTabelaAtributos_click()
   If Me.cboTabelaAtributos.Text <> "" Then
      CARREGA_COMBOS
   End If
End Sub

Private Sub cboCampoIPTU_Click()
   
   If VerificaSeNumerico(Me.cboCampoIPTU.Text, Me.cboTabelaAtributos.Text) = True Then
      blnIPTU = True
   Else
      blnIPTU = False
      Me.cmdSalvar.Enabled = False
      MsgBox "O campo selecionado possui registro não numérico ou vazio e não pode ser utilizado.", vbInformation, ""
   End If
   
End Sub

Private Sub cboCampoNroLigacao_Click()
  
   If VerificaSeNumerico(Me.cboCampoNroLigacao.Text, Me.cboTabelaAtributos.Text) = True Then
      If blnIPTU = True Then
         Me.cmdSalvar.Enabled = True
      End If
   Else
      Me.cmdSalvar.Enabled = False
      MsgBox "O campo selecionado possui registro não numérico ou vazio e não pode ser utilizado.", vbInformation, ""
   End If

End Sub


