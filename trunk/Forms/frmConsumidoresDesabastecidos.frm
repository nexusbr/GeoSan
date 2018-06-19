VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsumidoresDesabastecidos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ligações cadastras nos trechos selecionados"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   5520
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGerartxt 
      Caption         =   "Gerar txt"
      Height          =   315
      Left            =   6390
      TabIndex        =   4
      Top             =   4860
      Width           =   1215
   End
   Begin VB.TextBox txtQde 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   4830
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   7680
      TabIndex        =   1
      Top             =   4860
      Width           =   1155
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   8229
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro Ligação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Endereço"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usuário"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tel. Res."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tel. Com."
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Total de Ligações:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   4890
      Width           =   1635
   End
End
Attribute VB_Name = "frmConsumidoresDesabastecidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conexao As New ADODB.connection
Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
Dim nStr As String
Dim strConn As String
Dim connection As Integer
'Função para apresentar os usuários afetados em uma manobra de rede
'
'init -
'Object_id_trecho - string com todos os object_id_s dos trechos de rede em que serão consultados os ramais pertencentes aos mesmos
'
Public Function init(Object_id_trecho As String) As Boolean
    On Error GoTo Trata_Erro
    'LoozeXP1.InitIDESubClassing
    Dim TABELACOMERCIAL As String
    Dim count2 As Integer
    Dim lig As String
    Dim str As String
    Dim rs As ADODB.Recordset               'todas as ligações que estão localizadas nos trechos que estão conectados até encontrar uma válvula, os quais os object_id_s já foram fornecidos na estrada desta função
    Dim itmx As ListItem                    'lista com os dados dos consumidores que serão afetados pela manobra
    Dim QtdeLig As Integer                  'quantidade total de consumidores afetados pela manobra
    Dim rs2 As ADODB.Recordset              'da tabela que possui os dados temporários das ligações
    Dim fg As String
    Dim fh As String
    Dim fi As String
    Dim fk As String
    Dim fl As String
    Dim fm As String
    Dim numeroErro As String                'indica o número do erro que pode ocorrer
    Dim ddd As String
    Dim rsNro_Ligacao As ADODB.Recordset
    
    numeroErro = "Erro não identificado"
    count2 = 0
    fg = "RAMAIS_AGUA"
    fh = "RAMAIS_AGUA_LIGACAO"
    fi = "NRO_LIGACAO"
    fk = "ramais_agua"
    fl = "OBJECT_ID_"
    fm = "OBJECT_ID_TRECHO"
    If frmCanvas.TipoConexao = 4 Then
        'se for Postgres
        If connection <> 10 Then
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            nStr = frmCanvas.FunDecripta(Senha)
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
            conexao.Open strConn
            connection = 10
        End If
    End If
    'prepara consulta de todas as ligações que estão localizadas nos trechos que estão conectados até encontrar uma válvula, os quais os object_id_s já foram fornecidos na estrada desta função
    If frmCanvas.TipoConexao <> 4 Then
        'se for Oracle ou SQLServer
        str = "SELECT ramais_agua_ligacao.nro_ligacao from ramais_agua_ligacao inner join ramais_agua on ramais_agua.object_id_=ramais_agua_ligacao.object_id_ where object_id_trecho in(" & Object_id_trecho & ")"
        Set rs = Conn.execute(str)
    Else
        str = "SELECT " + """" + fh + """" + "." + """" + fi + """" + " from " + """" + fh + """" + "  inner join " + """" + fg + """" + "  on " + """" + fg + """" + "." + """" + fl + """" + "=" + """" + fh + """" + "." + """" + fl + """" + " where " + """" + fm + """" + " in(" & Object_id_trecho & ")"
        Set rs = conexao.execute(str)
    End If
    'While Not rs.EOF
     '  lig = rs.Fields(0).value
    
      ' rs.MoveNext
    'Wend
    'rs.Close
    TABELACOMERCIAL = GetQueryProcess(19)                               'obtem o nome da tabela que possui os dados temporários das ligações
    If frmCanvas.TipoConexao <> 4 Then
        'é Oracle ou SQLServer
        Dim dw As String
        dw = "GS_TEMP"
        Set rs2 = Conn.execute("SELECT * FROM GS_TEMP")
    Else
        'é Postgres
        Set rs2 = conexao.execute("SELECT  * FROM " + """" + "GS_TEMP" + """" + "")
    End If
    'conta o número total de linhas na tabela GS_TEMP
    While Not rs2.EOF
        count2 = count2 + 1
        rs2.MoveNext
    Wend
    rs2.Close
    If count2 >= 1 Then
        If frmCanvas.TipoConexao <> 4 Then
            ConnSec.execute "Delete  From " & TABELACOMERCIAL
        Else
            conexao.execute "Delete  From " + """" + TABELACOMERCIAL + """"
        End If
    End If
    'Conn.execute ("INSERT INTO GS_TEMP(NRO_LIGACAO) VALUES (" & Object_id_trecho & ")")
    Set rsNro_Ligacao = New ADODB.Recordset
    'abre o recordset da tabela temporária de ligações cujo nome está definido no banco de dados em GetQueryProcess(19)
    If frmCanvas.TipoConexao = 1 Then
        'é SQLServer
        rsNro_Ligacao.Open TABELACOMERCIAL, ConnSec, adOpenKeyset, adLockOptimistic, adCmdTable
    ElseIf frmCanvas.TipoConexao = 2 Then
        'é Oracle
        ddd = "SELECT  * FROM GS_TEMP"
        rsNro_Ligacao.Open ddd, ConnSec, adOpenDynamic, adLockOptimistic
    Else
        'é Postgres
        ddd = "SELECT  * FROM " + """" + "GS_TEMP" + """" + ""
        rsNro_Ligacao.Open ddd, conexao, adOpenDynamic, adLockOptimistic
    End If
    'adiciona os números de todas as ligações conectadas aos trechos de rede afetados pela pesquisa (nro_ligacao)
    While Not rs.EOF
       rsNro_Ligacao.AddNew                                 'adiciona nova linha na tabela GS_TEMP com todos os seguimentos (trechos de rede) que foram pintados. GS_TEMP já foi apagada anteriormente
       rsNro_Ligacao.Fields(0).value = rs.Fields(0).value
       rsNro_Ligacao.Update
       rs.MoveNext
    Wend
    '  rs.Close
    'Lv.ListItems.Clear
     ' Set itmx = Lv.ListItems.Add(, , 0)
    '  itmx.SubItems(1) = 0
    '  itmx.SubItems(2) = 0
    '  itmx.SubItems(3) = 0
    '  itmx.SubItems(4) = 0
    '   txtQde = QtdeLig
    ' Me.Show vbModal
    ' 'LoozeXP1.EndWinXPCSubClassing
    'prepara a querie para obter os dados das listas das ligações afetadas pela manobra
    str = GetQueryProcess(18)                               'vai consultar a vista que tem as ligações de água do comercial, mas somente aqueles trechos registrados em GS_TEMP
    If frmCanvas.TipoConexao <> 4 Then
        'caso SQLServer ou Oracle
        numeroErro = "String conexão: " & str
        Set rs = Conn.execute(str)
    Else
        'caso Postgres
        Set rs = conexao.execute(str)
    End If
    'Set rs = ConnSec.execute("SELECT LI.NRO_LIGACAO, (LI.ENDERECO + '-' + ISNULL(LI.NUM_CASA,'') + '-' +  ISNULL(LI.COMPL_LOGRADOURO,'') + '-' + ISNULL(LI.BAIRRO,'')) as Endereco, LI.CONSUMIDOR, LI.TEL_RES AS TELRES, LI.TEL_COM AS TELCOM FROM NXGS_V_LIG_COMERCIAL LI INNER JOIN gs_temp G ON G.NRO_LIGACAO = LI.NRO_LIGACAO")
    'prepara a lista para ser apresentada na caixa de diálogo para o usuário
    If frmCanvas.TipoConexao <> 4 Then
        'caso SQLServer ou Oracle
        lv.ListItems.Clear
        While Not rs.EOF
            Set itmx = lv.ListItems.Add(, , rs.Fields(0).value)
            itmx.SubItems(1) = IIf(IsNull(rs.Fields(1).value), "", rs.Fields(1).value) 'foi inserido o IsNull, pois foi verificado que em alguns bancos comerciais algumas colunas podem vir com valores nulos o que ocasiona um erro
            itmx.SubItems(2) = IIf(IsNull(rs.Fields(2).value), "", rs.Fields(2).value)
            itmx.SubItems(3) = IIf(IsNull(rs.Fields(3).value), "", rs.Fields(3).value)
            itmx.SubItems(4) = IIf(IsNull(rs.Fields(4).value), "", rs.Fields(4).value)
            QtdeLig = QtdeLig + 1                                   'incrementa a quantidade total de ligações
            rs.MoveNext
        Wend
        rs.Close
    Else
        'caso Postgres
        lv.ListItems.Clear
        While Not rs.EOF
            Set itmx = lv.ListItems.Add(, , rs.Fields(0).value)
            itmx.SubItems(1) = rs.Fields(1).value
            itmx.SubItems(2) = rs.Fields(5).value
            itmx.SubItems(3) = rs.Fields(6).value
            itmx.SubItems(4) = rs.Fields(7).value
            QtdeLig = QtdeLig + 1                                   'incrementa a quantidade total de ligações
            rs.MoveNext
        Wend
         rs.Close
    End If
    txtQde = QtdeLig
    Me.Show vbModal
    'LoozeXP1.EndWinXPCSubClassing
    init = True
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmConsumidoresDesabastecidos", "init, object_id selecionado: " & Object_id_trecho & " querie: " & numeroErro, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
        init = False
    End If
End Function
'Subrotina para gerar um arquivo texto, com os dados já gerados pela Função init, contendo todos os consumidores que serão afetados por uma manobra na rede
'
'
'
Private Sub cmdGerartxt_Click()
On Error GoTo Trata_Erro
    Dim a As Integer
    cdl1.filename = GetMyDocumentsDirectory() & "\ClientesAfetadosManobra_" & Format(Now, "YYYY-MM-DD-HHMMSS") & ".txt"
    cdl1.Filter = "Arquivos texto (*.txt)|*.txt"
    cdl1.ShowSave
    If cdl1.filename <> "" Then
        Open cdl1.filename For Output As #1
        For a = 1 To lv.ListItems.count
            Print #1, lv.ListItems.Item(a).Text & ";" & _
                        lv.ListItems.Item(a).SubItems(1) & ";" & _
                        lv.ListItems.Item(a).SubItems(2) & ";" & _
                        lv.ListItems.Item(a).SubItems(3) & ";" & _
                        lv.ListItems.Item(a).SubItems(4)
        Next
        Close #1
        MsgBox "Arquivo gerado com sucesso e disponível no no seguinte endereço: " & cdl1.filename, vbInformation
        Shell "notepad.exe " & cdl1.filename, vbNormalFocus
    End If

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmConsumidoresDesabastecidos", "cmdGerartxt_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub


