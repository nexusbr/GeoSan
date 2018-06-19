VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsumoLote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consultar Consumo"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCarregar 
      Caption         =   "Carregar"
      Height          =   375
      Left            =   3555
      TabIndex        =   9
      Top             =   2865
      Width           =   1140
   End
   Begin VB.CheckBox chkAgrupar 
      Caption         =   "AGRUPAR LIGAÇÕES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4950
      TabIndex        =   8
      Top             =   2910
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin MSComCtl2.DTPicker dtData 
      Height          =   375
      Index           =   0
      Left            =   285
      TabIndex        =   5
      Top             =   2880
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM/yyyy"
      Format          =   59310083
      UpDown          =   -1  'True
      CurrentDate     =   39339
      MaxDate         =   44196
      MinDate         =   32874
   End
   Begin MSComctlLib.ListView lvLigacoes 
      Height          =   1935
      Left            =   315
      TabIndex        =   3
      Top             =   390
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   3413
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
         Text            =   "Inscrição"
         Object.Width           =   3000
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ligação / Matricula"
         Object.Width           =   3000
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ENDEREÇO"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CONSUMIDOR"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   7890
      TabIndex        =   1
      Top             =   6480
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Grafico"
      Height          =   375
      Left            =   7575
      TabIndex        =   0
      Top             =   2895
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComctlLib.ListView lvConsumo 
      Height          =   2130
      Left            =   285
      TabIndex        =   4
      Top             =   3795
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   3757
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MÊS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ANO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CONSUMO"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtData 
      Height          =   375
      Index           =   1
      Left            =   1905
      TabIndex        =   6
      Top             =   2880
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM/yyyy"
      Format          =   59310083
      UpDown          =   -1  'True
      CurrentDate     =   39339
      MaxDate         =   44196
      MinDate         =   32874
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   120
      TabIndex        =   10
      Top             =   3450
      Width           =   8955
      Begin VB.Label LblPeriodo 
         Height          =   345
         Left            =   225
         TabIndex        =   11
         Top             =   2535
         Width           =   5565
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consumidores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   165
      TabIndex        =   12
      Top             =   90
      Width           =   8910
   End
   Begin VB.Label Label2 
      Caption         =   "Inicial:"
      Height          =   255
      Left            =   285
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Final:"
      Height          =   255
      Left            =   1905
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "frmConsumoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Inscricao As String
Dim Ligacoes As String

Private Sub Form_Load()

   Me.dtData(0).value = Now
   Me.dtData(1).value = Now
   
End Sub

Public Sub init(object_id As String, layer_name As String)

On Error GoTo Trata_Erro

   If ConnSec.State = 0 Then
        MsgBox "Não há conexão disponível com o banco de dados comercial para executar esta função.", vbInformation, "Conexão não disponível"
        Exit Sub
   End If

   If layer_name <> GetLayerLot() Then ' forçar como LOTES
      MsgBox "O plano de lotes não está cadastrado corretamente ou não está selecionado  " & GetLayerLot, vbExclamation
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   
   

   
   
   Dim i As ListItem
   Dim rs As ADODB.Recordset
   Dim str As String
   Dim vb As String ' alterado em 20/10/2010
   Dim vc As String
   Dim vk As String
   Dim vm As String
  vb = "QUERYSTRING"
   vc = "GS_QUERYS_CLIENT"
   vk = "QUERY_ID"
   vm = "CLIENT_ID"
   
   If frmCanvas.TipoConexao <> 4 Then

   Set rs = Conn.execute("SELECT querystring from gs_querys_client where query_id=3 and client_id=1")
   str = rs.Fields("querystring").value
   Else
   Set rs = Conn.execute("SELECT " + """" + vb + """" + " from " + """" + vc + """" + " where " + """" + vk + """" + "='3' and " + """" + vm + """" + "='1'")
   str = rs.Fields("QUERYSTRING").value
   End If
   
   
   str = Replace(str, "@OBJECT_ID_", object_id)
   rs.Close
   Set rs = Conn.execute(str) 'SELECT inscrição from lotes where object_id_in()
   
   'Inscricao = rs(0).value 'Alterado pela debaixo para atender ao DAE
   Ligacoes = rs(0).value
If frmCanvas.TipoConexao <> 4 Then
   Set rs = Conn.execute("SELECT querystring from gs_querys_client where client_id=1 and query_id=2")
    str = rs.Fields("querystring").value
   Else
    Set rs = Conn.execute("SELECT " + """" + vb + """" + " from " + """" + vc + """" + " where " + """" + vm + """" + "='1' and " + """" + vk + """" + "='2'")
     str = rs.Fields("QUERYSTRING").value
    End If
  
   
      
   'str = Replace(str, "@CLASSIFICACAO_FISCAL", "'" & Inscricao & "'") 'Alterado pela debaixo para atender ao DAE
   
   str = Replace(str, "@CLASSIFICACAO_FISCAL", "'" & Ligacoes & "'")
   
   
   str = Replace(str, "@NRO_LIGACAO", "''")
   rs.Close
   Set rs = ConnSec.execute(str)
   lvLigacoes.ListItems.Clear
   While Not rs.EOF
      Set i = lvLigacoes.ListItems.Add(, , rs(0).value)
      i.SubItems(1) = IIf(IsNull(rs(1).value), "", rs(1).value)
      i.SubItems(2) = ""
      i.SubItems(3) = IIf(IsNull(rs(2).value), "", rs(2).value)
      rs.MoveNext
   Wend
   rs.Close
   dtData(1).Month = Month(Date)
   dtData(1).Year = Year(Date)
   
   Select Case Month(Date)
      Case 1
         dtData(0).Month = 11
         dtData(0).Year = Year(Date) - 1
      Case 2
         dtData(0).Month = 12
         dtData(0).Year = Year(Date) - 1
      Case Else
         dtData(0).Month = Month(Date) - 2
   End Select
   Screen.MousePointer = vbNormal
   Me.Show vbModal
   Exit Sub
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Sub Init", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Sub

'Private Sub chkAgrupar_Click()
'   dtData_Click 0
'End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCarregar_Click()
   
   Screen.MousePointer = vbHourglass
   
   lvConsumo.ListItems.Clear
   ResultadoAgrupado
   
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmdGrafico_Click()
On Error GoTo Trata_Erro
   
  
    Screen.MousePointer = vbHourglass
    
    Dim Ligacoes As String, a As Integer
    Ligacoes = ""
    For a = 1 To lvLigacoes.ListItems.Count
       If lvLigacoes.ListItems(a).Checked Then
          If Ligacoes = "" Then
             Ligacoes = lvLigacoes.ListItems(a).SubItems(1)
          Else
             Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).SubItems(1)
          End If
       End If
    Next
    Screen.MousePointer = vbNormal
    If Ligacoes <> "" Then
       frmConsumoLoteGraf.init dtData(0).Month, IIf(dtData(1).Month = 12, 1, dtData(1).Month + 1), dtData(0).Year, IIf(dtData(1).Month = 12, dtData(1).Year + 1, dtData(1).Year), _
                               Inscricao, Ligacoes, IIf(chkAgrupar = 1, True, False)
    End If

Trata_Erro:
    Screen.MousePointer = vbNormal
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Sub cmdGrafico_Click", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Sub





'Private Sub lvLigacoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'   dtData_Click 0
'End Sub


Private Function ResultadoAgrupado() As Boolean


On Error GoTo Trata_Erro
   Dim i As ListItem
   
   Dim str As String
   Dim a As Integer

   Dim rs As New ADODB.Recordset


   
   Dim mChecked As Boolean

   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then mChecked = True
         
   Next
   
   If Not mChecked Then
      'MsgBox "Selecione ao menos um hidrometro", vbExclamation
      lvConsumo.ListItems.Clear
      Exit Function
   End If
      
   Dim strVALOR As String
   
   Set rs = New ADODB.Recordset
   
   Ligacoes = ""
   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then
         If Ligacoes = "" Then
            Ligacoes = lvLigacoes.ListItems(a).SubItems(1) '.Text
         Else
            Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).SubItems(1) '.Text
         End If
      End If
   Next
   
   Dim anoIni As String
   Dim dtIni As Date
   Dim ea As String ' alterado em 20/10/2010
   Dim eb As String
   Dim ec As String
   Dim ed As String
   Dim ee As String
   Dim ef As String
   Dim eg As String
   Dim eh As String
   Dim ei As String
   Dim ej As String
   ea = "NXGS_V_LIG_COMERCIAL_CONSUMO"
   eb = "NXGS_V_LIG_COMERCIAL"
   ec = "NRO_LIGACAO"
   ed = "CONSUMO_FATURADO"
   ee = "CONSUMO_MEDIDO"
   ef = "ANO"
   eg = "MES"
   eh = "CLASSIFICACAO_FISCAL"
   ei = "ENDERECO"
   ej = "CONSUMIDOR"
            
   
   
   dtIni = Me.dtData(0).value
   anoIni = Format(dtIni, "YYYY")
   If frmCanvas.TipoConexao <> 4 Then
   If Me.chkAgrupar.value = 0 Then
   

      
      str = "SELECT CO.NRO_LIGACAO, CO.CONSUMO_FATURADO, CO.CONSUMO_MEDIDO, CO.ANO, CO.MES,LI.CLASSIFICACAO_FISCAL,LI.ENDERECO,LI.CONSUMIDOR " & _
            "FROM NXGS_V_LIG_COMERCIAL_CONSUMO CO INNER JOIN NXGS_V_LIG_COMERCIAL LI ON CO.NRO_LIGACAO = LI.NRO_LIGACAO " & _
            "WHERE CO.NRO_LIGACAO IN (" & Ligacoes & ") AND ANO >= " & anoIni & " ORDER BY CO.NRO_LIGACAO, CO.ANO ASC, CO.MES ASC"
   
   Else
      
      str = "SELECT CO.NRO_LIGACAO, CO.CONSUMO_FATURADO, CO.CONSUMO_MEDIDO, CO.ANO, CO.MES,LI.CLASSIFICACAO_FISCAL,LI.ENDERECO,LI.CONSUMIDOR " & _
            "FROM NXGS_V_LIG_COMERCIAL_CONSUMO CO INNER JOIN NXGS_V_LIG_COMERCIAL LI ON CO.NRO_LIGACAO = LI.NRO_LIGACAO " & _
            "WHERE CO.NRO_LIGACAO IN (" & Ligacoes & ") AND ANO >= " & anoIni & " ORDER BY CO.ANO ASC, CO.MES ASC"
   End If
   
   Else
   If Me.chkAgrupar.value = 0 Then
   
   str = "SELECT " + """" + ea + """" + "." + """" + ec + """" + ", " + """" + ea + """" + "." + """" + ed + """" + ", " + """" + ea + """" + "." + """" + ee + """" + ", " + """" + ea + """" + "." + """" + ef + """" + ", " + """" + ea + """" + "." + """" + eg + """" + "," + """" + eb + """" + "." + """" + eh + """" + "," + """" + eb + """" + "." + """" + ei + """" + "," + """" + eb + """" + "." + """" + ej + """" + _
            "FROM " + """" + ea + """" + " INNER JOIN " + """" + eb + """" + " ON " + """" + ea + """" + "." + """" + ec + """" + " =" + """" + eb + """" + "." + """" + ec + """" + _
            "WHERE " + """" + ea + """" + "." + """" + ec + """" + " (" & Ligacoes & ") AND " + """" + ef + """" + ">= " & anoIni & " ORDER BY " + """" + ea + """" + "." + """" + ec + """" + ", " + """" + ea + """" + "." + """" + ef + """" + " ASC, " + """" + ea + """" + "." + """" + eg + """" + " ASC"
   
   Else
      
      str = "SELECT " + """" + ea + """" + "." + """" + ec + """" + ", " + """" + ea + """" + "." + """" + ed + """" + ", " + """" + ea + """" + "." + """" + ee + """" + ", " + """" + ea + """" + "." + """" + ef + """" + ", " + """" + ea + """" + "." + """" + eg + """" + "," + """" + eb + """" + "." + """" + eh + """" + "," + """" + eb + """" + "." + """" + ei + """" + "," + """" + eb + """" + "." + """" + ej + """" + _
            "FROM " + """" + ea + """" + " INNER JOIN " + """" + eb + """" + " ON " + """" + ea + """" + "." + """" + ec + """" + " =" + """" + eb + """" + "." + """" + ec + """" + _
            "WHERE " + """" + ea + """" + "." + """" + ec + """" + " (" & Ligacoes & ") AND " + """" + ef + """" + ">= " & anoIni & " ORDER BY " + """" + ea + """" + "." + """" + ec + """" + ", " + """" + ea + """" + "." + """" + ef + """" + " ASC, " + """" + ea + """" + "." + """" + eg + """" + " ASC"
   'Imprima str
   End If
   End If
   Set rs = New ADODB.Recordset
   rs.Open str, ConnSec, adOpenDynamic, adLockOptimistic
   
   
   lvConsumo.ColumnHeaders.Clear
   
   If Me.chkAgrupar.value = 0 Then lvConsumo.ColumnHeaders.Add , , "Ligação"
   
   lvConsumo.ColumnHeaders.Add , , "Mes / Ano"
   lvConsumo.ColumnHeaders.Add , , "Cons. Medido"
   lvConsumo.ColumnHeaders.Add , , "Cons. Faturado"
   lvConsumo.ListItems.Clear
   
   
   Dim dtFim As Date
   dtFim = Me.dtData(1).value 'data final selecionada no combo
   
   Dim anoMesIni As String
   Dim mesIni As String
   Dim anoMesFim As String
   Dim mesFim As String
   
   Dim NroLigaOld As String
   
   'AnoIni = Format(dtIni, "YYYY") 'carregada acima
   mesIni = Format(dtIni, "MM")
   mesFim = Format(dtFim, "MM")
   anoMesIni = Format(dtIni, "YYYY") & Format(dtIni, "MM")
   anoMesFim = Format(dtFim, "YYYY") & Format(dtFim, "MM")

   
   Dim TotalMed As Double
   Dim TotalFat As Double
   Dim Media As Double
   
   Dim qtdMeses As Integer
   Dim mesAtual As String
   Dim strRsAnoMes As String
   
   TotalMed = 0
   TotalFat = 0
   
   Dim mesOld As Integer, anoOld As Integer, somaMesFat As Double, somaMesMed As Double

novo_consumidor:

   'NRO_LIGACAO | CONSUMO_FATURADO | CONSUMO_MEDIDO | ANO | MES | CLASS.FISCAL | ENDERECO | CONSUMIDOR
   If rs.EOF = False Then
      
      Do While Not rs.EOF
         strRsAnoMes = rs!ano & Format(rs!Mes, "00")
         If CLng(anoMesIni) > CLng(strRsAnoMes) Then 'localiza ano e mes selecionados para a pesquisa
            rs.MoveNext
         Else
            Exit Do
         End If
      Loop
      
      If rs.EOF = False Then
         mesOld = rs!Mes 'determina que o mesOld é atual para entrar na primeira soma
         NroLigaOld = rs!NRO_LIGACAO
      End If
      
      Do While Not rs.EOF
         strRsAnoMes = rs!ano & Format(rs!Mes, "00")
         
         If Me.chkAgrupar.value = 0 Then ' se não esta mandando agrupar...
            If NroLigaOld <> rs!NRO_LIGACAO Then ' verifica se mudou de Nro_ligacao
               
               GoTo novo_consumidor
            
            End If
         End If
         
         If CLng(strRsAnoMes) <= CLng(anoMesFim) Then
                           
            If rs!Mes = mesOld Then
               
               'soma os consumos do mês atual
               somaMesMed = somaMesMed + IIf(IsNull(rs!consumo_medido), 0, rs!consumo_medido)
               somaMesFat = somaMesFat + IIf(IsNull(rs!consumo_faturado), 0, rs!consumo_faturado)
            
            Else
               'imprime resultado do mes
               If Me.chkAgrupar.value = 0 Then

                  Set i = lvConsumo.ListItems.Add(, , NroLigaOld)
                  i.SubItems(1) = Format(mesOld, "00") & "/" & anoOld
                  i.SubItems(2) = somaMesMed
                  i.SubItems(3) = somaMesFat
               
               Else
                  Set i = lvConsumo.ListItems.Add(, , Format(mesOld, "00") & "/" & anoOld)
                  i.SubItems(1) = somaMesMed
                  i.SubItems(2) = somaMesFat
               
               End If
               
               qtdMeses = qtdMeses + 1
               
               'Reseta os somadores do mes
               somaMesMed = IIf(IsNull(rs!consumo_medido), 0, rs!consumo_medido)
               somaMesFat = IIf(IsNull(rs!consumo_faturado), 0, rs!consumo_faturado)
               
            End If
            
            TotalMed = TotalMed + IIf(IsNull(rs!consumo_medido), 0, rs!consumo_medido)
            TotalFat = TotalFat + IIf(IsNull(rs!consumo_faturado), 0, rs!consumo_faturado)
         
         Else
            If Me.chkAgrupar.value = 0 Then
               Set i = lvConsumo.ListItems.Add(, , NroLigaOld)
               i.SubItems(1) = Format(mesOld, "00") & "/" & anoOld
               i.SubItems(2) = somaMesMed
               i.SubItems(3) = somaMesFat
               somaMesMed = 0
               somaMesFat = 0
               Do While Not rs.EOF
                  If rs!NRO_LIGACAO = NroLigaOld Then
                     rs.MoveNext
                  Else
                     Exit Do
                  End If
               Loop
               GoTo novo_consumidor
            End If

            Exit Do
         End If
         mesOld = rs!Mes: anoOld = rs!ano: NroLigaOld = rs!NRO_LIGACAO
         rs.MoveNext
      Loop
      
      'fim da seleção, imprime o resultado do ultimo mês
      If Me.chkAgrupar.value = 0 Then
         Set i = lvConsumo.ListItems.Add(, , NroLigaOld)
         i.SubItems(1) = Format(mesOld, "00") & "/" & anoOld
         i.SubItems(2) = somaMesMed
         i.SubItems(3) = somaMesFat
      Else
         Set i = lvConsumo.ListItems.Add(, , Format(mesOld, "00") & "/" & anoOld)
         i.SubItems(1) = somaMesMed
         i.SubItems(2) = somaMesFat
      End If
      qtdMeses = qtdMeses + 1
   
   End If
   
   rs.Close
   
   Set rs = Nothing
   
   If Me.chkAgrupar.value = 1 Then
      
      Set i = lvConsumo.ListItems.Add(, , "--")
      i.SubItems(1) = "----"
      i.SubItems(2) = "----"
      
      Set i = lvConsumo.ListItems.Add(, , "TOTAL")
      i.SubItems(1) = CStr(TotalMed)
      i.SubItems(2) = CStr(TotalFat)
      
      Set i = lvConsumo.ListItems.Add(, , "--")
      i.SubItems(1) = "----"
      i.SubItems(2) = "----"
      
      Set i = lvConsumo.ListItems.Add(, , "MEDIA")
      If qtdMeses > 0 Then
         i.SubItems(1) = Format(CStr(TotalMed / qtdMeses), "0.0000")
         i.SubItems(2) = Format(CStr(TotalFat / qtdMeses), "0.0000")
         LblPeriodo.Caption = "Localizados / mostrando " & qtdMeses & " meses de consumo."
      Else
         i.SubItems(1) = CStr(0)
         i.SubItems(2) = CStr(0)
         LblPeriodo.Caption = ""
      End If
   
   Else
      
      Set i = lvConsumo.ListItems.Add(, , "--")
      i.SubItems(1) = "----"
      i.SubItems(2) = "----"
      i.SubItems(3) = "----"
      
      Set i = lvConsumo.ListItems.Add(, , "TOTAL")
      i.SubItems(1) = "-->"
      i.SubItems(2) = CStr(TotalMed)
      i.SubItems(3) = CStr(TotalFat)
      
      Set i = lvConsumo.ListItems.Add(, , "--")
      i.SubItems(1) = "----"
      i.SubItems(2) = "----"
      i.SubItems(3) = "----"
      
      Set i = lvConsumo.ListItems.Add(, , "MEDIA")
      i.SubItems(1) = "-->"
      If qtdMeses > 0 Then
         i.SubItems(2) = Format(CStr(TotalMed / qtdMeses), "0.0000")
         i.SubItems(3) = Format(CStr(TotalFat / qtdMeses), "0.0000")
         LblPeriodo.Caption = "Localizados / mostrando " & qtdMeses & " meses de consumo."
      Else
         i.SubItems(2) = CStr(0)
         i.SubItems(3) = CStr(0)
         LblPeriodo.Caption = ""
      End If
   
   End If

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       PrintErro CStr(Me.Name), "Private Sub ResultadoAgrupado", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Function

'Private Sub ResultadoDesagrupado()
'On Error GoTo Trata_Erro
'   Dim i As ListItem, rs As ADODB.Recordset, str As String, a As Integer, Ligacoes As String
'
'   Dim Total As Double
'   Dim Media As Double
'   Dim TotalConsumo As Double
'
'   Dim mChecked As Boolean, UltimaLigacao As String
'
'   For a = 1 To lvLigacoes.ListItems.Count
'      If lvLigacoes.ListItems(a).Checked Then mChecked = True
'   Next
'
'   If Not mChecked Then
'      lvConsumo.ListItems.Clear
'      Exit Sub
'   End If
'
'   Set rs = Conn.execute("SELECT querystring from gs_querys_client where client_id=3 and query_id=4")
'   str = rs.Fields("querystring").value
'
'   'Imprima str
'
'   'str = Replace(str, "@CLASSIFICACAO_FISCAL", Inscricao)
'   str = Replace(str, "@NRO_LIGACAO", Inscricao)
'   str = Replace(str, "@PERIODO1", dtData(0).Month)
'   str = Replace(str, "@PERIODO2", IIf(dtData(1).Month = 12, 1, dtData(1).Month + 1))
'   str = Replace(str, "@ANO1", dtData(0).Year)
'   str = Replace(str, "@ANO2", IIf(dtData(1).Month = 12, dtData(1).Year + 1, dtData(1).Year))
'
'   Ligacoes = ""
''   For a = 1 To lvLigacoes.ListItems.Count
''      If lvLigacoes.ListItems(a).Checked Then
''         If Ligacoes = "" Then
''            Ligacoes = lvLigacoes.ListItems(a).SubItems(1)
''         Else
''            Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).SubItems(1)
''         End If
''      End If
''   Next
'
'   'original
'   For a = 1 To lvLigacoes.ListItems.Count
'      If lvLigacoes.ListItems(a).Checked Then
'         If Ligacoes = "" Then
'            Ligacoes = lvLigacoes.ListItems(a).Text
'         Else
'            Ligacoes = Ligacoes & "," & lvLigacoes.ListItems(a).Text
'         End If
'      End If
'   Next
'
'
'   str = Replace(str, "@CLASSIFICACAO_FISCAL", Ligacoes)
'
'   'str = Replace(str, "@NRO_LIGACAO", Ligacoes)
'
'   rs.Close
'   'Imprima str
'
'   Set rs = ConnSec.execute(str)
'
'   'l.NRO_LIGACAO, h.ano, h.PERIODO, h.consumo as Consumo_Faturado, h.CONSUMO_MEDIDO as Consumo_Medido
'
'   lvConsumo.ColumnHeaders.Clear
'   lvConsumo.ColumnHeaders.Add , , "Ligação"
'   lvConsumo.ColumnHeaders.Add , , "Mes_Ano"
'   lvConsumo.ColumnHeaders.Add , , "Cons. Medido"
'   lvConsumo.ColumnHeaders.Add , , "Cons. Faturado"
'   lvConsumo.ListItems.Clear
'   If rs.EOF Then
'      Exit Sub
'   Else
'      UltimaLigacao = rs(0).value
'   End If
'   While Not rs.EOF
'
'      If rs(0).value <> UltimaLigacao Then
'         Set i = lvConsumo.ListItems.Add(, , "--")
'         i.SubItems(1) = "----"
'         i.SubItems(2) = "----"
'         i.SubItems(3) = "----"
'
'
'         Set i = lvConsumo.ListItems.Add(, , "TOTAL")
'         i.SubItems(1) = "-->"
'         i.SubItems(2) = CStr(TotalConsumo)
'         i.SubItems(3) = CStr(Total)
'
'         Set i = lvConsumo.ListItems.Add(, , "--")
'         i.SubItems(1) = "----"
'         i.SubItems(2) = "----"
'         i.SubItems(3) = "----"
'
'         Set i = lvConsumo.ListItems.Add(, , "MEDIA")
'         i.SubItems(1) = "-->"
'         i.SubItems(2) = CStr(TotalConsumo / Media)
'         i.SubItems(3) = CStr(Total / Media)
'
'         Total = 0
'         TotalConsumo = 0
'         Media = 0
'
'         Set i = lvConsumo.ListItems.Add(, , "")
'
'      End If
'
'      Set i = lvConsumo.ListItems.Add(, , rs(0).value)
'      i.SubItems(1) = rs(1).value
'      i.SubItems(2) = CStr(IIf(IsNull(rs(2).value), 0, rs(2).value))
'      i.SubItems(3) = CStr(IIf(IsNull(rs(3).value), 0, rs(3).value))
'      TotalConsumo = TotalConsumo + IIf(IsNull(rs(2).value), 0, rs(2).value)
'      Total = Total + IIf(IsNull(rs(3).value), 0, rs(3).value)
'      Media = Media + 1
'      UltimaLigacao = rs(0).value
'      rs.MoveNext
'
'
'   Wend
'
'   rs.Close
'
'   Set i = lvConsumo.ListItems.Add(, , "--")
'   i.SubItems(1) = "----"
'   i.SubItems(2) = "----"
'   i.SubItems(3) = "----"
'
'
'   Set i = lvConsumo.ListItems.Add(, , "TOTAL")
'   i.SubItems(1) = "-->"
'   i.SubItems(2) = CStr(TotalConsumo)
'   i.SubItems(3) = CStr(Total)
'
'   Set i = lvConsumo.ListItems.Add(, , "--")
'   i.SubItems(1) = "----"
'   i.SubItems(2) = "----"
'   i.SubItems(3) = "----"
'
'   Set i = lvConsumo.ListItems.Add(, , "MEDIA")
'   i.SubItems(1) = "-->"
'   If Media > 0 Then
'      i.SubItems(2) = CStr(TotalConsumo / Media)
'      i.SubItems(3) = CStr(Total / Media)
'   Else
'      i.SubItems(2) = CStr(0)
'      i.SubItems(3) = CStr(0)
'   End If
'   Set rs = Nothing
'Trata_Erro:
'    If Err.Number = 0 Or Err.Number = 20 Then
'       Resume Next
'    Else
'       Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'       Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmConsumoLote - Private Sub ResultadoDesagrupado - " & Err.Number & " - " & Err.Description
'       Close #1
'       MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'    End If
'End Sub


Private Function GetLayerLot() As String
On Error GoTo Trata_Erro
   
   Dim rs As ADODB.Recordset
   Dim sd As String
   Dim st As String
   sd = "GS_QUERYS_SYSTEM" 'alterado em 20/10/2010
   st = "NAME"
   
   If frmCanvas.TipoConexao <> 4 Then

   Set rs = Conn.execute("SELECT * from gs_querys_system where name='INSCRICAO_LOTE'")
   GetLayerLot = rs.Fields("reference_layer").value
   Else
   Set rs = Conn.execute("SELECT  from " + """" + sd + """" + " where " + """" + st + """" + "='INSCRICAO_LOTE'")
   End If
   GetLayerLot = rs.Fields("reference_layer").value
   rs.Close
   Set rs = Nothing

Trata_Erro:
    
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       PrintErro CStr(Me.Name), "Private Function GetLayerLot", CStr(Err.Number), CStr(Err.Description), True
    End If

End Function




