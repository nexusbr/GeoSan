VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAtualizacaoConsumo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualizações de Consumo"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   5760
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton optAtualizaConsumo 
      Caption         =   "Atualiza todas as ligações de água com o Consumo Médio"
      Height          =   255
      Left            =   195
      TabIndex        =   6
      Top             =   375
      Value           =   -1  'True
      Width           =   5985
   End
   Begin VB.OptionButton optDistDem 
      Caption         =   "Distribuir as demandas de consumo, em l/s, em todos os nós das redes"
      Height          =   255
      Left            =   195
      TabIndex        =   4
      Top             =   840
      Width           =   6600
   End
   Begin VB.OptionButton optImpMedAtuConsDistDem 
      Caption         =   "Importar Medias de Consumo"
      Enabled         =   0   'False
      Height          =   255
      Left            =   195
      TabIndex        =   3
      Top             =   1290
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caminho de Arquivo com Médias de Consumo"
      Enabled         =   0   'False
      Height          =   990
      Left            =   180
      TabIndex        =   1
      Top             =   1785
      Width           =   6810
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   6360
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   390
      Left            =   5850
      TabIndex        =   0
      Top             =   2910
      Width           =   1140
   End
   Begin VB.Label Label1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2970
      Width           =   1350
   End
End
Attribute VB_Name = "frmAtualizacaoConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

' Inicia a execução da atualização dos consumos médios ou distribuição dos consumos
'
'
'
Private Sub Command1_Click()
    On Error GoTo Trata_Erro

    ' if we want to update medium consum in each consumer
    If Me.optAtualizaConsumo.value = True Then
        DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            varGlobais.pararExecucao = False
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        If AtualizaConsumo = True Then
            MsgBox "Consumo de ligações atualizados com sucesso!", vbInformation, ""
        Else
            MsgBox "Falha na atualização de consumo.", vbInformation, ""
        End If
        Exit Sub
    End If
    ' if we want to distribute consume demands
    If Me.optDistDem.value = True Then
        If DISTRIBUI_DEMANDAS = True Then
            MsgBox "Atualização de demanda concluída com sucesso!", vbInformation, "Concluído"
        End If
    End If
    If Me.optImpMedAtuConsDistDem = True Then
        If Me.Text1.Text <> "" Then
            Screen.MousePointer = vbHourglass
            importa_media
            Me.Command1.Enabled = False
        Else
            MsgBox "Caminho de arquivo inválido!", vbExclamation, ""
            Exit Sub
        End If
    End If
    Command1.Enabled = True
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmAtualizacaoConsumo", "Command1", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
'Esta função atualiza os consumos médios da vista do banco de dados comercial (NXGS_V_LIG_COM_CONSUMO_MEDIO) para o banco a tabela do GeoSan RAMAIS_AGUA_LIGACAO
'
'AtualizaConsumo - Retorna True se atualizou corretamente, False se não atualizou o consumo na tabela RAMAIS_AGUA_LIGACAO a partir da vista do sistema comercial
'
Private Function AtualizaConsumo() As Boolean
    On Error GoTo Trata_Erro
    Dim strsql As String                    'string sql
    'ATUALIZA O CONSUMO DAS LIGACOES DE UM RAMAL PUXANDO O VALOR EXISTENTE EM NXGS_V_LIG_COM_CONSUMO_MEDIO
    Screen.MousePointer = vbHourglass
    ' if the database is SqlServer
    If frmCanvas.TipoConexao = 1 Then
        strsql = "UPDATE RAMAIS_AGUA_LIGACAO SET CONSUMO_LPS = (NXGS_V_LIG_COM_CONSUMO_MEDIO.CONSUMO_MEDIO * 0.00038580246) FROM NXGS_V_LIG_COM_CONSUMO_MEDIO WHERE RAMAIS_AGUA_LIGACAO.NRO_LIGACAO / 10 = NXGS_V_LIG_COM_CONSUMO_MEDIO.NRO_LIGACAO_SEM_DV"
        Conn.execute (strsql)
    ' if the database is oracle
    ElseIf frmCanvas.TipoConexao = 2 Then
        'Esta querie necessita ser validada para este banco de dados, a mesma já foi validadada para a conexão do tipo 1, SQLServer
        Conn.execute ("UPDATE RAMAIS_AGUA_LIGACAO SET CONSUMO_LPS = (SELECT NXGS.CONSUMO_MEDIO * 0.00038580246 FROM NXGS_V_LIG_COM_CONSUMO_MEDIO NXGS WHERE RAMAIS_AGUA_LIGACAO.NRO_LIGACAO = NXGS.NRO_LIGACAO)")
    ' if the database is postgres
    ElseIf frmCanvas.TipoConexao = 4 Then
        a = "RAMAIS_AGUA_LIGACAO"
        b = "CONSUMO_LPS"
        c = "CONSUMO_MEDIDO"
        d = "NXGS_V_LIG_COMERCIAL_CONSUMO"
        e = "NRO_LIGACAO"
        Dim conexao As String
        'UPDATE "RAMAIS_AGUA_LIGACAO" SET "CONSUMO_LPS" = (rc."CONSUMO_MEDIDO" *
        ''0.00038580246') FROM "RAMAIS_AGUA_LIGACAO" ra INNER JOIN"NXGS_V_LIG_COMERCIAL_CONSUMO" rc ON ra."NRO_LIGACAO" = rc."NRO_LIGACAO"
        'MsgBox "UPDATE " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " SET " + """" + "CONSUMO_LPS" + """" + " = (N." + """" + "CONSUMO_MEDIDO" + """" + " * 0.00038580246) FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + "  as R INNER JOIN " + """" + "NXGS_V_LIG_COMERCIAL_CONSUMO" + """" + " N  ON R." + """" + "NRO_LIGACAO" + """" + " = N." + """" + "NRO_LIGACAO" + """" + ""
        'MsgBox "ARQUIVO DEBUG SALVO"
        'WritePrivateProfileString "A", "A", "UPDATE " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " SET " + """" + "CONSUMO_LPS" + """" + " = (N." + """" + "CONSUMO_MEDIDO" + """" + " * 0.00038580246) FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + "  as R INNER JOIN " + """" + "NXGS_V_LIG_COMERCIAL_CONSUMO" + """" + " N  ON R." + """" + "NRO_LIGACAO" + """" + " = N." + """" + "NRO_LIGACAO" + """" + "", App.path & "\DEBUG.INI"                                                                                                                                                                                                                                                                                                      ' "CAST(" + """" + d4 + """" + "." + """" + e4 + """" + " AS INTEGER)"
        'Please verify this querie, is problably is wrong because it updates de consumo_medido instead of consumo_medio - 2012-08-21
        'It is not necessary the inerjoin, please look at the SQLServer querie that was tested
        'Esta querie necessita ser validada para este banco de dados, a mesma já foi validadada para a conexão do tipo 1, SQLServer
        Conn.execute ("UPDATE " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " SET " + """" + "CONSUMO_LPS" + """" + " = (N." + """" + "CONSUMO_MEDIDO" + """" + " * 0.00038580246) FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + "  as R INNER JOIN " + """" + "NXGS_V_LIG_COMERCIAL_CONSUMO" + """" + " N  ON CAST(R" + "." + """" + "NRO_LIGACAO" + """" + "AS INTEGER) = CAST(N." + """" + "NRO_LIGACAO" + """" + "AS INTEGER)" + "")
        'Conn.execute ("UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = (" + """" + d + """" + "." + """" + c + """" + " * '0.00038580246') FROM " + """" + a + """" + " INNER JOIN( " + """" + d + """" + " ON " + """" + a + """" + "." + """" + e + """" + " = " + """" + d + """" + "." + """" + e + """" + ")'")
        'Conn.execute (conexao)
    End If
    Screen.MousePointer = vbDefault
    AtualizaConsumo = True

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Screen.MousePointer = vbDefault
        PrintErro CStr(Me.Name), "frmAtualizacaoConsumo - Private Sub AtualizaConsumo(), querie: " & strsql, CStr(Err.Number), CStr(Err.Description), True
        AtualizaConsumo = False
    End If
End Function
' Distribui as demandas de consumo em todos os nós da rede
'
'
'
Private Function DISTRIBUI_DEMANDAS() As Boolean
    On Error GoTo Trata_Erro
    Dim rsCon As New ADODB.Recordset
    Dim rsWATER As New ADODB.Recordset
    Dim redeOld As String, Inicial As String, Final As String
    Dim soma_consumo As Double, metade As Double
    Dim strsql As String
    Dim strMetade As String, strConsumo As String
    Dim rsWTC As New ADODB.Recordset
    Dim TEMPOINI As Date, TEMPOFIM As Date
    Dim contador As Long
    Dim ma As String
    Dim mb As String
    Dim mc As String
    Dim md As String
    Dim mf As String
    Dim mg As String
    Dim mh As String
    Dim mi As String
    Dim mj As String

    Me.Timer1.Enabled = True                                    'habilita o timer
    Me.Command1.Enabled = False
    Screen.MousePointer = vbHourglass
    TEMPOINI = Now
    b = "WATERCOMPONENTS"
    c = "DEMAND"
    ' if it is not Postgres
    If frmCanvas.TipoConexao <> 4 Then
        Conn.execute ("UPDATE WATERCOMPONENTS SET DEMAND = 0")
    Else
        Conn.execute ("UPDATE " + """" + b + """" + " SET " + """" + c + """" + " = '0'")
    End If
    ' open connection to distribute demands
    ' if it is not Postgres
    If frmCanvas.TipoConexao <> 4 Then
        strsql = "SELECT SUM(RAL.CONSUMO_LPS)/2 AS " + """" + "MEDIA_TRECHO" + """" + ",RA.OBJECT_ID_TRECHO,WTR.INITIALCOMPONENT,WTR.FINALCOMPONENT "
        strsql = strsql & "FROM RAMAIS_AGUA_LIGACAO RAL "
        strsql = strsql & "INNER JOIN RAMAIS_AGUA RA ON RAL.OBJECT_ID_ = RA.OBJECT_ID_ INNER JOIN WATERLINES WTR ON WTR.OBJECT_ID_ = RA.OBJECT_ID_TRECHO "
        strsql = strsql & "Where RAL.CONSUMO_LPS > 0 "
        strsql = strsql & "GROUP BY RA.OBJECT_ID_TRECHO,WTR.INITIALCOMPONENT,WTR.FINALCOMPONENT "
        strsql = strsql & "ORDER BY RA.OBJECT_ID_TRECHO,WTR.INITIALCOMPONENT "
        ' if it is Postgres
    Else
        ma = "RAMAIS_AGUA_LIGACAO"
        mb = "CONSUMO_LPS"
        mc = "OBJECT_ID_"
        md = "WATERLINES"
        mf = "INITIALCOMPONENT"
        mg = "FINALCOMPONENT"
        mh = "RAMAIS_AGUA"
        mi = "OBJECT_ID_TRECHO"
        mj = "RAMAIS_AGUA_LIGACAO"
        strsql = "SELECT SUM(" + """" + ma + """" + "." + """" + mb + """" + ")/'2' AS " + """" + "MEDIA_TRECHO" + """" + "," + """" + mh + """" + "." + """" + mi + """" + "," + """" + md + """" + "." + """" + mf + """" + "," + """" + md + """" + "." + """" + mg + """" + " "
        strsql = strsql & "FROM " + """" + ma + """" + ""
        strsql = strsql & "INNER JOIN " + """" + mh + """" + " ON " + """" + ma + """" + "." + """" + mc + """" + " = " + """" + mh + """" + "." + """" + mc + """" + " INNER JOIN " + """" + md + """" + "  ON " + """" + md + """" + "." + """" + mc + """" + " = " + """" + mh + """" + "." + """" + mi + """" + " "
        strsql = strsql & "Where " + """" + ma + """" + "." + """" + mb + """" + " > '0' "
        strsql = strsql & "GROUP BY " + """" + mh + """" + "." + """" + mi + """" + "," + """" + md + """" + "." + """" + mf + """" + "," + """" + md + """" + "." + """" + mg + """" + " "
        strsql = strsql & "ORDER BY " + """" + mh + """" + "." + """" + mi + """" + "," + """" + md + """" + "." + """" + mf + """" + " "
    End If
    Set rsCon = New ADODB.Recordset
    rsCon.Open strsql, Conn, adOpenDynamic, adLockReadOnly
    Set rsCon = New ADODB.Recordset
    rsCon.Open strsql, Conn, adOpenDynamic, adLockReadOnly
    Do While Not rsCon.EOF = True
        DoEvents                                                                'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
        If varGlobais.pararExecucao = True Then
            varGlobais.pararExecucao = False
            Screen.MousePointer = vbNormal
            rsCon.Close
            Exit Function
        End If
        contador = contador + 1
        rsCon.MoveNext
    Loop
    rsCon.Close
    Set rsCon = Nothing
    Me.ProgressBar1.Max = contador + 1
    Me.ProgressBar1.value = 1
    Me.ProgressBar1.Visible = True
    Set rsCon = New ADODB.Recordset
    rsCon.Open strsql, Conn, adOpenDynamic, adLockReadOnly
    'MEDIA_TRECHO,OBJECT_ID_TRECHO,INITIALCOMPONENT,FINALCOMPONENT
    If rsCon.EOF = False Then
        Do While Not rsCon.EOF = True
            DoEvents                                                                'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
            If varGlobais.pararExecucao = True Then
                varGlobais.pararExecucao = False
                Screen.MousePointer = vbNormal
                rsCon.Close
                Exit Function
            End If
            strMetade = Replace(rsCon!MEDIA_TRECHO, ",", ".")                       'já vem dividido por dois pelo Select acima
            STRINICIAL = rsCon!INITIALCOMPONENT
            STRFINAL = rsCon!FinalComponent
            a = "WATERCOMPONENTS"
            b = "DEMAND"
            c = "INSCRICAO_LOTE"
            d = "OBJECT_ID_"
            e = strMetade
            'f = "e'"
            If frmCanvas.TipoConexao <> 4 Then
                'comando que joga para os 2 nós ponta o consumo da rede
                On Error Resume Next
                Conn.execute ("UPDATE WATERCOMPONENTS SET DEMAND = DEMAND + " & strMetade & " WHERE OBJECT_ID_ IN ('" & STRINICIAL & "','" & STRFINAL & "')")
            Else
                On Error Resume Next
                Conn.execute ("UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = '" & strMetade & "' WHERE " + """" + d + """" + " IN ('" & STRINICIAL & "','" & STRFINAL & "')")
            End If
            rsCon.MoveNext
            Me.ProgressBar1.value = Me.ProgressBar1.value + 1
        Loop
    End If
    rsCon.Close
    Set rsCon = Nothing
    Screen.MousePointer = vbNormal
    Me.ProgressBar1.Visible = False
    DISTRIBUI_DEMANDAS = True
    Exit Function

Trata_Erro:
        If Err.Number = 0 Or Err.Number = 20 Then
            Resume Next
        Else
            DISTRIBUI_DEMANDAS = False
            ErroUsuario.Registra "frmAtualizacaoConsumo", "DISTRIBUI_DEMANDAS", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
        End If
End Function

Private Function importa_media()

On Error GoTo Trata_Erro

Dim linha As String
Dim Vetor As Variant
Dim i As Integer
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Cont As Long

   'anoMês;imov_ID;cons_medio
   a = "NXGS_V_LIG_COM_CONSUMO_MEDIO"


     If frmCanvas.TipoConexao <> 4 Then
   Conn.execute ("DELETE FROM NXGS_V_LIG_COM_CONSUMO_MEDIO")
   Else
      Conn.execute ("DELETE FROM " + a + "")
   End If
   
   'Conn.execute ("DROP TABLE CONSUMO")
   
a = "CONSUMO"
b = "MESANO"
c = "IMOVEL"
d = "CONSUMO"

     If frmCanvas.TipoConexao <> 4 Then
   Conn.execute ("CREATE TABLE CONSUMO (MESANO [char] (12),IMOVEL [char](12), CONSUMO [FLOAT])")
   Else
      Conn.execute ("CREATE TABLE " + """" + a + """" + " (" + """" + b + """" + " character varying(50) ," + """" + c + """" + " character varying(50) ," + """" + d + """" + " float)")
   End If
   Cont = 0
   Open Me.Text1.Text For Input As #3
   Do While Not EOF(3)
      Input #3, linha
      Cont = Cont + 1
   Loop
   Close #3
   
   ProgressBar1.value = 1
   ProgressBar1.Max = Cont + 10
   ProgressBar1.Visible = True
   
   Open Me.Text1.Text For Input As #3

   Do While Not EOF(3)
      DoEvents
      
      Input #3, linha
      Vetor = Split(linha, ";")
      
      a = "CONSUMO"
      b = "MESANO"
      c = "IMOVEL"
      d = "CONSUMO"
      e = "HIDROMETRADO"
      f = "ECONOMIAS"
      g = "CONSUMO_LPS"
      h = "TB_LIGACOES"
      i = "HIDROMETRADO"
      j = "ECONOMIAS"
      k = "CONSUMO_LPS"
      l = "IMOVEL"


     If frmCanvas.TipoConexao <> 4 Then
         
      Conn.execute ("INSERT INTO CONSUMO (MESANO,IMOVEL,CONSUMO) VALUES ('" & Vetor(0) & "','" & Vetor(1) & "','" & Vetor(2) & "')")
     Else
     
       Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ") VALUES ('" & Vetor(0) & "','" & Vetor(1) & "','" & Vetor(2) & "')")
     End If
      
                        
      ProgressBar1.value = ProgressBar1.value + 1
      
      
   Loop
   Close #3
   
   Dim ID_IMOVEL As String
   Dim Media As String
   
      a = "CONSUMO"
      b = "IMOVEL"

      
   
   If frmCanvas.TipoConexao <> 4 Then
   SQL = "SELECT CONSUMO.IMOVEL, SUM(CONSUMO.CONSUMO)/COUNT(CONSUMO.CONSUMO) AS " + """" + "MEDIA_CONS" + """" + " FROM CONSUMO AS " + """" + "CONSUMO" + """" + " GROUP BY CONSUMO.IMOVEL"
   Else
   SQL = "SELECT " + """" + a + """" + "." + """" + b + """" + "," + """" + " SUM(" + """" + a + """" + "." + """" + a + """" + ")/COUNT(" + """" + a + """" + "." + """" + a + """" + ") AS " + """" + "MEDIA_CONS" + """" + " FROM " + """" + a + """" + " AS " + """" + "CONSUMO" + """" + " GROUP BY " + """" + a + """" + "." + """" + b + """" + ""
   End If
   
   
   Set rs = Conn.execute(SQL)
   
   Cont = 0
   If rs.EOF = False Then
      Do While Not rs.EOF
         Cont = Cont + 1
         rs.MoveNext
      Loop
      Me.ProgressBar1.Max = Cont + 10
      Me.ProgressBar1.value = 1
      Set rs = Conn.execute(SQL)
      Do While Not rs.EOF
         DoEvents
         Media = Trim(rs!MEDIA_CONS)
         Media = Replace(Media, ",", ".")
         
         ID_IMOVEL = Trim(rs!IMOVEL)
         
         
      a = "NXGS_V_LIG_COM_CONSUMO_MEDIO"
      b = "NRO_LIGACAO"
      c = "CONSUMO_MEDIO"
      

     If frmCanvas.TipoConexao <> 4 Then
         
      Conn.execute ("INSERT INTO NXGS_V_LIG_COM_CONSUMO_MEDIO (NRO_LIGACAO,CONSUMO_MEDIO) VALUES ('" & ID_IMOVEL & "'," & Media & ")")
     
     Else
     
       Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + ") VALUES ('" & ID_IMOVEL & "'," & Media & ")")
     End If
         
         Me.ProgressBar1.value = Me.ProgressBar1.value + 1
         
         rs.MoveNext
      
      Loop
   End If
   
   Me.ProgressBar1.Visible = False

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = "-2147217900" Or Err.Number = "-2147217913" Then
   Resume Next
Else
   MsgBox Err.Number & " - " & Err.Description
   Resume Next
End If

End Function

Private Sub Form_Load()
    Me.Timer1.Interval = 100                               'define o intervalo em que ele vai verificar se alguma tecla foi pressionada
    Me.Timer1.Enabled = False                              'inicia com o timer desligado, só liga quando tiver cálculo intensivo
End Sub

Private Sub optAtualizaConsumo_Click()
   Frame1.Enabled = False
End Sub

Private Sub optDistDem_Click()
   Frame1.Enabled = False
End Sub


Private Sub optImpMedAtuConsDistDem_Click()
   Frame1.Enabled = True
End Sub




    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
   '
   'O PROCEDIMENTO DE TRANSFERIR A DEMANDA DO NÓ PARA O NÓ VIZINHO CASO ESTE SEJA VÁLVULA,BOMBA OU RESERVATÓRIO
   'QUE ESTÁ ABAIXO FOI COMENTADO APÓS VERIFICAR COM PINHEIRO QUE ESTE PROCEDIMENTO NÃO SEJA NECESSÁRIO


''
''   'SELECIONA-SE TODOS OS NÓS QUE SÃO DO TIPO VALVULA OU BOMBA OU RESERVATORIO QUE A DEMANDA SEJA MAIOR QUE ZERO
''   strSQL = "SELECT OBJECT_ID_,DEMAND FROM WATERCOMPONENTS WHERE ID_TYPE IN ( SELECT ID_TYPE FROM WATERCOMPONENTSTYPES WHERE DESCRIPTION_ IN ('VRP','BOMBA','REGISTRO')) AND DEMAND > 0 ORDER BY ID_TYPE"
''   Set rsWTC = Conn.execute(strSQL)
''
''   'RETORNA: | OBJECT_ID_ | DEMAND
''
''   If rsWTC.EOF = False Then
''      Do While Not rsWTC.EOF = True
''         Set rsWATER = Conn.execute("SELECT INITIALCOMPONENT, FINALCOMPONENT FROM WATERLINES WHERE INITIALCOMPONENT = " & rsWTC!Object_id_)
''         If rsWATER.EOF = False Then ' PELO NÓ INICIAL, ACHEI O NÓ FINAL... QUE RECEBE A DEMANDA
''
''            Final = rsWATER!FINALCOMPONENT
''            strConsumo = Replace(rsWTC!DEMAND, ",", ".")
''            strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = DEMAND + " & strConsumo & " WHERE OBJECT_ID_ = '" & Final & "'"
''
''         Else
''            Set rsWATER = Conn.execute("SELECT INITIALCOMPONENT, FINALCOMPONENT FROM WATERLINES WHERE FINALCOMPONENT = " & rsWTC!Object_id_)
''            If rsWATER.EOF = False Then ' PELO NÓ FINAL, ACHEI O NÓ INICIAL... QUE RECEBE A DEMANDA
''
''               Inicial = rsWATER!INITIALCOMPONENT
''               strConsumo = Replace(rsWTC!DEMAND, ",", ".")
''               strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = DEMAND + " & strConsumo & " WHERE OBJECT_ID_ = '" & Inicial & "'"
''            Else
''               MsgBox ""
''            End If
''
''         End If
''
''         Conn.execute (strSQL) 'EXECUTA O COMANDO DO UPDATE ACIMA
''
''         strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = 0 WHERE OBJECT_ID_ = '" & rsWTC!Object_id_ & "'"
''         Conn.execute (strSQL) 'ZERA O VALOR DE DEMANDA DA VALVULA, BOMBA OU RESERVATORIO QUE ESTAVA EM ANÁLISE
''
''         rsWTC.MoveNext
''      Loop
''   End If
''
''
''   'PODE OCORRER DE UM NÓ DO TIPO VÁLVULA OU BOMBA OU RESERVATORIO TER SIDO ATUALIZADO COM O VALOR DE SEU VIZINHO ..
''
''   'INICIA-SE UM NOVO PROCESSO, ESPECÍFICO NOS COMPONENTES DO TIPO VÁLVULA OU BOMBA OU RESERVATORIO
''   'VERIFICANDO SE AINDA HÁ VÁLVULA OU BOMBA OU RESERVATORIO COM VALORES MAIORES QUE ZERO NA DEMANDA
''   strSQL = "SELECT OBJECT_ID_,DEMAND FROM WATERCOMPONENTS WHERE ID_TYPE IN ( SELECT ID_TYPE FROM WATERCOMPONENTSTYPES WHERE DESCRIPTION_ IN ('VRP','BOMBA','REGISTRO')) AND DEMAND > 0 ORDER BY ID_TYPE"
''   Set rsWTC = Conn.execute(strSQL)
''
''   'RETORNA: | OBJECT_ID_ | DEMAND
''
''   If rsWTC.EOF = False Then
''      Do While Not rsWTC.EOF = True
''         'SELECT INITIALCOMPONENT, FINALCOMPONENT FROM WATERLINES WHERE INITIALCOMPONENT = 302
''         strSQL = "SELECT INITIALCOMPONENT, FINALCOMPONENT FROM WATERLINES WHERE INITIALCOMPONENT = " & rsWTC!Object_id_
''         Set rsCon = Conn.execute(strSQL)
''         If rsCon.EOF = False Then
''            Do While Not rsCon.EOF = True
''               '--PELO COMPONENTE 302 ACHEI O 298
''               '--PESQUISA SE ELE É UM COMPONENTE VÁLIDO PARA RECEBER A DEMANDA, OU SEJA DIFERENTE DE 'VRP','BOMBA','REGISTRO'
''               strSQL = "SELECT * FROM WATERCOMPONENTS WHERE OBJECT_ID_ = " & rsCon!FINALCOMPONENT & " AND ID_TYPE NOT IN ( SELECT ID_TYPE FROM WATERCOMPONENTSTYPES WHERE DESCRIPTION_ IN ('VRP','BOMBA','REGISTRO'))"
''               Set rsCon = Conn.execute(strSQL)
''               '--CASO O SELECT ACIMA RETORNE VALORES, O COMPONENTE 298 RECEBE A DEMANDA DO COMPONENTE 302, SE NÃO PESQUISA DE NOVO USANDO O 302 COMO FINALCOMPONENT
''               If rsCon.EOF = False Then
''
''                  strConsumo = Replace(rsWTC!DEMAND, ",", ".")
''                  strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = DEMAND + " & strConsumo & " WHERE OBJECT_ID_ = '" & rsCon!Object_id_ & "'"
''                  Conn.execute (strSQL) 'EXECUTA O COMANDO DO UPDATE ACIMA
''
''                  strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = 0 WHERE OBJECT_ID_ = '" & rsWTC!Object_id_ & "'"
''                  Conn.execute (strSQL) 'ZERA O VALOR DE DEMANDA DA VALVULA, BOMBA OU RESERVATORIO QUE ESTAVA EM ANÁLISE
''                  Exit Do
''
''               End If
''               If rsCon.EOF = False Then
''                  rsCon.MoveNext
''               Else
''                  'NENHUM VIZINHO DO PONTO É VÁLIDO
''               End If
''            Loop
''         Else
''            'PROCURA POR INICIALCOMPONENT
''            'SELECT INITIALCOMPONENT, FINALCOMPONENT FROM WATERLINES WHERE FINALCOMPONENT = 302
''            strSQL = "SELECT INITIALCOMPONENT, FINALCOMPONENT FROM WATERLINES WHERE FINALCOMPONENT = " & rsWTC!Object_id_
''            Set rsCon = Conn.execute(strSQL)
''            If rsCon.EOF = False Then
''               Do While Not rsCon.EOF = True
''                  '--PESQUISA SE ELE É UM COMPONENTE VÁLIDO PARA RECEBER A DEMANDA, OU SEJA DIFERENTE DE 'VRP','BOMBA','REGISTRO'
''                  strSQL = "SELECT * FROM WATERCOMPONENTS WHERE OBJECT_ID_ = " & rsCon!INITIALCOMPONENT & " AND ID_TYPE NOT IN ( SELECT ID_TYPE FROM WATERCOMPONENTSTYPES WHERE DESCRIPTION_ IN ('VRP','BOMBA','REGISTRO'))"
''                  Set rsCon = Conn.execute(strSQL)
''                  '--CASO O SELECT ACIMA RETORNE VALORES, O COMPONENTE 298 RECEBE A DEMANDA DO COMPONENTE 302, SE NÃO PESQUISA DE NOVO USANDO O 302 COMO INITIALCOMPONENT
''                  If rsCon.EOF = False Then
''
''                     strConsumo = Replace(rsWTC!DEMAND, ",", ".")
''                     strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = DEMAND + " & strConsumo & " WHERE OBJECT_ID_ = '" & rsCon!Object_id_ & "'"
''                     Conn.execute (strSQL) 'EXECUTA O COMANDO DO UPDATE ACIMA
''
''                     strSQL = "UPDATE WATERCOMPONENTS SET DEMAND = 0 WHERE OBJECT_ID_ = '" & rsWTC!Object_id_ & "'"
''                     Conn.execute (strSQL) 'ZERA O VALOR DE DEMANDA DA VALVULA, BOMBA OU RESERVATORIO QUE ESTAVA EM ANÁLISE
''                     Exit Do
''
''                  End If
''                  If rsCon.EOF = False Then
''                     rsCon.MoveNext
''                  Else
''                     'NENHUM VIZINHO DO PONTO É VÁLIDO
''                  End If
''               Loop
''            End If
''         End If
''
''         rsWTC.MoveNext
''
''      Loop
''
''   End If
''
''   'AO FIM DESSE PROCESSO, SE AINDA EXISTEM COMPONENTES DO TIPO VÁLVULA OU BOMBA OU RESERVATORIO COM VALOR DE DEMANDA
''   'ISSO QUER DIZER QUE PROVÁVELMENTE HÁ UM ERRO NO CADASTRO DOS COMPONENTES
''
''   strSQL = "SELECT OBJECT_ID_,DEMAND FROM WATERCOMPONENTS WHERE ID_TYPE IN ( SELECT ID_TYPE FROM WATERCOMPONENTSTYPES WHERE DESCRIPTION_ IN ('VRP','BOMBA','REGISTRO')) AND DEMAND > 0 ORDER BY ID_TYPE"
''   Set rsWTC = Conn.execute(strSQL)
''
''   'RETORNA: | OBJECT_ID_ | DEMAND
''
''   If rsWTC.EOF = False Then
''      Do While Not rsWTC.EOF = True
''         MsgBox "VERIFICAR O COMPONENTE: " & rsWTC!Object_id_
''         rsWTC.MoveNext
''      Loop
''   Else
''      MsgBox "Demandas de consumo atualizadas com sucesso!", vbExclamation, ""
''   End If

   'MsgBox "Demandas de consumo atualizadas com sucesso!", vbExclamation, ""

' Configura o Timer para o usuário poder apertar a tecla ESC e ele cancelar a operação
Private Sub Timer1_Timer()
    If GetAsyncKeyState(VK_ESCAPE) Then         'pressionou ESC e vai avisar para parar a execução
        MsgBox ("Comando cancelado.")
        varGlobais.pararExecucao = True
    End If
End Sub
