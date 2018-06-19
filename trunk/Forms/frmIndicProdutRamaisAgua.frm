VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndicProdutRamaisAgua 
   Caption         =   "Indicador de Produtividade - Ligações de Água"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   150
      TabIndex        =   3
      Top             =   1065
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.TextBox txtCaminho 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   435
      Width           =   6060
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   360
      Left            =   5025
      TabIndex        =   0
      Top             =   885
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   165
      Left            =   150
      TabIndex        =   4
      Top             =   930
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.Label lblCaminho 
      Caption         =   "Caminho do Arquivo"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   1605
   End
End
Attribute VB_Name = "frmIndicProdutRamaisAgua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gera o relatório de produtividade do cadastro de ramais e ligações de água
'
'
Private Sub cmdGerar_Click()
    On Error GoTo Trata_Erro
    MousePointer = vbHourglass
    Dim rs As ADODB.Recordset
    Dim rsMeta As ADODB.Recordset
    Dim strDataR, strUserR As String
    Dim contBar As Long
    Dim strsql As String
    Dim dataOld As String
    Dim TotalLigacoes As Long
    Dim TotalLigacoesDoDia As Long
    Dim TotalHistoricoLicacoes As Long

    Conn.execute ("UPDATE RAMAIS_AGUA SET USUARIO_LOG = 'DESCONHECIDO' WHERE USUARIO_LOG is null")
    Conn.execute ("UPDATE RAMAIS_AGUA SET DATA_LOG = '01/01/01 01:01' WHERE DATA_LOG is null")
    strDataR = Format(Now, "DD/MM/YY")

    'IMPRIME O RELATÓRIO DO DIA QUE É DEFINIDO PELA DATA DA MAQUINA
    ProgressBar1.value = 2
    strsql = "SELECT COUNT(*) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE LEFT(DATA_LOG,8) = '" & strDataR & "')"
    Set rs = New ADODB.Recordset
    rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
    Open txtCaminho.Text For Output As #2
    Print #2, "****************** SISTEMA GEOSAN **********************"
    Print #2, "######### RELATÓRIO INDICATIVO DE PRODUTIVIDADE ########"
    Print #2, "############ CADASTRO DE LIGAÇÕES DE ÁGUA ##############"
    Print #2, "INÍCIO - *************************** " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    Print #2, ""
    Print #2, ""
    If rs.EOF = False Then
        Print #2, "********************************************************"
        Print #2, "****************** RESUMO DO DIA *****************INÍCIO"
        Print #2, ""
        Print #2, "DATA"; Tab(30); "LIGAÇÕES"
        Print #2, "========================================================"
        Print #2, strDataR; Tab(15); "Total do Dia"; Tab(30); rs!linhas
        Print #2, ""
        Print #2, "****************** RESUMO DO DIA ******************* FIM"
        Print #2, "********************************************************"
        Print #2, ""
        Print #2, ""
        Print #2, ""
    End If
    Close #2
    rs.Close
    
    'MONTAGEM DO RELATÓRIO DIÁRIO DE LIGAÇÕES CADASTRAS - NÃO CONTA RAMAIS, SOMENTE LIGAÇÕES
    'SELECT USUARIO_LOG, LEFT(LEFT(DATA_LOG,8),2) AS DIA, RIGHT(LEFT(DATA_LOG,5),2) AS MES, RIGHT(LEFT(DATA_LOG,10),2) AS ANO, LEFT(DATA_LOG,10) AS DATA, COUNT(Object_id_) As Ligacoes FROM RAMAIS_AGUA_LIGACAO
    'Where Len(USUARIO_LOG) > 0 And Len(DATA_LOG) > 0
    'GROUP BY USUARIO_LOG, LEFT(DATA_LOG,10), LEFT(LEFT(DATA_LOG,8),2), RIGHT(LEFT(DATA_LOG,5),2), RIGHT(LEFT(DATA_LOG,10),2)
    'ORDER BY ANO,MES,DIA,USUARIO_LOG
    strsql = "SELECT USUARIO_LOG, LEFT(LEFT(DATA_LOG,8),2) AS DIA, RIGHT(LEFT(DATA_LOG,5),2) AS MES, RIGHT(LEFT(DATA_LOG,10),2) AS ANO, LEFT(DATA_LOG,10) AS DATA, COUNT(Object_id_) As Ligacoes FROM RAMAIS_AGUA_LIGACAO"
    strsql = strsql & " Where Len(USUARIO_LOG) > 0 And Len(DATA_LOG) > 0"
    strsql = strsql & " GROUP BY USUARIO_LOG, LEFT(DATA_LOG,10), LEFT(LEFT(DATA_LOG,8),2), RIGHT(LEFT(DATA_LOG,5),2), RIGHT(LEFT(DATA_LOG,10),2)"
    strsql = strsql & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
    Set rs = New ADODB.Recordset
    rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
    TotalLigacoesDoDia = 0
    TotalHistoricoLicacoes = 0
    Open txtCaminho.Text For Append As #2
    Print #2, "********************************************************"
    Print #2, "**** HISTÓRICO DIÁRIO DE LIGAÇÕES CADASTRADAS *** INÍCIO"
    Print #2, "========================================================"
    Print #2, "DATA"; Tab(15); "USUARIO"; Tab(30); "LIGAÇÕES"
    Print #2, "========================================================"
    If rs.EOF = False Then
        dataOld = rs!data
        Do While Not rs.EOF
            'IMPRIME O TOTAL DIA DO USUÁRIO
            If dataOld = rs!data Then
                TotalLigacoesDoDia = TotalLigacoesDoDia + rs!Ligacoes
                TotalHistoricoLicacoes = TotalHistoricoLicacoes + rs!Ligacoes
                Print #2, Trim(rs!data); Tab(15); Trim(rs!USUARIO_LOG); Tab(30); Trim(rs!Ligacoes)
            Else ' TROCOU DE DATA
                Print #2, "========================================================"
                Print #2, dataOld; Tab(15); "Total do Dia"; Tab(30); CStr(TotalLigacoesDoDia)
                Print #2, ""
                Print #2, ""
                TotalLigacoesDoDia = rs!Ligacoes
                TotalHistoricoLicacoes = TotalHistoricoLicacoes + rs!Ligacoes
                Print #2, rs!data; Tab(15); Trim(rs!USUARIO_LOG); Tab(30); Trim(rs!Ligacoes)
            End If
            dataOld = rs!data
            rs.MoveNext
        Loop
        'Imprime o último total do dia até a data do relatório
        Print #2, "========================================================"
        Print #2, dataOld; Tab(15); "Total do Dia"; Tab(30); CStr(TotalLigacoesDoDia)
        Print #2, ""
        Print #2, ""
        'Imprime o total geral de todos os dias
        Print #2, "========================================================"
        Print #2, ""
        Print #2, dataOld; Tab(15); "Total geral de ligações cadastradas"; Tab(30); CStr(Trim(TotalHistoricoLicacoes))
        Print #2, ""
        Print #2, "Obs: este relatório apresenta apenas as ligações de água"
        Print #2, "cadastradas a partir do GeoSan versão 7.5.0"
        Print #2, ""
    Else
        Print #2, "NÃO HÁ INFORMAÇÕES PARA HISTÓRICO DIÁRIO DE USUÁRIO ****"
        Print #2, ""
    End If
    Print #2, "********** HISTÓRICO DIÁRIO POR USUÁRIO ************ FIM"
    Print #2, "********************************************************"
    Print #2, ""
    Print #2, ""
    Print #2, ""
    Close #2
    
    'MONTAGEM DO RELATÓRIO RESUMO CONSOLIDADO (ACUMULADO) DE USUÁRIO
    '1 - SELECT DISTINCT LEFT(DATA_LOG,8)as data,USUARIO_LOG FROM WATERLINES ORDER BY DATA,USUARIO_LOG
    '2 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = 'Jonathas'
    '3 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES
    TotalLigacoes = 0
    Set rsMeta = Conn.execute("SELECT DISTINCT USUARIO_LOG FROM RAMAIS_AGUA WHERE LEN(USUARIO_LOG) > 0 ORDER BY USUARIO_LOG")
    contBar = 0
    If rsMeta.EOF = False Then
        Do While Not rsMeta.EOF = True
            rsMeta.MoveNext
            contBar = contBar + 1
        Loop
    End If
    ProgressBar2.value = 0
    ProgressBar2.Max = contBar + 5
    rsMeta.Requery
    ProgressBar1.value = 6
    Open txtCaminho.Text For Append As #2
    Print #2, "********************************************************"
    Print #2, "******** RESUMO CONSOLIDADO POR USUÁRIO ********* INÍCIO"
    If rsMeta.EOF = False Then
        strUserR = rsMeta!USUARIO_LOG
        Print #2, "========================================================"
        Print #2, ""; Tab(15); "USUARIO"; Tab(30); "LIGAÇÕES"
        Print #2, "========================================================"
        Do While Not rsMeta.EOF = True
            DoEvents
            Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS TotalLigacoesPorUsuario FROM RAMAIS_AGUA_LIGACAO WHERE USUARIO_LOG = '" & strUserR & "'")
            If rs.EOF = False Then
                'IMPRIME O TOTAL DIA DO USUÁRIO
                Print #2, ""; Tab(15); strUserR; Tab(30); rs!TotalLigacoesPorUsuario
                TotalLigacoes = TotalLigacoes + rs!TotalLigacoesPorUsuario
            End If
                rsMeta.MoveNext
                ProgressBar2.value = ProgressBar2.value + 1
            If rsMeta.EOF = False Then
                strUserR = rsMeta!USUARIO_LOG
            Else
                'IMPRIME O TOTAL GERAL DA BASE DE DADOS
                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS TotalLigacoesGeral FROM RAMAIS_AGUA_LIGACAO")
                Print #2, ""
                Print #2, "TOTAL CADASTRADO ATÉ " & Format(Now, "DD/MM/YYYY HH:MM:SS"); Tab(30); CStr(TotalLigacoes)
                Print #2, ""
                Print #2, "********** RESUMO CONSOLIDADO POR USUÁRIO ********** FIM"
                Print #2, "********************************************************"
                Print #2, ""
                Print #2, "Obs: este relatório apresenta apenas as ligações de água"
                Print #2, "cadastradas a partir do GeoSan versão 7.5.0"
                Print #2, ""
                Print #2, ""
                Print #2, "TOTAL GERAL DE LIGAÇÕES E RAMAIS CADASTRADOS"; Tab(30); "LIGAÇÕES"
                Print #2, "========================================================"
                Print #2, "ATÉ " & Format(Now, "DD/MM/YYYY HH:MM:SS"); Tab(30); rs!TotalLigacoesGeral
                Print #2, ""
                Print #2, ""
                Print #2, ""
                Exit Do
            End If
        Loop
    Else
        'RESUMO CONSOLIDADO DE USUÁRIO
        Print #2, "NÃO HÁ INFORMAÇÕES PARA RESUMO CONSOLIDADO DE USUÁRIO **"
        Print #2, ""
    End If

    'MONTAGEM DO RELATÓRIO DIA A DIA DOS RAMAIS CADASTRADOS SEPARADO POR PONTO E VIRGULA
    'Para contar quantas ligações estão cadastradas e mostrar o andamento do processamento na barra de progresso
    Set rsMeta = Conn.execute("SELECT DISTINCT LEFT(DATA_LOG,10) AS DATA,LEFT(LEFT(DATA_LOG,8),2) AS DIA,RIGHT(LEFT(DATA_LOG,5),2) AS MES,RIGHT(LEFT(DATA_LOG,8),2) AS ANO,USUARIO_LOG FROM RAMAIS_AGUA_LIGACAO WHERE LEN(USUARIO_LOG) > 0 AND LEN(DATA_LOG) > 0 ORDER BY ANO,MES,DIA")
    contBar = 0
    If rsMeta.EOF = False Then
        Do While Not rsMeta.EOF = True
            rsMeta.MoveNext
            contBar = contBar + 1
        Loop
    End If
    ProgressBar2.value = 0
    ProgressBar2.Max = contBar
    rsMeta.Requery
    ProgressBar1.value = 10
    Print #2, "********************************************************"
    Print #2, "HISTÓRICO DIÁRIO DE USUÁRIO SEPARADO POR ; ****** INÍCIO"
    Print #2, "Representa o cadastro total por ramais cadastrados"
    Print #2, ""
    If rsMeta.EOF = False Then
        strDataR = rsMeta!data
        strUserR = rsMeta!USUARIO_LOG
        Print #2, "DATA;USUARIO;LIGAÇÕES"
        Do While Not rsMeta.EOF = True
            DoEvents
            Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE USUARIO_LOG = '" & strUserR & "' and LEFT(DATA_LOG,10) = '" & strDataR & "'")
            If rs.EOF = False Then
                'IMPRIME O TOTAL DIA DO USUÁRIO
                Print #2, strDataR & ";" & strUserR & ";" & rs!linhas
            End If
                rsMeta.MoveNext
                ProgressBar2.value = ProgressBar2.value + 1
            If rsMeta.EOF = False Then
                If rsMeta!data <> strDataR Then
                    'IMPRIME O TOTAL GERAL DIA
                    Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE LEFT(DATA_LOG,10) = '" & strDataR & "'")
                    Print #2, strDataR & ";" & "Total do Dia" & ";" & rs!linhas
                    strDataR = rsMeta!data
                End If
                strUserR = rsMeta!USUARIO_LOG
            Else 'CHEGOU AO FIM DA TABELA
                'IMPRIME O TOTAL GERAL DO ULTIMO DIA DA TABELA
                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE LEFT(DATA_LOG,10) = '" & strDataR & "'")
                Print #2, strDataR & ";Total do dia;" & rs!linhas
                Print #2, ""
                Print #2, "Obs: este relatório apresenta apenas as ligações de água"
                Print #2, "cadastradas a partir do GeoSan versão 7.5.0"
                Print #2, ""
            End If
        Loop
    Else
        Print #2, "NÃO HÁ INFORMAÇÕES PARA HISTÓRICO DIÁRIO DE USUÁRIO ****"
        Print #2, ""
    End If
    Print #2, "HISTÓRICO DIÁRIO DE USUÁRIO SEPARADO POR ; ********* FIM"
    Print #2, "********************************************************"
    Print #2, ""
    Print #2, "Obs: este relatório apresenta apenas as ligações de água"
    Print #2, "cadastradas a partir do GeoSan versão 7.5.0"
    Print #2, ""
    Print #2, ""
    Print #2, ""
    Print #2, "****************** SISTEMA GEOSAN **********************"
    Print #2, "######### RELATÓRIO INDICATIVO DE PRODUTIVIDADE ########"
    Print #2, "FIM - ****************************** " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    Close #2
    rsMeta.Close
    rs.Close
    MousePointer = Default
    MsgBox "Arquivo gerado com sucesso!", vbInformation, "Indicador"
    Unload Me
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 55 Then
        Resume Next
    Else
        Close #2
        MousePointer = vbDefault
        PrintErro CStr(Me.Name), "cmdGerar.Click ", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub

Private Sub Form_Load()
    txtCaminho.Text = App.path & "\Indicador_RamaisAgua_" & Format(Now, "YYYYMMDD") & ".txt"
End Sub


