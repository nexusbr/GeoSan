VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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

Private Sub cmdGerar_Click()
On Error GoTo Trata_Erro
MousePointer = vbHourglass
Dim rs As ADODB.Recordset
Dim rsMeta As ADODB.Recordset
Dim strDataR, strUserR As String
Dim contBar As Long
Dim strsql As String
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
Dim n As String
Dim o As String

a = "RAMAIS_AGUA"
b = "USUARIO_LOG"
c = "DATA_LOG"


   If frmCanvas.TipoConexao <> 4 Then
   Conn.execute ("UPDATE RAMAIS_AGUA SET USUARIO_LOG = 'DESCONHECIDO' WHERE USUARIO_LOG is null")
   Conn.execute ("UPDATE RAMAIS_AGUA SET DATA_LOG = '01/01/01 01:01' WHERE DATA_LOG is null")
    Else
  Conn.execute ("UPDATE " + """" + a + """" + " SET " + """" + b + """" + " = 'DESCONHECIDO' WHERE " + """" + b + """" + "is null")
   Conn.execute ("UPDATE " + """" + a + """" + " SET " + """" + c + """" + " = '01/01/01 01:01' WHERE " + """" + c + """" + " is null")
    End If
    strDataR = Format(Now, "DD/MM/YY")
    
    'IMPRIME O RELATÓRIO DO DIA QUE É DEFINIDO PELA DATA DA MAQUINA
    ProgressBar1.value = 2
    
    If frmCanvas.TipoConexao = 1 Then 'SQL
        
        strsql = "SELECT COUNT(*) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE LEFT(DATA_LOG,8) = '" & strDataR & "')"

    ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE


        strsql = "SELECT COUNT(*) AS " + """" + "LINHAS" + """" + " FROM RAMAIS_AGUA_LIGACAO RAL WHERE EXISTS (SELECT OBJECT_ID_ FROM RAMAIS_AGUA RA WHERE SUBSTR(DATA_LOG,1,8) = '" & strDataR & "' AND RA.OBJECT_ID_ = RAL.OBJECT_ID_)"
    ElseIf frmCanvas.TipoConexao = 4 Then
a = "RAMAIS_AGUA_LIGACAO"
b = "OBJECT_ID_"
c = "RAMAIS_AGUA"
d = Left(DATA_LOG, 8)

     strsql = "SELECT COUNT(*) AS " + """" + "LINHAS" + """" + " FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + " WHERE SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8)" + "=" + " '" + strDataR + "')"

    End If
    
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
        Print #2, strDataR; Tab(15); "Total Data"; Tab(30); rs!linhas
        Print #2, ""
        Print #2, "****************** RESUMO DO DIA ******************* FIM"
        Print #2, "********************************************************"
        Print #2, ""
        Print #2, ""
        Print #2, ""
    End If
    Close #2
    rs.Close


    'MONTAGEM DO RELATÓRIO DIA A DIA
    '1 - SELECT DISTINCT LEFT(DATA_LOG,8)as data,USUARIO_LOG FROM WATERLINES ORDER BY DATA,USUARIO_LOG
    '2 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = 'Adm' and LEFT(DATA_LOG,8) = '01/12/08'
    '3 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE LEFT(DATA_LOG,8) = '01/12/08'
    
    
   If frmCanvas.TipoConexao = 1 Then 'SQL
      'Set rsMeta = Conn.execute("SELECT DISTINCT LEFT(DATA_LOG,8) AS DATA,LEFT(LEFT(DATA_LOG,8),2) AS DIA,RIGHT(LEFT(DATA_LOG,5),2) AS MES,RIGHT(LEFT(DATA_LOG,8),2) AS ANO,USUARIO_LOG FROM RAMAIS_AGUA WHERE LEN(USUARIO_LOG) > 0 AND LEN(DATA_LOG) > 0 ORDER BY ANO,MES,DIA")
   
      strsql = "SELECT RA.USUARIO_LOG,"
      strsql = strsql & " LEFT(LEFT(RA.DATA_LOG,8),2) AS DIA,"
      strsql = strsql & " RIGHT(LEFT(RA.DATA_LOG,5),2) AS MES,"
      strsql = strsql & " RIGHT(LEFT(RA.DATA_LOG,8),2) AS ANO,"
      strsql = strsql & " LEFT(RA.DATA_LOG,8) AS DATA,"
      strsql = strsql & " COUNT(RAL.Object_id_) As Ligacoes"
      strsql = strsql & " FROM RAMAIS_AGUA RA JOIN RAMAIS_AGUA_LIGACAO RAL ON RAL.OBJECT_ID_ = RA.OBJECT_ID_"
      strsql = strsql & " Where Len(RA.USUARIO_LOG) > 0 And Len(RA.DATA_LOG) > 0"
      strsql = strsql & " GROUP BY RA.USUARIO_LOG, LEFT(RA.DATA_LOG,8),LEFT(LEFT(RA.DATA_LOG,8),2), RIGHT(LEFT(RA.DATA_LOG,5),2), RIGHT(LEFT(RA.DATA_LOG,8),2)"
      strsql = strsql & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
   
   ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
      'Set rsMeta = Conn.execute("SELECT DISTINCT SUBSTR(DATA_LOG,1,8) AS DATA, SUBSTR(DATA_LOG,1,2) AS DIA,SUBSTR(DATA_LOG,4,2) AS MES,SUBSTR(DATA_LOG,7,2) AS ANO, USUARIO_LOG FROM RAMAIS_AGUA WHERE LENGTH(USUARIO_LOG) > 0 AND LENGTH(DATA_LOG) > 0 ORDER BY ANO,MES,DIA")
   
      strsql = "SELECT RA.USUARIO_LOG,"
      strsql = strsql & " SUBSTR(DATA_LOG,1,2) AS " + """" + "DIA" + """" + ","
      strsql = strsql & " SUBSTR(DATA_LOG,4,2) AS " + """" + "MES" + """" + ","
      strsql = strsql & " SUBSTR(DATA_LOG,7,2) AS " + """" + "ANO" + """" + ","
      strsql = strsql & " SUBSTR(DATA_LOG,1,8) AS " + """" + "DATA" + """" + ","
      strsql = strsql & " COUNT(RAL.Object_id_) As " + """" + "Ligacoes" + """" + ""
      strsql = strsql & " FROM RAMAIS_AGUA RA JOIN RAMAIS_AGUA_LIGACAO RAL ON RAL.OBJECT_ID_ = RA.OBJECT_ID_"
      strsql = strsql & " Where Length(RA.USUARIO_LOG) > 0 And Length(RA.DATA_LOG) > 0"
      strsql = strsql & " GROUP BY RA.USUARIO_LOG, SUBSTR(RA.DATA_LOG,1,2), SUBSTR(RA.DATA_LOG,1,8), SUBSTR(RA.DATA_LOG,4,2), SUBSTR(RA.DATA_LOG,4,2), SUBSTR(RA.DATA_LOG,7,2)"
      strsql = strsql & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
   
   
   ElseIf frmCanvas.TipoConexao = 4 Then
      'Set rsMeta = Conn.execute("SELECT DISTINCT SUBSTR(DATA_LOG,1,8) AS DATA, SUBSTR(DATA_LOG,1,2) AS DIA,SUBSTR(DATA_LOG,4,2) AS MES,SUBSTR(DATA_LOG,7,2) AS ANO, USUARIO_LOG FROM RAMAIS_AGUA WHERE LENGTH(USUARIO_LOG) > 0 AND LENGTH(DATA_LOG) > 0 ORDER BY ANO,MES,DIA")
Dim a1 As String
Dim a2 As String
Dim a3 As String
Dim a4 As String
Dim a5 As String
Dim a6 As String
Dim a7 As String
Dim a8 As String
Dim a9 As String
Dim a10 As String
a = "USUARIO_LOG"
a10 = "DATA_LOG"
b = """" + "DATA_LOG" + """"

f = "OBJECT_ID_"
g = "RAMAIS_AGUA"
h = "RAMAIS_AGUA_LIGACAO"
i = "DATA_LOG"


a5 = "j"
a6 = "k"
a7 = "l"
a8 = "m"
a9 = "n"

'"SUBSTR(" + """" + a10 + """" + ",1,2) AS"
      strsql = "SELECT " + """" + g + """" + "." + """" + a + """" + ","
      strsql = strsql + "SUBSTR(" + """" + a10 + """" + ",1,2) AS " + """" + "DIA" + """" + ","
      strsql = strsql + "SUBSTR(" + """" + a10 + """" + ",4,2) AS " + """" + "MES" + """" + ","
      strsql = strsql + "SUBSTR(" + """" + a10 + """" + ",7,2) AS " + """" + "ANO" + """" + ","
      strsql = strsql + "SUBSTR(" + """" + a10 + """" + ",1,8) AS " + """" + "DATA" + """" + ","
      strsql = strsql & " COUNT(" + """" + h + """" + "." + """" + f + """" + ") As " + """" + "Ligacoes" + """" + ""
      strsql = strsql & " FROM " + """" + g + """" + " JOIN " + """" + h + """" + "  ON " + """" + h + """" + "." + """" + f + """" + " = " + """" + g + """" + "." + """" + f + """" + ""
      strsql = strsql & " Where length(" + """" + g + """" + "." + """" + a + """" + ") > '0' And length(" + """" + g + """" + "." + """" + a10 + """" + ") > '0'"
      strsql = strsql & " GROUP BY " + """" + g + """" + "." + """" + a + """" + "," + "SUBSTR(" + """" + a10 + """" + ",1,2)" + "," + "SUBSTR(" + """" + a10 + """" + ",1,8)" + "," + "SUBSTR(" + """" + a10 + """" + ",4,2)" + "," + "SUBSTR(" + """" + a10 + """" + ",4,2)" + "," + "SUBSTR(" + """" + a10 + """" + ",7,2)" + ""
      strsql = strsql & " ORDER BY " + """" + "ANO" + """" + "," + """" + "MES" + """" + "," + """" + "DIA" + """" + "," + """" + a + """" + ""
     '  WritePrivateProfileString "A", "A", strsql, App.path & "\DEBUG.INI"
      
   End If
    
   Set rs = New ADODB.Recordset
     rs.Open strsql, Conn, adOpenDynamic, adLockOptimistic
   Dim dataOld As String
   Dim SumRamais As Long
     
   SumRamais = 0
   Open txtCaminho.Text For Append As #2
   Print #2, "********************************************************"
   Print #2, "********** HISTÓRICO DIÁRIO DE USUÁRIO ********** INÍCIO"
   
   Print #2, "========================================================"
   Print #2, "DATA"; Tab(15); "USUARIO"; Tab(30); "LIGAÇÕES"
   Print #2, "========================================================"
     
   If rs.EOF = False Then
       dataOld = rs!Data
       Do While Not rs.EOF
          'IMPRIME O TOTAL DIA DO USUÁRIO
          If dataOld = rs!Data Then
             
             SumRamais = SumRamais + rs!Ligacoes
             
             Print #2, Trim(rs!Data); Tab(15); Trim(rs!USUARIO_LOG); Tab(30); Trim(rs!Ligacoes)
             
          Else ' TROCOU DE DATA
          
               Print #2, "========================================================"
               Print #2, dataOld; Tab(15); "Total Data"; Tab(30); SumRamais
               Print #2, ""
               Print #2, ""
             
               SumRamais = rs!Ligacoes
               
               Print #2, rs!Data; Tab(15); Trim(rs!USUARIO_LOG); Tab(30); Trim(rs!Ligacoes)
               
          End If
          dataOld = rs!Data
          rs.MoveNext
       
       Loop
       Print #2, "========================================================"
       Print #2, dataOld; Tab(15); "Total Data"; Tab(30); Trim(SumRamais)
       Print #2, ""

   Else
      
      Print #2, "NÃO HÁ INFORMAÇÕES PARA HISTÓRICO DIÁRIO DE USUÁRIO ****"
      Print #2, ""
   
   End If
        
   Print #2, "*********** HISTÓRICO DIÁRIO DE USUÁRIO ************ FIM"
   Print #2, "********************************************************"
   Print #2, ""
   Print #2, ""
   Print #2, ""
    
    
   
'''   If rsMeta.EOF = False Then
'''      Do While Not rsMeta.EOF = True
'''         rsMeta.MoveNext
'''         contBar = contBar +""""+ 1
'''      Loop
'''   End If
'''   ProgressBar2.value = 0
'''   ProgressBar2.Max = contBar +""""+ 5
'''   rsMeta.Requery
'''   ProgressBar1.value = 4
'''
'''    If rsMeta.EOF = False Then
'''
'''        strDataR = rsMeta!Data
'''        strUserR = rsMeta!usuario_log
'''
'''        Do While Not rsMeta.EOF = True
'''            DoEvents
'''            If frmCanvas.TipoConexao = 1 Then 'SQL
'''                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE USUARIO_LOG = '" & strUserR & "' and LEFT(DATA_LOG,8) = '" & strDataR & "')")
'''            ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
'''                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE USUARIO_LOG = '" & strUserR & "' and SUBSTR(DATA_LOG,1,8) = '" & strDataR & "')")
'''            End If
'''
'''            If rs.EOF = False Then
'''                'IMPRIME O TOTAL DIA DO USUÁRIO
'''                Print #2, strDataR; Tab(15); strUserR; Tab(30); rs!linhas
'''            End If
'''            rsMeta.MoveNext
'''            ProgressBar2.value = ProgressBar2.value +""""+ 1
'''            If rsMeta.EOF = False Then
'''                If rsMeta!Data <> strDataR Then
'''                    'IMPRIME O TOTAL GERAL DIA
'''
'''                    If frmCanvas.TipoConexao = 1 Then 'SQL
'''                        Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE LEFT(DATA_LOG,8) = '" & strDataR & "')")
'''                    ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
'''                        Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE SUBSTR(DATA_LOG,1,8) = '" & strDataR & "')")
'''                    End If
'''
'''                    Print #2, "========================================================"
'''                    Print #2, strDataR; Tab(15); "Total Data"; Tab(30); rs!linhas
'''                    Print #2, ""
'''                    Print #2, ""
'''                    strDataR = rsMeta!Data
'''                End If
'''                strUserR = rsMeta!usuario_log
'''            Else 'CHEGOU AO FIM DA TABELA
'''                 'IMPRIME O TOTAL GERAL DO ULTIMO DIA DA TABELA
'''
'''                If frmCanvas.TipoConexao = 1 Then 'SQL
'''                    Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE LEFT(DATA_LOG,8) = '" & strDataR & "')")
'''                ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
'''                    Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE SUBSTR(DATA_LOG,1,8) = '" & strDataR & "')")
'''                End If
'''
'''                Print #2, "========================================================"
'''                Print #2, strDataR; Tab(15); "Total Data"; Tab(30); rs!linhas
'''                Print #2, ""
'''            End If
'''        Loop
'''    Else
'''        Print #2, "NÃO HÁ INFORMAÇÕES PARA HISTÓRICO DIÁRIO DE USUÁRIO ****"
'''        Print #2, ""
'''    End If
'''    Print #2, "*********** HISTÓRICO DIÁRIO DE USUÁRIO ************ FIM"
'''    Print #2, "********************************************************"
'''    Print #2, ""
'''    Print #2, ""
'''    Print #2, ""
    
    
    
    
    Close #2
    
    'MONTAGEM DO RELATÓRIO RESUMO CONSOLIDADO (ACUMULADO) DE USUÁRIO
    '1 - SELECT DISTINCT LEFT(DATA_LOG,8)as data,USUARIO_LOG FROM WATERLINES ORDER BY DATA,USUARIO_LOG
    '2 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = 'Jonathas'
    '3 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES
a = "USUARIO_LOG"
b = "RAMAIS_AGUA"

    If frmCanvas.TipoConexao = 1 Then 'SQL
        Set rsMeta = Conn.execute("SELECT DISTINCT USUARIO_LOG FROM RAMAIS_AGUA WHERE LEN(USUARIO_LOG) > 0 ORDER BY USUARIO_LOG")
    ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
        Set rsMeta = Conn.execute("SELECT DISTINCT USUARIO_LOG FROM RAMAIS_AGUA WHERE LENGTH(USUARIO_LOG) > 0 ORDER BY USUARIO_LOG")
        ElseIf frmCanvas.TipoConexao = 4 Then
        Set rsMeta = Conn.execute("SELECT DISTINCT " + """" + "USUARIO_LOG" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + " WHERE LENgth(" + """" + "USUARIO_LOG" + """" + ") > 0 ORDER BY " + """" + "USUARIO_LOG" + """" + "")
        
    End If
   
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
    Print #2, "********* RESUMO CONSOLIDADO DE USUÁRIO ********* INÍCIO"
    
    If rsMeta.EOF = False Then

        strUserR = rsMeta!USUARIO_LOG
        Print #2, "========================================================"
        Print #2, ""; Tab(15); "USUARIO"; Tab(30); "LIGAÇÕES"
        Print #2, "========================================================"
        Do While Not rsMeta.EOF = True
        DoEvents
a = "NRO_LIGACAO"
b = "RAMAIS_AGUA_LIGACAO"
c = "OBJECT_ID_"
d = "RAMAIS_AGUA"
e = "USUARIO_LOG"

            If frmCanvas.TipoConexao <> 4 Then
            Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS " + """" + "LINHAS" + """" + " FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE USUARIO_LOG = '" & strUserR & "')")
            Else
            Set rs = Conn.execute("SELECT COUNT(" + """" + "NRO_LIGACAO" + """" + ") AS LINHAS FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + " WHERE " + """" + "USUARIO_LOG" + """" + " = '" & strUserR & "')")
            End If
            If rs.EOF = False Then
                'IMPRIME O TOTAL DIA DO USUÁRIO
                Print #2, ""; Tab(15); strUserR; Tab(30); rs!linhas
            End If
            rsMeta.MoveNext
            ProgressBar2.value = ProgressBar2.value + 1
            If rsMeta.EOF = False Then
                strUserR = rsMeta!USUARIO_LOG
            Else
                'IMPRIME O TOTAL GERAL DA BASE DE DADOS
a = "NRO_LIGACAO"
b = "RAMAIS_AGUA_LIGACAO"
c = "OBJECT_ID_"
d = "RAMAIS_AGUA"
e = "USUARIO_LOG"

                If frmCanvas.TipoConexao <> 4 Then
                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS " + """" + "LINHAS" + """" + " FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA)")
                Else
                Set rs = Conn.execute("SELECT COUNT(" + """" + "NRO_LIGACAO" + """" + ") AS " + """" + "LINHAS" + """" + " FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + ")")
                End If
                Print #2, ""
                Print #2, "*********** RESUMO CONSOLIDADO DE USUÁRIO ********** FIM"
                Print #2, "********************************************************"
                Print #2, ""
                Print #2, ""
                Print #2, ""
                Print #2, "TOTAL GERAL"; Tab(30); "LIGAÇÕES"
                Print #2, "========================================================"
                Print #2, "ATÉ " & Format(Now, "DD/MM/YYYY HH:MM:SS"); Tab(30); rs!linhas
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
    
    
    'MONTAGEM DO RELATÓRIO DIA A DIA SEPARADO POR PONTO E VIRGULA
    '1 - SELECT DISTINCT LEFT(DATA_LOG,8)as data,USUARIO_LOG FROM WATERLINES ORDER BY DATA,USUARIO_LOG
    '2 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = 'Adm' and LEFT(DATA_LOG,8) = '01/12/08'
    '3 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE LEFT(DATA_LOG,8) = '01/12/08'
    'Set rsMeta = Conn.execute("SELECT DISTINCT LEFT(DATA_LOG,8) AS DATA,USUARIO_LOG FROM WATERLINES ORDER BY DATA,USUARIO_LOG")
    
a = Left(DATA_LOG, 8)

c = Left(Left(DATA_LOG, 8), 2)

e = Right(Left(DATA_LOG, 5), 2)

g = "USUARIO_LOG"
h = "RAMAIS_AGUA"
i = "DATA_LOG"
j = Right(Left(DATA_LOG, 8), 2)

Dim g1 As String
Dim g2 As String

Dim g3 As String
g1 = "ANO"
g2 = "MES"
g3 = "DIA"


    If frmCanvas.TipoConexao = 1 Then 'SQL
        Set rsMeta = Conn.execute("SELECT DISTINCT LEFT(DATA_LOG,8) AS DATA,LEFT(LEFT(DATA_LOG,8),2) AS DIA,RIGHT(LEFT(DATA_LOG,5),2) AS MES,RIGHT(LEFT(DATA_LOG,8),2) AS ANO,USUARIO_LOG FROM RAMAIS_AGUA WHERE LEN(USUARIO_LOG) > 0 AND LEN(DATA_LOG) > 0 ORDER BY ANO,MES,DIA")
    ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
        Set rsMeta = Conn.execute("SELECT DISTINCT SUBSTR(DATA_LOG,1,8) AS " + """" + "DATA" + """" + ", SUBSTR(DATA_LOG,1,2) AS " + """" + "DIA" + """" + ",SUBSTR(DATA_LOG,4,2) AS " + """" + "MES" + """" + ",SUBSTR(DATA_LOG,7,2) AS " + """" + "ANO" + """" + ", USUARIO_LOG FROM RAMAIS_AGUA WHERE LENGTH(USUARIO_LOG) > 0 AND LENGTH(DATA_LOG) > 0 ORDER BY ANO,MES,DIA")
        ElseIf frmCanvas.TipoConexao = 4 Then
         Set rsMeta = Conn.execute("SELECT DISTINCT " + "SUBSTR(" + """" + "DATA_LOG" + """" + ", 1, 8)" + " AS " + """" + "DATA" + """" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ", 1, 2)" + " AS " + """" + "DIA" + """" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ", 4, 2)" + " AS " + """" + "MES" + """" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ", 7, 2)" + " AS " + """" + "ANO" + """" + "," + """" + g + """" + " FROM " + """" + h + """" + " WHERE LENgth(" + """" + g + """" + ") > '0' AND LENgth(" + """" + i + """" + ") > '0' ORDER BY " + """" + "ANO" + """" + "," + """" + "MES" + """" + "," + """" + "ANO" + """" + "")
    End If
    
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
    Print #2, ""
    If rsMeta.EOF = False Then

        strDataR = rsMeta!Data
        strUserR = rsMeta!USUARIO_LOG
        
        Print #2, "DATA;USUARIO;LIGAÇÕES"
        Do While Not rsMeta.EOF = True
            DoEvents
            
a = NRO_LIGACAO
b = "RAMAIS_AGUA_LIGACAO"
c = "OBJECT_ID_"
d = "RAMAIS_AGUA"
g = "USUARIO_LOG"
h = "RAMAIS_AGUA"
i = "DATA_LOG"
j = Left(DATA_LOG, 8)


            If frmCanvas.TipoConexao = 1 Then 'SQL
                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE USUARIO_LOG = '" & strUserR & "' and LEFT(DATA_LOG,8) = '" & strDataR & "')")
            ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
                Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS " + """" + "LINHAS" + """" + " FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE USUARIO_LOG = '" & strUserR & "' and SUBSTR(DATA_LOG,1,8) = '" & strDataR & "')")
             ElseIf frmCanvas.TipoConexao = 4 Then
            Set rs = Conn.execute("SELECT COUNT(" + """" + "NRO_LIGACAO" + """" + ") AS " + """" + "LINHAS" + """" + " FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + " WHERE " + """" + "USUARIO_LOG" + """" + " = '" & strUserR & "' and " + "SUBSTR(" + """" + "DATA_LOG" + """" + ", 1, 8)" + " = '" & strDataR & "')")
    End If
            
            If rs.EOF = False Then
                'IMPRIME O TOTAL DIA DO USUÁRIO
                Print #2, strDataR & ";" & strUserR & ";" & rs!linhas
            End If
            rsMeta.MoveNext
            ProgressBar2.value = ProgressBar2.value + 1
            If rsMeta.EOF = False Then
                If rsMeta!Data <> strDataR Then
                    'IMPRIME O TOTAL GERAL DIA
                    
                    If frmCanvas.TipoConexao = 1 Then 'SQL
                        Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE LEFT(DATA_LOG,8) = '" & strDataR & "')")
                    ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
                        Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS " + """" + "LINHAS" + """" + " FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE SUBSTR(DATA_LOG,1,8) = '" & strDataR & "')")
                   ElseIf frmCanvas.TipoConexao = 4 Then
            Set rs = Conn.execute("SELECT COUNT(" + """" + "NRO_LIGACAO" + """" + ") AS " + """" + "LINHAS" + """" + " FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + " WHERE " + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8)" + " = '" & strDataR & "')")
    End If
                    Print #2, strDataR & ";" & "Total Data" & ";" & rs!linhas

                    strDataR = rsMeta!Data
                End If
                strUserR = rsMeta!USUARIO_LOG
            Else 'CHEGOU AO FIM DA TABELA
                 'IMPRIME O TOTAL GERAL DO ULTIMO DIA DA TABELA
                 
                If frmCanvas.TipoConexao = 1 Then 'SQL
                    Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS LINHAS FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE LEFT(DATA_LOG,8) = '" & strDataR & "')")
                ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
                    Set rs = Conn.execute("SELECT COUNT(NRO_LIGACAO) AS " + """" + "LINHAS" + """" + " FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE SUBSTR(DATA_LOG,1,8) = '" & strDataR & "')")
               ElseIf frmCanvas.TipoConexao = 4 Then
            Set rs = Conn.execute("SELECT COUNT(" + """" + "NRO_LIGACAO" + """" + ") AS " + """" + "LINHAS" + """" + " FROM " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + " IN (SELECT " + """" + "OBJECT_ID_" + """" + " FROM " + """" + "RAMAIS_AGUA" + """" + " WHERE " + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8)" + " = '" & strDataR & "')")
    End If

                Print #2, strDataR & ";Total Data;" & rs!linhas
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


