Attribute VB_Name = "mdlRelatoriosProdutividade"
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

Public Function RelProdutividade(TipoRede As String) As Boolean

If TipoRede = "ESGOTO" Then
   TipoRede = "SEWERLINES"
Else
   TipoRede = "WATERLINES"
End If

On Error GoTo Trata_Erro
    MousePointer = vbHourglass
    Dim rs As ADODB.Recordset
    Dim rsMeta As ADODB.Recordset
    Dim strDataR, strUserR As String
    
    strDataR = Format(Now, "DD/MM/YY")
    
    'IMPRIME O RELATÓRIO DO DIA QUE É DEFINIDO PELA DATA DA MAQUINA
   
   frmIndicProdutRedesDeAgua.ProgBar1.Visible = True
   frmIndicProdutRedesDeAgua.ProgBar1.value = 1
   DoEvents
   
   
a = TipoRede
b = "a"
c = "USUARIO_LOG"
d = "DATA_LOG"


    If frmCanvas.TipoConexao <> 4 Then
   Conn.execute ("UPDATE " & TipoRede & " SET USUARIO_LOG = 'DESCONHECIDO' WHERE USUARIO_LOG is null")
   Conn.execute ("UPDATE " & TipoRede & " SET DATA_LOG = '01/01/01 01:01' WHERE DATA_LOG is null")
   Else

   Conn.execute ("UPDATE " + """" + TipoRede + """" + " SET " + """" + c + """" + " = 'DESCONHECIDO' WHERE " + """" + c + """" + " is null")
   Conn.execute ("UPDATE " + """" + TipoRede + """" + " SET " + """" + d + """" + " = '01/01/01 01:01' WHERE " + """" + d + """" + " is null")
   End If
   
   If frmCanvas.TipoConexao = 1 Then 'SQL
       Set rs = Conn.execute("SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM " & TipoRede & " WHERE LEFT(DATA_LOG,8) = '" & strDataR & "'")

   ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
       
       Set rs = Conn.execute("SELECT COUNT(*) AS " + """" + "LINHAS" + """" + ",SUM(LENGTHCALCULATED) AS " + """" + "COMPRIMENTO" + """" + " FROM " & TipoRede & " WHERE SUBSTR(DATA_LOG,1,8) = '" & strDataR & "'")
   ElseIf frmCanvas.TipoConexao = 4 Then 'POSTGRES
   
a = TipoRede
c = "LENGTHCALCULATED"
d = "DATA_LOG"
'e = Left("+""""+"DATA_LOG", 8)

'b = "SELECT COUNT(*) AS " + """" + "LINHAS" + """" + ",SUM(" + """" + "LENGTHCALCULATED" + """" + ") AS " + """" + "COMPRIMENTO" + """" + " FROM " + """" + TipoRede + """" + " WHERE " + """" + Left(DATA_LOG, 8) + """" + " = '" + strDataR + "'"
'MsgBox b

 'WritePrivateProfileString "A", "A", b, App.path & "\DEBUG.INI"
   
   Set rs = Conn.execute("SELECT COUNT(*) AS " + """" + "LINHAS" + """" + ",SUM(" + """" + "LENGTHCALCULATED" + """" + ") AS " + """" + "COMPRIMENTO" + """" + " FROM " + """" + TipoRede + """" + " WHERE  SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8)   = '" + strDataR + "'")

   End If
   
   Open frmIndicProdutRedesDeAgua.txtCaminho.Text For Output As #2
       
   Print #2, "****************** SISTEMA GEOSAN **********************"
   Print #2, "######### RELATÓRIO INDICATIVO DE PRODUTIVIDADE ########"
    
   If TipoRede = "SEWERLINES" Then
      Print #2, "############## DESENHO DE REDES DE ESGOTO ##############"
   Else
      Print #2, "############### DESENHO DE REDES DE AGUA ###############"
   End If
   
    Print #2, "INÍCIO - *************************** " & Format(Now, "DD/MM/YYYY HH:MM:SS")
        
    Print #2, ""
    Print #2, ""
    If rs.EOF = False Then
        
        Print #2, "********************************************************"
        Print #2, "****************** RESUMO DO DIA *****************INÍCIO"
        Print #2, ""
        Print #2, "DATA"; Tab(30); "LINHAS"; Tab(45); "COMPRIMENTO"
        Print #2, "========================================================"
        Print #2, strDataR; Tab(15); "Total Data"; Tab(30); rs!linhas; Tab(45); Format(rs!comprimento, "0.00")
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
   
   
     frmIndicProdutRedesDeAgua.ProgBar1.value = 2
     DoEvents
   
     Open frmIndicProdutRedesDeAgua.txtCaminho.Text For Append As #2
     Print #2, "********************************************************"
     Print #2, "********** HISTÓRICO DIÁRIO DE USUÁRIO ********** INÍCIO"
      
     Print #2, "========================================================"
     Print #2, "DATA"; Tab(15); "USUARIO"; Tab(30); "LINHAS"; Tab(45); "COMPRIMENTO"
     Print #2, "========================================================"
     Dim str As String
     Set rs = New ADODB.Recordset

     If frmCanvas.TipoConexao = 1 Then 'SQL
         
        'Set rs = Conn.execute("SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = '" & strUserR & "' and LEFT(DATA_LOG,8) = '" & strDataR & "'")
        str = "SELECT USUARIO_LOG,"
        str = str & "LEFT(LEFT(DATA_LOG,8),2) AS DIA,"
        str = str & "RIGHT(LEFT(DATA_LOG,5),2) AS MES,"
        str = str & "RIGHT(LEFT(DATA_LOG,8),2) AS ANO,"
        str = str & "LEFT(DATA_LOG,8) AS DATA,"
        str = str & "COUNT(*) AS LINHAS,"
        str = str & "SUM(LENGTHCALCULATED) As comprimento"
        str = str & " FROM " & TipoRede
        str = str & " WHERE Len(USUARIO_LOG) > 0 And Len(DATA_LOG) > 0"
        str = str & " GROUP BY USUARIO_LOG,LEFT(LEFT(DATA_LOG,8),2),LEFT(DATA_LOG,8),RIGHT(LEFT(DATA_LOG,5),2),RIGHT(LEFT(DATA_LOG,8),2)"
        str = str & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
        
     ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
         
        'Set rs = Conn.execute("SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = '" & strUserR & "' and SUBSTR(DATA_LOG,1,8) = '" & strDataR & "'")
        str = "SELECT USUARIO_LOG,"
        str = str & " SUBSTR(DATA_LOG,1,2) AS " + """" + "DIA" + """" + ","
        str = str & " SUBSTR(DATA_LOG,4,2) AS " + """" + "MES" + """" + ","
        str = str & " SUBSTR(DATA_LOG,7,2) AS " + """" + "ANO" + """" + ","
        str = str & " SUBSTR(DATA_LOG,1,8) AS " + """" + "DATA" + """" + ","
        str = str & " COUNT(*) AS LINHAS,"
        str = str & " SUM(LENGTHCALCULATED) As " + """" + "comprimento" + """" + ""
        str = str & " From " & TipoRede
        str = str & " GROUP BY USUARIO_LOG,SUBSTR(DATA_LOG,1,2),SUBSTR(DATA_LOG,1,8),SUBSTR(DATA_LOG,4,2),SUBSTR(DATA_LOG,4,2),SUBSTR(DATA_LOG,7,2)"
        str = str & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
     
       ElseIf frmCanvas.TipoConexao = 4 Then 'Postgres
     Dim ut As String
     c = "LENGTHCALCULATED"
d = "DATA_LOG"
ut = "USUARIO_LOG"
a = "ANO"
b = "MES"
e = "DIA"
f = "USUARIO_LOG"

         str = "SELECT " + """" + "USUARIO_LOG" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,2) AS" + """" + "DIA" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",4,2) AS" + """" + "MES" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",7,2) AS " + """" + "ANO" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8) AS " + """" + "DATA" + """" + ","
        str = str & "COUNT(*) AS " + """" + "LINHAS" + """" + ","
        str = str & "SUM(" + """" + "LENGTHCALCULATED" + """" + ") As " + """" + "Comprimento" + """" + ""
        str = str & " FROM " + """" + TipoRede + """" + ""
        str = str & " WHERE " + "length(" + """" + "USUARIO_LOG" + """" + ")" + " > '0'" + " And " + "length(" + """" + "DATA_LOG" + """" + ")" + " > '0'"
        str = str & " GROUP BY " + """" + "USUARIO_LOG" + """" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,2)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",4,2)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",4,2)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",7,2)" + ""
        str = str & " ORDER BY " + """" + a + """" + "," + """" + b + """" + "," + """" + e + """" + "," + """" + f + """" + ""
     
     
     
     ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
     End If
     
        rs.Open str, Conn, adOpenDynamic, adLockOptimistic
     Dim dataOld As String
     Dim SumLinhas As Long, SumComp As Double
     
     SumLinhas = 0
     SumComp = 0
     
     If rs.EOF = False Then
         dataOld = rs!Data
         Do While Not rs.EOF
            'IMPRIME O TOTAL DIA DO USUÁRIO
            If dataOld = rs!Data Then
               
               SumLinhas = SumLinhas + rs!linhas
               SumComp = SumComp + rs!comprimento
               
               Print #2, Trim(rs!Data); Tab(15); Trim(rs!USUARIO_LOG); Tab(30); Trim(rs!linhas); Tab(45); Format(rs!comprimento, "0.00")
               
            Else ' TROCOU DE DATA
            
                 Print #2, "========================================================"
                 Print #2, dataOld; Tab(15); "Total Data"; Tab(30); SumLinhas; Tab(45); Format(SumComp, "0.00")
                 Print #2, ""
                 Print #2, ""
               
                 SumLinhas = rs!linhas
                 SumComp = rs!comprimento
                 
                 Print #2, rs!Data; Tab(15); Trim(rs!USUARIO_LOG); Tab(30); rs!linhas; Tab(45); Format(rs!comprimento, "0.00")
                 
            End If
            dataOld = rs!Data
            rs.MoveNext
         
         Loop
         Print #2, "========================================================"
         Print #2, dataOld; Tab(15); "Total Data"; Tab(30); SumLinhas; Tab(45); Format(SumComp, "0.00")
         Print #2, ""
         Print #2, "*********** HISTÓRICO DIÁRIO DE USUÁRIO ************ FIM"
         Print #2, "********************************************************"
         Print #2, ""
         Print #2, ""
         Print #2, ""
     
     Else

        Print #2, "NÃO HÁ INFORMAÇÕES PARA HISTÓRICO DIÁRIO DE USUÁRIO ****"
        Print #2, ""
        Print #2, "*********** HISTÓRICO DIÁRIO DE USUÁRIO ************ FIM"
        Print #2, "********************************************************"
        Print #2, ""
        Print #2, ""
        Print #2, ""
     
     End If
         
     Close #2 'fecha arquivo salvando

   frmIndicProdutRedesDeAgua.ProgBar1.value = 3
   DoEvents

    'MONTAGEM DO RELATÓRIO RESUMO CONSOLIDADO (ACUMULADO) DE USUÁRIO
    '1 - SELECT DISTINCT LEFT(DATA_LOG,8)as data,USUARIO_LOG FROM WATERLINES ORDER BY DATA,USUARIO_LOG
    '2 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = 'Jonathas'
    '3 - SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES
   
   Set rs = New ADODB.Recordset

   If frmCanvas.TipoConexao = 1 Then 'SQL
       'Set rsMeta = Conn.execute("SELECT DISTINCT USUARIO_LOG FROM WATERLINES WHERE LEN(USUARIO_LOG) > 0 ORDER BY USUARIO_LOG")
       
       rs.Open "SELECT USUARIO_LOG, COUNT(*) AS LINHAS, SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM " & TipoRede & " WHERE LEN(USUARIO_LOG) > 0 GROUP BY USUARIO_LOG ORDER BY USUARIO_LOG", Conn, adOpenForwardOnly, adLockReadOnly

   ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
       'Set rsMeta = Conn.execute("SELECT DISTINCT USUARIO_LOG FROM WATERLINES WHERE LENGTH(USUARIO_LOG) > 0 ORDER BY USUARIO_LOG")
       
       rs.Open "SELECT USUARIO_LOG, COUNT(*) AS " + """" + "LINHAS" + """" + ", SUM(LENGTHCALCULATED) AS " + """" + "COMPRIMENTO" + """" + " FROM " & TipoRede & " WHERE LENGTH(USUARIO_LOG) > 0 GROUP BY USUARIO_LOG ORDER BY USUARIO_LOG", Conn, adOpenForwardOnly, adLockReadOnly
   
    ElseIf frmCanvas.TipoConexao = 4 Then 'Postgres
   c = "LENGTHCALCULATED"
d = "DATA_LOG"
ut = "USUARIO_LOG"

   
   rs.Open "SELECT " + """" + ut + """" + ", COUNT(*) AS LINHAS, SUM(" + """" + c + """" + ") AS " + """" + "COMPRIMENTO" + """" + " FROM " + """" + TipoRede + """" + " WHERE LENGTH(" + """" + ut + """" + ") > '0' GROUP BY " + """" + ut + """" + " ORDER BY " + """" + ut + """" + "", Conn, adOpenDynamic, adLockOptimistic
   End If
   
   Close #2
   Open frmIndicProdutRedesDeAgua.txtCaminho.Text For Append As #2
   Print #2, "********************************************************"
   Print #2, "********* RESUMO CONSOLIDADO DE USUÁRIO ********* INÍCIO"
    

   Print #2, "========================================================"
   Print #2, ""; Tab(15); "USUARIO"; Tab(30); "LINHAS"; Tab(45); "COMPRIMENTO"
   Print #2, "========================================================"
         
   If rs.EOF = False Then
      Do While Not rs.EOF
         'IMPRIME O TOTAL GERAL DO USUÁRIO
         Print #2, ""; Tab(15); Trim(rs!USUARIO_LOG); Tab(30); rs!linhas; Tab(45); Format(rs!comprimento, "0.00")
         rs.MoveNext
      Loop
   
 
      Set rs = New ADODB.Recordset
      
c = "LENGTHCALCULATED"
d = "DATA_LOG"
ut = "USUARIO_LOG"
a = "ANO"
b = "MES"
e = "DIA"
f = "USUARIO_LOG"

       If frmCanvas.TipoConexao <> 4 Then 'Postgres
      rs.Open "SELECT COUNT(*) AS " + """" + "LINHAS" + """" + ",SUM(LENGTHCALCULATED) AS " + """" + "COMPRIMENTO" + """" + " FROM " & TipoRede, Conn, adOpenForwardOnly, adLockReadOnly
      Else
      Dim zaza As String
      zaza = "SELECT COUNT(*) AS " + """" + "LINHAS" + """" + ",SUM(" + """" + c + """" + ") AS " + """" + "COMPRIMENTO" + """" + " FROM " + """" + TipoRede + """" + ""
       rs.Open zaza, Conn, adOpenDynamic, adLockOptimistic
      End If
      'IMPRIME O TOTAL GERAL DA BASE DE DADOS
       Print #2, ""
       Print #2, "*********** RESUMO CONSOLIDADO DE USUÁRIO ********** FIM"
       Print #2, "********************************************************"
       Print #2, ""
       Print #2, ""
       Print #2, ""
       Print #2, "TOTAL GERAL"; Tab(30); "LINHAS"; Tab(45); "COMPRIMENTO"
       Print #2, "========================================================"
       Print #2, "ATÉ " & Format(Now, "DD/MM/YYYY HH:MM:SS"); Tab(30); rs!linhas; Tab(45); Format(rs!comprimento, "0.00")
       Print #2, ""
       Print #2, ""

   Else
      
      'RESUMO CONSOLIDADO DE USUÁRIO
      Print #2, "NÃO HÁ INFORMAÇÕES PARA RESUMO CONSOLIDADO DE USUÁRIO **"
      Print #2, ""
   
   End If
         
   Close #2
   Open frmIndicProdutRedesDeAgua.txtCaminho.Text For Append As #2
            
            

   frmIndicProdutRedesDeAgua.ProgBar1.value = 4
   DoEvents
    
    Print #2, "********************************************************"
    Print #2, "HISTÓRICO DIÁRIO DE USUÁRIO SEPARADO POR ; ****** INÍCIO"
    Print #2, ""
        
    Print #2, "DATA;USUARIO;LINHAS;COMPRIMENTO"
        
      
    If frmCanvas.TipoConexao = 1 Then 'SQL
         
      'Set rs = Conn.execute("SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = '" & strUserR & "' and LEFT(DATA_LOG,8) = '" & strDataR & "'")
      str = "SELECT USUARIO_LOG,"
      str = str & "LEFT(LEFT(DATA_LOG,8),2) AS DIA,"
      str = str & "RIGHT(LEFT(DATA_LOG,5),2) AS MES,"
      str = str & "RIGHT(LEFT(DATA_LOG,8),2) AS ANO,"
      str = str & "LEFT(DATA_LOG,8) AS DATA,"
      str = str & "COUNT(*) AS LINHAS,"
      str = str & "SUM(LENGTHCALCULATED) As comprimento"
      str = str & " FROM " & TipoRede
      str = str & " WHERE Len(USUARIO_LOG) > 0 And Len(DATA_LOG) > 0"
      str = str & " GROUP BY USUARIO_LOG,LEFT(LEFT(DATA_LOG,8),2),LEFT(DATA_LOG,8),RIGHT(LEFT(DATA_LOG,5),2),RIGHT(LEFT(DATA_LOG,8),2)"
      str = str & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
        
    ElseIf frmCanvas.TipoConexao = 2 Then 'ORACLE
         
      'Set rs = Conn.execute("SELECT COUNT(*) AS LINHAS,SUM(LENGTHCALCULATED) AS COMPRIMENTO FROM WATERLINES WHERE USUARIO_LOG = '" & strUserR & "' and SUBSTR(DATA_LOG,1,8) = '" & strDataR & "'")
      str = "SELECT USUARIO_LOG,"
      str = str & " SUBSTR(DATA_LOG,1,2) AS " + """" + "DIA" + """" + ","
      str = str & " SUBSTR(DATA_LOG,4,2) AS " + """" + "MES" + """" + ","
      str = str & " SUBSTR(DATA_LOG,7,2) AS " + """" + "ANO" + """" + ","
      str = str & " SUBSTR(DATA_LOG,1,8) AS " + """" + "DATA" + """" + ","
      str = str & " COUNT(*) AS " + """" + "LINHAS" + """" + ","
      str = str & " SUM(LENGTHCALCULATED) As " + """" + "comprimento" + """" + ""
      str = str & " From " & TipoRede
      str = str & " GROUP BY USUARIO_LOG,SUBSTR(DATA_LOG,1,2),SUBSTR(DATA_LOG,1,8),SUBSTR(DATA_LOG,4,2),SUBSTR(DATA_LOG,4,2),SUBSTR(DATA_LOG,7,2)"
      str = str & " ORDER BY ANO,MES,DIA,USUARIO_LOG"
     
   ElseIf frmCanvas.TipoConexao = 4 Then 'Postgres
  c = "LENGTHCALCULATED"
d = "DATA_LOG"
ut = "USUARIO_LOG"
a = "ANO"
b = "MES"
e = "DIA"
f = "USUARIO_LOG"

           str = "SELECT " + """" + "USUARIO_LOG" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,2) AS" + """" + "DIA" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",4,2) AS" + """" + "MES" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",7,2) AS " + """" + "ANO" + """" + ","
        str = str + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8) AS " + """" + "DATA" + """" + ","
        str = str & "COUNT(*) AS " + """" + "LINHAS" + """" + ","
        str = str & "SUM(" + """" + "LENGTHCALCULATED" + """" + ") As " + """" + "Comprimento" + """" + ""
        str = str & " FROM " + """" + TipoRede + """" + ""
        str = str & " WHERE " + "length(" + """" + "USUARIO_LOG" + """" + ")" + " > '0'" + " And" + " length(" + """" + "DATA_LOG" + """" + ")" + " > '0'"
        str = str & " GROUP BY " + """" + "USUARIO_LOG" + """" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,2)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",1,8)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",4,2)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",4,2)" + "," + "SUBSTR(" + """" + "DATA_LOG" + """" + ",7,2)" + ""
        str = str & " ORDER BY " + """" + a + """" + "," + """" + b + """" + "," + """" + e + """" + "," + """" + f + """" + ""
     
     End If
     ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
     
     

   Set rs = New ADODB.Recordset
   rs.Open str, Conn, adOpenDynamic, adLockOptimistic
   
   SumLinhas = 0
   SumComp = 0
   
   If rs.EOF = False Then
       dataOld = rs!Data
       Do While Not rs.EOF
          'IMPRIME O TOTAL DIA DO USUÁRIO
          If dataOld = rs!Data Then
             
             SumLinhas = SumLinhas + rs!linhas
             SumComp = SumComp + rs!comprimento
             
             Print #2, rs!Data & ";" & Trim(rs!USUARIO_LOG) & ";" & rs!linhas & ";" & Format(rs!comprimento, "0.00")
             
          Else ' TROCOU DE DATA
          
               Print #2, dataOld & ";" & "Total Data" & ";" & SumLinhas & ";" & Format(SumComp, "0.00")
             
               SumLinhas = rs!linhas
               SumComp = rs!comprimento
               
               Print #2, rs!Data & ";" & Trim(rs!USUARIO_LOG) & ";" & rs!linhas & ";" & Format(rs!comprimento, "0.00")
   
          End If
          dataOld = rs!Data
          rs.MoveNext
       
       Loop
       
       Print #2, dataOld & ";" & "Total Data" & ";" & SumLinhas & ";" & Format(SumComp, "0.00")
       Print #2, ""
   
   Else
   
      Print #2, "NÃO HÁ INFORMAÇÕES PARA HISTÓRICO DIÁRIO DE USUÁRIO ****"
      Print #2, ""
      Print #2, "*********** HISTÓRICO DIÁRIO DE USUÁRIO ************ FIM"
      Print #2, "********************************************************"
      Print #2, ""
      Print #2, ""
      Print #2, ""
   
   End If
   
   Print #2, "HISTÓRICO DIÁRIO DE USUÁRIO SEPARADO POR ; ********* FIM"
   Print #2, "********************************************************"
   Print #2, ""
   Print #2, ""
   Print #2, "****************** SISTEMA GEOSAN **********************"
   Print #2, "######### RELATÓRIO INDICATIVO DE PRODUTIVIDADE ########"
   Print #2, "FIM - ****************************** " & Format(Now, "DD/MM/YYYY HH:MM:SS")
      
    
   frmIndicProdutRedesDeAgua.ProgBar1.value = 5
    
   RelProdutividade = True
    
   Close #2
   rs.Close

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   Close #2
   MousePointer = vbDefault
   
   PrintErro "mdlRelatoriosProdutividade", "Public Function RelProdutividade(TipoRede As String) As Boolean", CStr(Err.Number), CStr(Err.Description), True
   
End If


End Function

