Attribute VB_Name = "calcDetecta"
Option Explicit
Public nomeIn, nomeOut As String
Public quantTubos, quantNos, quantBombas As Integer
Public linhas
Dim conteudo As String
Dim recWaterLines As Recordset
Dim recWaterComponents As Recordset
Dim recMoment As Recordset
Public Function obtemRede(ByVal Tc As TeCanvas) As Boolean

Dim linhasSelecionadas As String
Dim nosSelecionados As String
Dim tipo As Long, contador As Integer

'Em 19/10/2010
Dim ww As String
Dim ws As String
Dim wx As String
Dim wc As String
Dim wd As String
Dim wt As String
Dim at As String
Dim bt As String
Dim ct As String
ww = "WATERLINES"
ws = "OBJECT_ID_"
wx = "ID_TYPE"
wc = " WATERCOMPONENTS "
wt = " PIPENUMBER "
at = " CALCULANODE"
bt = " pipenumber "
ct = " pipenumber "



 If frmCanvas.TipoConexao <> 4 Then

   obtemRede = False
   'Retorna tipo de rede a exportar
   Dim frm  As frmSelectnetWorkTypes
   Set frm = New frmSelectnetWorkTypes
   If Not frm.init(tipo) Then Exit Function
   Set frm = Nothing
   
   '########################
   'Captura todos os tubos selecionados
   For contador = 0 To Tc.getSelectCount(lines) - 1
     If contador = 0 Then
        linhasSelecionadas = "'" & Tc.getSelectObjectId(contador, lines) & "'"
     Else
        linhasSelecionadas = linhasSelecionadas & ",'" & Tc.getSelectObjectId(contador, lines) & "'"
     End If
   Next
   Set recWaterLines = Conn.execute("SELECT * from WaterLines where object_id_ in(" & linhasSelecionadas & ") and id_Type=" & tipo & " order by pipenumber")
   
   '########################
   'Captura todos os nós ligados aos tubos selecionados
   If Not recWaterLines.EOF Then
      For contador = 0 To recWaterLines.RecordCount - 1
         If contador = 0 Then
            nosSelecionados = "'" & recWaterLines.Fields(9).value & "'"
         Else
            nosSelecionados = nosSelecionados & ",'" & recWaterLines.Fields(9).value & "'"
         End If
         nosSelecionados = nosSelecionados & ",'" & recWaterLines.Fields(10).value & "'"
         recWaterLines.MoveNext
      Next
   End If
   
   'alterado em 19/10/2010
   Else
   obtemRede = False
   'Retorna tipo de rede a exportar
  
   Set frm = New frmSelectnetWorkTypes
   If Not frm.init(tipo) Then Exit Function
   Set frm = Nothing
   
   '########################
   'Captura todos os tubos selecionados
   For contador = 0 To Tc.getSelectCount(lines) - 1
     If contador = 0 Then
        linhasSelecionadas = "'" & Tc.getSelectObjectId(contador, lines) & "'"
     Else
        linhasSelecionadas = linhasSelecionadas & ",'" & Tc.getSelectObjectId(contador, lines) & "'"
     End If
   Next
   Set recWaterLines = Conn.execute("SELECT  from " + """" + ww + """" + " where " + """" + ws + """" + " in(" & linhasSelecionadas & ") and " + """" + wx + """" + "=' " & tipo & "' order by " + """" + wt + """" + "")
   
   '########################
   'Captura todos os nós ligados aos tubos selecionados
   If Not recWaterLines.EOF Then
      For contador = 0 To recWaterLines.RecordCount - 1
         If contador = 0 Then
            nosSelecionados = "'" & recWaterLines.Fields(9).value & "'"
         Else
            nosSelecionados = nosSelecionados & ",'" & recWaterLines.Fields(9).value & "'"
         End If
         nosSelecionados = nosSelecionados & ",'" & recWaterLines.Fields(10).value & "'"
         recWaterLines.MoveNext
      Next
   End If
   
    If frmCanvas.TipoConexao <> 4 Then
   Set recWaterComponents = Conn.execute("SELECT * from WATERCOMPONENTS where OBJECT_ID_ in(" & nosSelecionados & ")" & " order by  CALCULANODE  ")
   Else
    ' Set recWaterComponents = Conn.execute("SELECT * from Watercomponents where object_id_ in(" & nosSelecionados & ")" & " order by calculenode")
   'Obtém a quantidade nós selecionados considerando considerando os nós de cálculo
   Set recMoment = Conn.execute("SELECT * from " + """" + wc + """" + " where " + """" + ws + """" + " in(" & nosSelecionados & ")" + """" + " order by " + """" + at + """" + "")
   
   End If
   
   If frmCanvas.TipoConexao <> 4 Then
   Set recMoment = Conn.execute("SELECT * from Watercomponents where object_id_ in(" & nosSelecionados & ") and calculeNode <>0")
   Else
   Set recMoment = Conn.execute("SELECT * from " + """" + wc + """" + " where " + """" + ws + """" + " in(" & nosSelecionados & ") and " + """" + at + """" + " <>'0'")
   End If
   
   
   quantNos = recMoment.RecordCount
   End If
    If frmCanvas.TipoConexao = 4 Then
   Set recMoment = Conn.execute("SELECT pipenumber from WaterLines where object_id_ in(" & linhasSelecionadas & ") and id_Type=" & tipo & " group by pipenumber")
  Else
   Set recMoment = Conn.execute("SELECT  from " + """" + wc + """" + " where " + """" + ws + """" + " in(" & linhasSelecionadas & ") and " + """" + wx + """" + "='" & tipo & "' group by " + """" + wt + """" + "")
   End If
   quantTubos = recMoment.RecordCount
   
   Tc.setCurrentLayer "WATERLINES"
   
   If quantTubos = 0 Or quantNos = 0 Then
      MsgBox "Trecho não encontrado" '
      obtemRede = False
      Exit Function
   End If
   
   If Not (recWaterLines Is Nothing) Then
      recWaterLines.MoveFirst
   End If
   
   If Not (recWaterComponents Is Nothing) Then
      recWaterComponents.MoveFirst
   End If
      obtemRede = True
End Function

Public Sub openForm()
    Dim i As Integer
    Dim nomeArq As String
    recWaterComponents.Filter = "id_type=19"
    frmEntradasDetecta.txtReservatorio.Text = recWaterComponents.RecordCount
    recWaterComponents.Filter = ""
    recWaterComponents.Filter = "demand<>0"
    frmEntradasDetecta.txtNosVazao.Text = recWaterComponents.RecordCount
    recWaterComponents.Filter = ""
    recWaterComponents.Filter = "id_type=20"
    frmEntradasDetecta.txtBombas.Text = recWaterComponents.RecordCount
    recWaterComponents.Filter = ""
    recWaterComponents.Filter = "id_type=1"
    frmEntradasDetecta.txtValvulas.Text = recWaterComponents.RecordCount
    recWaterComponents.Filter = ""
    frmEntradasDetecta.txtNos.Text = quantNos
    frmEntradasDetecta.txtTubos.Text = quantTubos
    frmEntradasDetecta.txtViscosidade.Text = "0.000001"
    frmEntradasDetecta.txtMassaEspeficac.Text = "1000"
    frmEntradasDetecta.txtVazaoInicial = "0"
    recWaterLines.MoveFirst
    While Not recWaterLines.EOF
      frmEntradasDetecta.listLeft.AddItem (recWaterLines.Fields("PipeNumber").value)
      recWaterLines.MoveNext
    Wend
    
    recWaterComponents.Filter = "demand<>0"
    i = 1
    While Not recWaterComponents.EOF
      frmEntradasDetecta.fgVazoesNos.AddItem ("")
      frmEntradasDetecta.fgVazoesNos.TextMatrix(i, 0) = recWaterComponents.Fields("CalculeNode").value
      i = i + 1
      recWaterComponents.MoveNext
    Wend
    
    frmEntradasDetecta.Show vbModal
    If frmEntradasDetecta.ok = True Then
      'nomeArq = frmEntradasDetecta.txtFile.Text
      'nomeArq = "c:\temp\detecta.dat"
      nomeArq = App.path & "\detecta.dat"
      If writeIn(nomeArq) Then
         If execute(nomeArq) Then
            readOut (mid(nomeArq, 1, Len(nomeArq) - 4) & ".out")
            'readOut (Mid(nomeArq, 1, Len(nomeArq) - 4) & "ext.out")
         End If
      End If
    End If
End Sub
  
Public Function writeIn(nomeArq As String) As Boolean
    writeIn = False
    Dim arqrede, noInicial, noFinal As String
    Dim comprimento As Double
    Dim linha As String
    arqrede = FreeFile
    
    Open nomeArq For Output As arqrede
    
    Print #arqrede, "Entre com o numero de tubos,de nos,de vazoes de entrada/saida e de reservato-"
    Print #arqrede, "rios respectivamente e salte uma linha:"
    Print #arqrede, frmEntradasDetecta.txtTubos.Text & " " & frmEntradasDetecta.txtNos.Text & " " & frmEntradasDetecta.txtNosVazao.Text & " " & frmEntradasDetecta.txtReservatorio.Text
    Print #arqrede, ""
    
    Print #arqrede, "Entre com o numero do trecho (entre primeiro com os numeros dos trechos fictici-"
    Print #arqrede, "os, que contem os reservatorios), o no inicial, o no final(os trechos ficticios"
    Print #arqrede, "nao tem no final),o diametro(m),o comprimento(m), a vazao inicial(adotada)(m3/s)"
    Print #arqrede, "e a rugosidade(m) respectivamente para cada trecho e salte uma linha:"
    
    recMoment.MoveFirst
    noInicial = -1
    noFinal = -1
    While Not recMoment.EOF
      recWaterLines.Filter = "pipenumber=" & recMoment.Fields("pipenumber").value
      While Not recWaterLines.EOF
         recWaterComponents.Filter = "object_id_ =" & recWaterLines.Fields(9).value
         If noInicial = -1 Then
            noInicial = recWaterComponents.Fields(14).value
         End If
         'noInicial = recWaterComponents.Fields(14).Value
         recWaterComponents.Filter = "object_id_ =" & recWaterLines.Fields(10).value
         noFinal = recWaterComponents.Fields(14).value
         If noFinal = "0" Then
            noFinal = ""
         End If
         comprimento = comprimento + recWaterLines.Fields(13).value
         linha = recWaterLines.Fields(21).value & " " & noInicial & " " & noFinal & " " & (recWaterLines.Fields(7).value / 1000) & " " & comprimento & " " & frmEntradasDetecta.txtVazaoInicial.Text & " " & recWaterLines.Fields(19).value
         recWaterComponents.Filter = ""
         recWaterLines.MoveNext
      Wend
      Print #arqrede, Replace(linha, ",", ".")
      noInicial = -1
      noFinal = -1
      comprimento = 0
      recMoment.MoveNext
    Wend
    recWaterComponents.MoveFirst
    recWaterLines.Filter = ""
    recWaterLines.MoveFirst
    Print #arqrede, ""
    
    Print #arqrede, "Entre com a viscosidade cinematica(m2/s) e a massa especifica(kg/m3) do fluido"
    Print #arqrede, "ra cada no para o regime permanente e salte uma linha:"
    Print #arqrede, frmEntradasDetecta.txtViscosidade.Text & " " & frmEntradasDetecta.txtMassaEspeficac.Text
    Print #arqrede, ""
    
    Print #arqrede, "Entre com o valor da pressao maxima(m) e da pressao minima admissiveis na re-"
    Print #arqrede, "de e salte uma linha:"
    Print #arqrede, frmEntradasDetecta.txtPressaoMax.Text & " " & frmEntradasDetecta.txtPressaoMin.Text
    Print #arqrede, ""
    
    Print #arqrede, "Entre com o numero do no e a vazao de entrada/saida(m3/s) respectivamente pa-"
    Print #arqrede, "ra cada no para o regime permanente e salte uma linha:"
    recWaterComponents.Filter = "demand<>0"
    While Not recWaterComponents.EOF
      linha = recWaterComponents.Fields(14).value & " " & recWaterComponents.Fields(13)
      Print #arqrede, Replace(linha, ",", ".")
      recWaterComponents.MoveNext
    Wend
    recWaterComponents.Filter = ""
    recWaterComponents.MoveFirst
    Print #arqrede, ""
    
    Print #arqrede, "Comecando pelo reservatorio pulmao,entre com o numero do reservatorio,o no do"
    Print #arqrede, "reservatorio, o nivel(m),a area(m2),o nivel maximo e o nivel minimo admissi-"
    Print #arqrede, "veis respectivamente para cada reservatorio e salte uma linha:"
    
    recWaterComponents.Filter = "id_type=19"
    Dim numRes As Integer, area As String, nivel As String, nivelMax As String, nivelMin As String
    numRes = 1
    While Not recWaterComponents.EOF
      retornaDadosReservatorio recWaterComponents.Fields("object_id_").value, area, nivel, nivelMax, nivelMin
      linha = numRes & " " & recWaterComponents.Fields(14).value & " " & nivel & " " & area & " " & nivelMax & " " & nivelMin
      Print #arqrede, Replace(linha, ",", ".")
      numRes = numRes + 1
      recWaterComponents.MoveNext
    Wend
    Print #arqrede, ""
    
    Print #arqrede, "Entre com o numero do no e a cota do no(m) respectivamente para cada no e sal-"
    Print #arqrede, "te uma linha:"
    recWaterComponents.Filter = "calculenode<>0"
    While Not recWaterComponents.EOF
      linha = recWaterComponents.Fields(14).value & " " & recWaterComponents.Fields(9)
      Print #arqrede, Replace(linha, ",", ".")
      recWaterComponents.MoveNext
    Wend
    recWaterComponents.Filter = ""
    recWaterComponents.MoveFirst
    Print #arqrede, ""
    
    Print #arqrede, "Deseja calcular o periodo extensivo (s/n)? (salte uma linha apos a resposta)(se"
    Print #arqrede, "o período extensivo não for ser calculado, digite 'n' e desconsidere todas as"
    Print #arqrede, "instruções seguintes)."
    If frmEntradasDetecta.optPermanente.value = True Then
        Print #arqrede, "n"
    Else
        Print #arqrede, "s"
    End If
    Print #arqrede, ""
    If frmEntradasDetecta.optPermanente.value = False Then
      Print #arqrede, "Entre com o numero de valvulas da rede(digite '0' se não houver válvulas na rede"
      Print #arqrede, "e desconsidere a próxima instrução)e salte uma linha:"
      'recWaterComponents.Filter = "id_Type=1"
      'Print #arqrede, recWaterComponents.RecordCount
      'recWaterComponents.Filter = ""
      'recWaterComponents.MoveFirst
      Print #arqrede, frmEntradasDetecta.txtValvulas.Text
      Print #arqrede, ""
    
      Print #arqrede, "Entre com o numero do tubo que contem a valvula e o coeficiente de perda de"
      Print #arqrede, "carga para a valvula fechada(kf) respectivamente e salte uma linha:"
      If frmEntradasDetecta.txtValvulas.Text = 0 Then
        'Print #arqrede, ""
      Else
       recWaterComponents.Filter = "id_type=1"
       Dim numTuboVal As Integer, perdaVal As String
         
       While Not recWaterComponents.EOF
           recWaterLines.Filter = "FinalComponent=" & recWaterComponents.Fields(1).value
           numTuboVal = recWaterLines.Fields("PipeNumber").value
           recWaterLines.Filter = ""
           retornaDadosValvula recWaterComponents.Fields("object_id_").value, perdaVal
           linha = numTuboVal & " " & perdaVal
           Print #arqrede, Replace(linha, ",", ".")
           recWaterComponents.MoveNext
       Wend
      End If
      Print #arqrede, ""
           
      Print #arqrede, "Entre com o numero de bombas da rede e com o numero de intervalos de tempo que"
      Print #arqrede, "as bombas trabalharao(1 a 24)(digite "; 0; 0; " se não houver bombas na rede e des-"
      Print #arqrede, "considere as duas próximas instruções) e salte uma linha:"
      If frmEntradasDetecta.txtBombas.Text = 0 Then
      
         Print #arqrede, "0 0"
         Print #arqrede, ""
         
         Print #arqrede, "Na mesma linha,entre com o numero dos intervalos de tempo nos quais as bombas"
         Print #arqrede, "trabalharao e salte uma linha:"
         'Print #arqrede, ""
         Print #arqrede, ""
         
         Print #arqrede, "Entre com o numero do tubo que contem a bomba,a carga de shut-off(m),a car-"
         Print #arqrede, "ga(m) e a vazao(m3/s) para o rendimento maximo da bomba e uma carga qual-"
         Print #arqrede, "quer(m) com a correspondente vazao(m3/s) respectivamente e salte uma linha:"
         'Print #arqrede, ""
         Print #arqrede, ""
         
      Else
      
         linha = frmEntradasDetecta.txtBombas.Text & " " & frmEntradasDetecta.ListIntervalos.SelCount
         Print #arqrede, Replace(linha, ",", ".")
         Print #arqrede, ""
         
         Print #arqrede, "Na mesma linha,entre com o numero dos intervalos de tempo nos quais as bombas"
         Print #arqrede, "trabalharao e salte uma linha:"
         Dim i As Integer
         linha = ""
         For i = 0 To frmEntradasDetecta.ListIntervalos.ListCount - 1
            If frmEntradasDetecta.ListIntervalos.Selected(i) = True Then
                If linha = "" Then
                    linha = i + 1
                Else
                    linha = linha & " " & i + 1
                End If
            End If
         Next i
         Print #arqrede, Replace(linha, ",", ".")
         Print #arqrede, ""
         
         Print #arqrede, "Entre com o numero do tubo que contem a bomba,a carga de shut-off(m),a car-"
         Print #arqrede, "ga(m) e a vazao(m3/s) para o rendimento maximo da bomba e uma carga qual-"
         Print #arqrede, "quer(m) com a correspondente vazao(m3/s) respectivamente e salte uma linha:"
         recWaterComponents.Filter = "id_type=20"
         Dim numTuboBom As Integer
         Dim shutoff As String, carga As String, vazao As String, cargaq As String, vazaoq As String
         While Not recWaterComponents.EOF
           recWaterLines.Filter = "FinalComponent=" & recWaterComponents.Fields(1).value
           numTuboBom = recWaterLines.Fields("PipeNumber").value
           recWaterLines.Filter = ""
           retornaDadosBomba recWaterComponents.Fields("object_id_").value, shutoff, carga, vazao, cargaq, vazaoq
           linha = numTuboBom & " " & shutoff & " " & carga & " " & vazao & " " & cargaq & " " & vazaoq
           Print #arqrede, Replace(linha, ",", ".")
           recWaterComponents.MoveNext
         Wend
         Print #arqrede, ""
         
      End If
      
      Print #arqrede, "Existe tubos com vazamento na rede?(salte uma linha apos a resposta)(se não hou-"
      Print #arqrede, "ver tubos com vazamentos na rede, digite 'n' e desconsidere as duas próximas"
      Print #arqrede, "instruções)."
      If frmEntradasDetecta.listLeft.SelCount = 0 Then
      
        Print #arqrede, "n"
        Print #arqrede, ""
        
        Print #arqrede, "Entre com o número de tubos que contêm vazamentos e salte uma linha:"
        'Print #arqrede, ""
        Print #arqrede, ""
    
        Print #arqrede, "Entre na mesma linha com os números dos tubos que contêm vazamentos e salte uma"
        Print #arqrede, "linha:"
        'Print #arqrede, ""
        Print #arqrede, ""
        
      Else
      
        Print #arqrede, "s"
        Print #arqrede, ""
        
        Print #arqrede, "Entre com o número de tubos que contêm vazamentos e salte uma linha:"
        linha = frmEntradasDetecta.listLeft.SelCount
        Print #arqrede, linha
        Print #arqrede, ""
    
        Print #arqrede, "Entre na mesma linha com os números dos tubos que contêm vazamentos e salte uma"
        Print #arqrede, "linha:"
        Dim j As Integer
        linha = ""
        For j = 0 To frmEntradasDetecta.listLeft.ListCount - 1
            If frmEntradasDetecta.listLeft.Selected(j) = True Then
                If linha = "" Then
                    linha = j + 1
                Else
                    linha = linha & " " & j + 1
                End If
            End If
        Next j
        Print #arqrede, Replace(linha, ",", ".")
        Print #arqrede, ""
        
      End If
      
      Print #arqrede, "Entre com o numero de intervalos de tempo (1 a 24) para se calcular o periodo"
      Print #arqrede, "extensivo e salte uma linha:"
      linha = frmEntradasDetecta.ListIntervalos.ListCount
      Print #arqrede, linha
      Print #arqrede, ""
    
      Print #arqrede, "Na primeira linha,entre com o numero dos nos de entrada e saida. A partir"
      Print #arqrede, "da segunda linha,entre com as vazoes de entrada e saida(m3/s) para cada perio-"
      Print #arqrede, "do e salte uma linha:"
      linha = ""
      For i = 1 To frmEntradasDetecta.txtNos
        linha = linha & i & " "
      Next i
      Print #arqrede, linha
      For i = 1 To frmEntradasDetecta.cbIntervalosCalc.Text
      linha = ""
        For j = 1 To frmEntradasDetecta.txtNos
            linha = linha & frmEntradasDetecta.fgVazoesNos.TextMatrix(j, i) & " "
        Next j
        Print #arqrede, Replace(linha, ",", ".")
      Next i
      Print #arqrede, ""
    
      Print #arqrede, "Escreva o nome do arquivo de saida para o periodo extensivo e salte uma linha:"
      'linha = frmEntradasDetecta.txtFile.Text
      linha = "detectaext.out"
      Print #arqrede, linha
      Print #arqrede, ""
    End If
    Print #arqrede, "##############################################################################"
    Print #arqrede, "#                                                                            #"
    Print #arqrede, "#        Este arquivo foi escrito pelo módulo calcDetecta do Geosan          #"
    Print #arqrede, "#        desenvolvido pela Nexus Geoengenharia e Comercio Ltda.              #"
    Print #arqrede, "#        com o apoio da FAPESP.                                              #"
    Print #arqrede, "#        O arquivo constitui os dados de entrada para calculo hidráulico     #"
    Print #arqrede, "#        de rede realizado pela aplicação Detecta desenvolvida pelo          #"
    Print #arqrede, "#        Engenheiro Victor Diniz.                                            #"
    Print #arqrede, "#                                                         Rodrigo Viviani    #"
    Print #arqrede, "#                                                      Analista de Sistemas  #"
    Print #arqrede, "#                                                                            #"
    Print #arqrede, "##############################################################################"
    
    Close #arqrede
    writeIn = True
    Unload frmEntradasDetecta
End Function

Public Function execute(nomeArq As String) As Boolean
'Dim a As Long
'a = ShellExecute(0, "open", App.path & "\bin\detecta.exe", "", "", 5)
'On Error GoTo execute:
ChDir (App.path)
   Shell App.path & "\detecta.exe", vbNormalFocus
   
   execute = True
   Exit Function
'execute:
'  execute = False
End Function

Public Function readOut(nomeArq As String) As Boolean
    Dim arqrede, linha As String
    
    arqrede = FreeFile

    Open nomeArq For Input As arqrede
    conteudo = Input(LOF(arqrede), arqrede)
    Close #arqrede
    
    linhas = Split(conteudo, vbCrLf)
    showResults
   'linha = linhas(UBound(linhas) - 1)

End Function

Public Sub showResults()
      Dim i, linhaParada, j, qualColuna, qualLinha As Integer
      Dim palavras
      linhaParada = 3
      quantTubos = 41
      quantNos = 34
      
      With frmResultadosDetecta
      
      .fg1.TextMatrix(0, 0) = "Trecho"
      .fg1.TextMatrix(0, 1) = "Nó Inicial"
      .fg1.TextMatrix(0, 2) = "Nó Final"
      .fg1.TextMatrix(0, 3) = "Carga Piezométrica do Nó Incial (m)"
      .fg1.TextMatrix(0, 4) = "Carga Piezométrica do Nó Final (m)"
      .fg1.TextMatrix(0, 5) = "Vazão (m3/s)"
      .fg1.TextMatrix(0, 6) = "Perda de Carga (m)"
      .fg1.TextMatrix(0, 7) = "Diâmetro"
      .fg1.TextMatrix(0, 8) = "Fator de Atrito"
      
      .fg2.TextMatrix(0, 0) = "Trecho"
      .fg2.TextMatrix(0, 1) = "Nó Inicial"
      .fg2.TextMatrix(0, 2) = "Nó Final"
      .fg2.TextMatrix(0, 3) = "Pressão Nó Inicial (m)"
      .fg2.TextMatrix(0, 4) = "Pressão Nó Final (m)"
      .fg2.TextMatrix(0, 5) = "Cota Nó Inicial (m)"
      .fg2.TextMatrix(0, 6) = "Cota Nó Final (m)"
  
      .fg3.TextMatrix(0, 0) = "Nó"
      .fg3.TextMatrix(0, 1) = "Vazão de Entrada e Saída (m3/s)"
      .fg3.TextMatrix(0, 2) = "Pressão (m)"
      
      qualLinha = 1
      For j = linhaParada To (linhaParada + quantTubos)
         .fg1.AddItem ("")
         palavras = Split(linhas(j), " ")
         qualColuna = 0
         If j = linhaParada Then
             For i = 0 To UBound(palavras)
               If palavras(i) <> "" Then
                  .fg1.TextMatrix(qualLinha, qualColuna) = palavras(i)
                  If qualColuna = 1 Then
                     qualColuna = qualColuna + 2
                  Else
                     qualColuna = qualColuna + 1
                  End If
               End If
            Next i
         Else
            For i = 0 To UBound(palavras)
               If palavras(i) <> "" Then
                  .fg1.TextMatrix(qualLinha, qualColuna) = palavras(i)
                  qualColuna = qualColuna + 1
               End If
            Next i
         End If
         
         qualLinha = qualLinha + 1
      Next j
      
      linhaParada = j + 3
      
      qualLinha = 1
      For j = linhaParada To (linhaParada + quantTubos)
         .fg2.AddItem ("")
         palavras = Split(linhas(j), " ")
         qualColuna = 0
         
         If j = linhaParada Then
             For i = 0 To UBound(palavras)
               If palavras(i) <> "" Then
                  .fg2.TextMatrix(qualLinha, qualColuna) = palavras(i)
                  If qualColuna = 1 Or qualColuna = 3 Then
                     qualColuna = qualColuna + 2
                  Else
                     qualColuna = qualColuna + 1
                  End If
               End If
            Next i
         Else
            For i = 0 To UBound(palavras)
               If palavras(i) <> "" Then
                  .fg2.TextMatrix(qualLinha, qualColuna) = palavras(i)
                  qualColuna = qualColuna + 1
               End If
            Next i
         End If
         
         qualLinha = qualLinha + 1
      Next j

      linhaParada = j + 3
      
      qualLinha = 1
      For j = linhaParada To (linhaParada + quantNos)
         .fg3.AddItem ("")
         palavras = Split(linhas(j), " ")
         qualColuna = 0
         For i = 0 To UBound(palavras)
         If palavras(i) <> "" Then
            .fg3.TextMatrix(qualLinha, qualColuna) = palavras(i)
            qualColuna = qualColuna + 1
         End If
         Next i
         qualLinha = qualLinha + 1
      Next j

      linhaParada = j + 2
      
      qualLinha = 1
      For j = linhaParada To (linhaParada + quantNos)
         palavras = Split(linhas(j), " ")
         qualColuna = 1
         For i = 0 To UBound(palavras)
         If palavras(i) <> "" Then
            If qualColuna = 2 Then
               .fg3.TextMatrix(qualLinha, qualColuna) = palavras(i)
            End If
            qualColuna = qualColuna + 1
         End If
         Next i
         qualLinha = qualLinha + 1
      Next j

      linhaParada = j
      
      .txtIteracoes.Text = linhas(j)
   
      .Show vbModal
      End With
End Sub

Public Function retornaDadosReservatorio(ByVal object_id As String, ByRef area As String, _
                                       nivel As String, nivelMax As String, nivelMin As String) As Boolean
                                       
   Dim fa As String
   Dim fb As String
   Dim fc As String
   Dim fd As String
   Dim fe As String
   Dim ff As String
   Dim fg As String
   
    fa = "DESCRIPTION_"
   fb = "VALUE_"
   fc = "WATERCOMPONENTSDATA"
   fd = "WATERCOMPONENTSSUBTYPES"
   fe = "ID_TYPE"
   ff = "ID_SUBTYPE"
   fg = "OBJECT_ID"
   
   Dim RsPilha As New ADODB.Recordset
    If frmCanvas.TipoConexao <> 4 Then
   RsPilha.CursorType = adOpenDynamic
   RsPilha.Open "SELECT description_,value_ from watercomponentsdata d " & _
                              "inner join watercomponentssubtypes s " & _
                              "on s.id_type=d.id_type and s.id_subtype=d.id_subtype " & _
                              "Where Object_id_ ='" & object_id & "'", Conn
   While Not RsPilha.EOF
      Select Case RsPilha.Fields("description_").value
         Case "Nivel Maximo"
            nivelMax = RsPilha.Fields(1).value
         Case "Nivel Minimo"
            nivelMin = RsPilha.Fields(1).value
         Case "Nível d Água"
            nivel = RsPilha.Fields(1).value
         Case "Área"
            area = RsPilha.Fields(1).value
      End Select
      RsPilha.MoveNext
   Wend
   RsPilha.Close
   Set RsPilha = Nothing
   
   Else
   'alterado em 19/10/2010
     RsPilha.CursorType = adOpenDynamic
   RsPilha.Open "SELECT " + """" + fa + """" + "," + """" + fb + """" + " from " + """" + fc + """" + " " & _
                              "inner join " + """" + fd + """" + " " & _
                              "on " + """" + fd + """" + "." + """" + fe + """" + "=" + """" + fc + """" + "." + """" + fe + """" + " and " + """" + fd + """" + "." + """" + ff + """" + "=" + """" + fc + """" + "." + """" + ff + """" + " " & _
                              "Where " + """" + fg + " ='" & object_id & "'", Conn
   While Not RsPilha.EOF
      Select Case RsPilha.Fields("description_").value
         Case "Nivel Maximo"
            nivelMax = RsPilha.Fields(1).value
         Case "Nivel Minimo"
            nivelMin = RsPilha.Fields(1).value
         Case "Nível d Água"
            nivel = RsPilha.Fields(1).value
         Case "Área"
            area = RsPilha.Fields(1).value
      End Select
      RsPilha.MoveNext
   Wend
   RsPilha.Close
   Set RsPilha = Nothing
   End If
   
End Function

Public Function retornaDadosValvula(ByVal object_id As String, ByRef coeficiente As String) As Boolean
   Dim RsPilha As New ADODB.Recordset
    Dim fa As String
   Dim fb As String
   Dim fc As String
   Dim fd As String
   Dim fe As String
   Dim ff As String
   Dim fg As String
   
    fa = "DESCRIPTION_"
   fb = "VALUE_"
   fc = "WATERCOMPONENTSDATA"
   fd = "WATERCOMPONENTSSUBTYPES"
   fe = "ID_TYPE"
   ff = "ID_SUBTYPE"
   fg = "OBJECT_ID"
   If frmCanvas.TipoConexao <> 4 Then
   
   RsPilha.CursorType = adOpenDynamic
   RsPilha.Open "SELECT description_,value_ from watercomponentsdata d " & _
                              "inner join watercomponentssubtypes s " & _
                              "on s.id_type=d.id_type and s.id_subtype=d.id_subtype " & _
                              "Where Object_id_ ='" & object_id & "'", Conn
   While Not RsPilha.EOF
      If RsPilha.Fields("description_").value = "k p/ Válvula Fechada" Then
         coeficiente = RsPilha.Fields(1).value
      End If
      RsPilha.MoveNext
   Wend
   RsPilha.Close
   Set RsPilha = Nothing
   Else
   
   'alterado em 19/10/2010
     RsPilha.CursorType = adOpenDynamic
   RsPilha.Open "SELECT " + """" + fa + """" + "," + """" + fb + """" + " from " + """" + fc + """" + " " & _
                              "inner join " + """" + fd + """" + " " & _
                              "on " + """" + fd + """" + "." + """" + fe + """" + "=" + """" + fc + """" + "." + """" + fe + """" + " and " + """" + fd + """" + "." + """" + ff + """" + "=" + """" + fc + """" + "." + """" + ff + """" + " " & _
                              "Where " + """" + fg + """" + " ='" & object_id & "'", Conn
                              
                              End If
   
End Function

Public Function retornaDadosBomba(ByVal object_id As String, shutoff As String, carga As String, vazao As String, cargaq As String, vazaoq As String) As Boolean
   Dim RsPilha As New ADODB.Recordset
   Dim fa As String
   Dim fb As String
   Dim fc As String
   Dim fd As String
   Dim fe As String
   Dim ff As String
   Dim fg As String
   
    fa = "DESCRIPTION_"
   fb = "VALUE_"
   fc = "WATERCOMPONENTSDATA"
   fd = "WATERCOMPONENTSSUBTYPES"
   fe = "ID_TYPE"
   ff = "ID_SUBTYPE"
   fg = "OBJECT_ID"
   If frmCanvas.TipoConexao <> 4 Then
   
   RsPilha.CursorType = adOpenDynamic
   RsPilha.Open "SELECT description_,value_ from watercomponentsdata d " & _
                              "inner join watercomponentssubtypes s " & _
                              "on s.id_type=d.id_type and s.id_subtype=d.id_subtype " & _
                              "Where Object_id_ ='" & object_id & "'", Conn
   While Not RsPilha.EOF
      Select Case RsPilha.Fields("description_").value
         Case "Carga Shut-Off"
            shutoff = RsPilha.Fields(1).value
         Case "Carga p/ Rendimento Máximo"
            carga = RsPilha.Fields(1).value
         Case "Vazão p/ Rendimento Máximo"
            vazao = RsPilha.Fields(1).value
         Case "Carga Qualquer "
            cargaq = RsPilha.Fields(1).value
         Case "Vazão Qualquer"
            vazaoq = RsPilha.Fields(1).value
      End Select
      RsPilha.MoveNext
   Wend
   RsPilha.Close
   Set RsPilha = Nothing
   Else
    
   
   'alterado em 19/10/2010
     RsPilha.CursorType = adOpenDynamic
   RsPilha.Open "SELECT " + """" + fa + """" + "," + """" + fb + """" + " from " + """" + fc + """" + " " & _
                              "inner join " + """" + fd + """" + " " & _
                              "on " + """" + fd + """" + "." + """" + fe + """" + "=" + """" + fc + """" + "." + """" + fe + """" + " and " + """" + fd + """" + "." + """" + ff + """" + "=" + """" + fc + """" + "." + """" + ff + """" + " " & _
                              "Where " + """" + fg + """" + " ='" & object_id & "'", Conn
                              
                              End If
   
End Function
