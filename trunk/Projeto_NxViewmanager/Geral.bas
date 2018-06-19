Attribute VB_Name = "Geral"
Option Explicit
'Pontos
'0 - mais
'1 - asterico
'2 - circulo sólido
'3 - xis
'4 - quadrado sólido
'5 - diamante sólido
'6 - circulo transparente
'7 - quadrado transparente
'8 - diamante transparente
'
'Linhas
'0 - Linha Sólida
'1 - Linha Tracejada
'2 - Linha pontilhada
'3 - Linha ponto linha
'4 - Linha ponto, ponto, Linha
'5 - Linha nula
'
'100 - Linha orientada
'101 - Linha Delimitada
'102 - Linha com seta dupla
'103 - Linha com seta e barra
'104 - Linha com direçãoF
'105 - Linha com barra 45 graus
'106 - Linha com barra 135 graus
'
'Poligonos
'0 - transparente
'1 - sólido
'2 - Hachurado Horizontal
'3 - Hachurado Vertical
'4 - Hachurado Diagonal Crescente
'5 - Hachurado Diagonal Descendente
'6 - Hachurado Cruzado
'7 - Hachurado Cruzado Diagonal
Public conn As ADODB.Connection
Public TypeConn As Integer
Public intTema As Integer
Public strCmdFiltro As String
 Dim aa As String
   Dim bb As String
   Dim sql As String, lista() As String
   Dim sa As String
   Dim sb As String
   Dim sc As String
   Dim sd As String
   Dim se As String
   Dim sf As String
   Dim sg As String
   Dim sh As String
   Dim si As String
   Dim sj As String
   Dim sl As String
   Dim sm As String
   Dim sn As String
   Dim so As String
   Dim sp As String
   Dim sq As String
   Dim sr As String
   Dim ss As String
   Dim st As String
   Dim su As String
   Dim sv As String
   Dim sx As String
   Dim sz As String
   Dim sk As String
   Dim sw As String
   Dim swx As String
   Dim sss As String
   Dim ssv As String
   Dim ssz As String
    Dim ssr As String
    Dim ssa As String
    Dim sst As String
    Dim ssq As String
    Dim sse As String
    Dim ssj As String
    Dim ssd As String
'FUNÇÕES PARA LER E GRAVAR NO ARQUIVO .INI-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Function LoadCboPolygon(mCbo As ComboBox)
   With mCbo
      .Clear
      .AddItem "transparente"
      .ItemData(.NewIndex) = 0
      .AddItem "Sólido"
      .ItemData(.NewIndex) = 1
      .AddItem "Hachurado Horizontal"
      .ItemData(.NewIndex) = 2
      .AddItem "Hachurado Vertical"
      .ItemData(.NewIndex) = 3
      .AddItem "Hachurado Diagonal Crescente"
      .ItemData(.NewIndex) = 4
      .AddItem "Hachurado Diagonal Descendente"
      .ItemData(.NewIndex) = 5
      .AddItem "Hachurado Cruzado"
      .ItemData(.NewIndex) = 6
      .AddItem "Hachurado Hachurado Cruzado Diagonal"
      .ItemData(.NewIndex) = 7
   End With
End Function



Public Function LoadCboLine(mCbo As ComboBox)
   With mCbo
      .Clear
      .AddItem "Sólido"
      .ItemData(.NewIndex) = 1
      .AddItem "Traço"
      .ItemData(.NewIndex) = 2
      .AddItem "Ponto"
      .ItemData(.NewIndex) = 3
      .AddItem "Traço ponto"
      .ItemData(.NewIndex) = 4
      .AddItem "Traço ponto ponto"
      .ItemData(.NewIndex) = 5
      .AddItem "Linha com direção"
      .ItemData(.NewIndex) = 107
   End With
End Function

Public Function GetCboListIndex(ID As Long, mCbo As ComboBox, Rep As Integer) As Integer
   Dim a As Integer
   
   If Rep = 2 Then
      If ID < 100 Then ID = ID + 1
   ElseIf Rep = 1 Then
      If ID = 1 Then
         ID = 0
      ElseIf ID = 0 Then
         ID = 1
      End If
   ElseIf Rep = 4 Then
      
   End If
   For a = 0 To mCbo.ListCount - 1
       If mCbo.ItemData(a) = ID Then
         GetCboListIndex = a
         Exit Function
       End If
   Next
   GetCboListIndex = -1
End Function

Public Function LoadCboPoints(mCbo As ComboBox)
   With mCbo
      .Clear
      .AddItem "mais"
      .ItemData(.NewIndex) = 0
      .AddItem "asterisco"
      .ItemData(.NewIndex) = 1
      .AddItem "circulo sólido"
      .ItemData(.NewIndex) = 2
      .AddItem "xis"
      .ItemData(.NewIndex) = 3
      .AddItem "quadrado sólido"
      .ItemData(.NewIndex) = 4
      .AddItem "diamante sólido"
      .ItemData(.NewIndex) = 5
      .AddItem "circulo transparente"
      .ItemData(.NewIndex) = 6
      .AddItem "quadrado transparente"
      .ItemData(.NewIndex) = 7
      .AddItem "diamante transparente"
      .ItemData(.NewIndex) = 8
   End With
End Function

'Public Function GetThemeWhere(ByRef Sql As String, LayerName As String, Lv As ListView) As String
'
'   Dim Rs As ADODB.Recordset, a As Integer, B As Integer, mValue As String, mFieldName As String, _
'         itmx As ListItem, mType As Integer, Sinal As String
'
'   Sql = Mid(Sql, InStr(1, Sql, "Where"), Len(Sql) - InStr(1, Sql, "Where"))
'
'   While Not Trim(Sql) = ""
'      GetFilter Sql, mFieldName, Sinal, mValue, mType
'      If mType = 1 Then 'Proceso que busca os atributos principais
'         Set Rs = Conn.Execute("PMSDP '" & LayerName & "',0,0,'0'")
'         For a = 0 To Rs.Fields.Count - 1
'            If Rs.Fields(a).Name = mFieldName Then
'               With frmTheme
'                  For B = 0 To .cboAttribute.ListCount - 1
'                     If .cboAttribute.ItemData(B) = a Then
'                        Set itmx = Lv.ListItems.Add(, , .cboAttribute.List(B))
'                        .cboAttribute.ListIndex = B
'                        itmx.SubItems(1) = Sinal
'                        GetCboListIndex CLng(mValue), .cboAttributeValue, 0
'                        .cboAttributeValue.ListIndex = GetCboListIndex(CLng(mValue), .cboAttributeValue, 0)
'                        itmx.SubItems(2) = .cboAttributeValue.Text
'                     End If
'                  Next
'               End With
'            End If
'         Next
'         Rs.Close
'      Else  'Proceso que busca os atributos Adicionais
'         Set Rs = Conn.Execute("PmSSP '" & LayerName & "', " & frmTheme.cboAttributeValue.ItemData(frmTheme.cboAttributeValue.ListIndex) & ", 0")
'         For a = 0 To Rs.Fields.Count - 1
'            If UCase(Rs.Fields(a).Name) = UCase(mFieldName) Then
'               With frmTheme
'                  For B = 0 To .cboAttribAdc.ListCount - 1
'                     If .cboAttribAdc.ItemData(B) = a Then
'                        Set itmx = Lv.ListItems.Add(, , .cboAttribAdc.List(B))
'                        .cboAttribAdc.ListIndex = B
'                        itmx.SubItems(1) = Sinal
'                        GetCboListIndex CLng(mValue), .cboAttribAdcValue, 0
'                        .cboAttribAdcValue.ListIndex = GetCboListIndex(CLng(mValue), .cboAttribAdcValue, 0)
'                        itmx.SubItems(2) = .cboAttribAdcValue.Text
'                     End If
'                  Next
'               End With
'            End If
'         Next
'         Rs.Close
'      End If
'   Wend
'
'   Set Rs = Nothing
'End Function
'
'Private Function GetFilter(Sql As String, mFiedName As String, Sinal As String, mValue As String, mType As Integer) As String
'
'   'Define o tipo
'
'
'   mType = IIf(Mid(Sql, IIf((InStr(1, Sql, ".") - 1) < 0, 1, (InStr(1, Sql, ".") - 1)), 1) = "A", 1, 2)
'   'lIMPA
'   Sql = Mid(Sql, InStr(1, Sql, "["), Len(Sql))
'   'Define o Nome do Campo
'   mFiedName = Mid(Sql, 2, InStr(1, Sql, "]") - 2)
'   'Limpar
'   Sql = Mid(Sql, InStr(1, Sql, "]") + 1, Len(Sql))
'   'Define o Sinal
'   If IsNumeric(Mid(Sql, 2, 1)) Or Mid(Sql, 2, 1) = "'" Then
'      Sinal = Left(Sql, 1)
'      Sql = Mid(Sql, 2, Len(Sql) - 1)
'   Else
'      If IsNull(Mid(Sql, 2, 2)) Or Mid(Sql, 2, 2) = "'" Then
'         Sinal = Trim(Left(Sql, 2))
'         Sql = Mid(Sql, 3, Len(Sql) - 2)
'      Else
'         Sinal = Trim(Left(Sql, 3))
'         Sql = Mid(Sql, 4, Len(Sql) - 3)
'      End If
'   End If
'   'Definir o Valor
'   mValue = Trim(Mid(Sql, 1, InStr(1, Sql, "and") - 1))
'
'
'End Function

Public Function LoadFilter(LayerName_ As String, StrCrt As String, SqlgetPmsp As String) As Boolean
   'On Error GoTo LoadFilter_Err
   Exit Function
   Dim a As Integer, rs As ADODB.Recordset
   Dim Inicio As Integer, Final As Integer
   
   With frmTheme
   Set rs = Nothing
   
   If StrCrt <> "" Then
      Dim RsComp As ADODB.Recordset, StrCrtII As String
      'Set RsComp = conn.Execute("PMSDP '" & LayerName_ & "',0,0,'0'")
      Set RsComp = conn.Execute(SqlgetPmsp)
      StrCrt = Replace(StrCrt, "A.", "")
      StrCrt = Replace(StrCrt, "B.", "")
      'DEIXA SOMENTE O CRITERIO
      Inicio = InStr(1, StrCrt, "Where ") + 5
      Final = InStr(1, StrCrt, ")") - Inicio
      StrCrt = Mid(StrCrt, Inicio, Final)
      
      'RETORNA O SINAL DO PRIMENIRO CRITERIO
      If InStr(1, StrCrt, "=") > 0 Then
         .cboOperador = "Igual"
      ElseIf InStr(1, StrCrt, ">") > 0 Then
         .cboOperador = "Maior"
      ElseIf InStr(1, StrCrt, "<") > 0 Then
         .cboOperador = "Menor"
      ElseIf InStr(1, StrCrt, "<>") > 0 Then
         .cboOperador = "Diferente"
      End If
      If .cboOperador <> "" Then
         'RETORNA O CAMPO DO PROMEIRO CRITERIO
         Inicio = InStr(1, StrCrt, "[") + 1
         Final = InStr(1, StrCrt, "]")
         StrCrtII = Mid(StrCrt, Inicio, Final - Inicio)
         For a = 0 To RsComp.Fields.Count - 1
            If StrCrtII = RsComp.Fields(a).Name Then
               .cboColunas.ListIndex = .cboColunas.ItemData(a)
               StrCrt = Replace(StrCrt, StrCrtII, .cboColunas)
               Exit For
            End If
         Next
         'RETORNA O VALOR DO PRIMEIRO CAMPOS(CRITERIO)
         Dim str As String
         Inicio = InStr(1, StrCrt, "]") + 2
         Final = InStr(Inicio, StrCrt, "and") - Inicio
         If Final <= 0 Then Final = 20
         If .cboFiltro.ListCount <= 0 Then .cboFiltro = Trim(Mid(StrCrt, Inicio, Final))
         For a = 0 To .cboFiltro.ListCount - 1
            If .cboFiltro.ItemData(a) = CInt(Mid(StrCrt, Inicio, Final)) Then
               .cboFiltro.ListIndex = a
               Exit For
            End If
         Next
      End If
      Dim Posicao As Integer
      'RETORNA O SINAL DO SEGUNDO CRITERIO
      
      If InStr(1, StrCrt, " = ") > 0 Then
         .cboOperador2 = "Igual"
         Posicao = InStr(1, StrCrt, " = ")
      ElseIf InStr(1, StrCrt, " > ") > 0 Then
         .cboOperador2 = "Maior"
         Posicao = InStr(1, StrCrt, " > ")
      ElseIf InStr(1, StrCrt, " < ") > 0 Then
         .cboOperador2 = "Menor"
         Posicao = InStr(1, StrCrt, " < ")
      ElseIf InStr(1, StrCrt, " <> ") > 0 Then
         .cboOperador2 = "Diferente"
         Posicao = InStr(1, StrCrt, " <> ")
      End If
      If .cboOperador2 <> "" Then
         Inicio = InStrRev(StrCrt, "[", Posicao) + 1
         Final = InStrRev(StrCrt, "]", Posicao) - InStrRev(StrCrt, "[", Posicao) - 1
         StrCrtII = Trim(Mid(StrCrt, Inicio, Final))
         For a = 0 To RsComp.Fields.Count - 1
            If StrCrtII = RsComp.Fields(a).Name Then
               .cboColunas2.ListIndex = .cboColunas.ItemData(a)
               StrCrt = Replace(StrCrt, StrCrtII, .cboColunas2)
               Exit For
            End If
         Next
         Inicio = Posicao + 3
         Final = IIf(InStr(Posicao, StrCrt, " and") = 0, Len(StrCrt) - Posicao, InStr(Posicao, StrCrt, " and"))
         If Final <= 0 Then Final = Inicio + 20
         If Final > 0 Then
            'Final = Final - Inicio
            If .cboFiltro2.ListCount > 0 Then
               For a = 0 To .cboFiltro2.ListCount - 1
                  If .cboFiltro2.ItemData(a) = CInt(Mid(StrCrt, Inicio, Final)) Then
                     .cboFiltro2.ListIndex = a
                     Exit For
                  End If
               Next
            Else
              If .cboColunas2.ListIndex >= 0 Then .cboFiltro2.Text = Mid(StrCrt, Inicio + 3, Final)
            End If
         End If
      End If
      'RETORNA O SINAL DO SUB CRITERIO
      If InStr(1, StrCrt, "  =  ") > 0 Then
         .cboOperadorSub = "Igual"
      ElseIf InStr(1, StrCrt, "  >  ") > 0 Then
         .cboOperadorSub = "Maior"
      ElseIf InStr(1, StrCrt, "  <  ") > 0 Then
         .cboOperadorSub = "Menor"
      ElseIf InStr(1, StrCrt, "  <>  ") > 0 Then
         .cboOperadorSub = "Diferente"
      End If
      If .cboOperadorSub <> "" Then
         'RETORNA O CAMPO DO SUBCAMPOS (SUBCRITERIO)
         Inicio = InStr(1, StrCrt, "[Id_SubType]") + 14
         Final = InStr(1, StrCrt, "and [value_") - Inicio
         If Inicio > 14 Then
            For a = 0 To .cboColunasSub.ListCount - 1
               If .cboColunasSub.ItemData(a) = CInt(Mid(StrCrt, Inicio, Final)) Then
                  .cboColunasSub.ListIndex = a
                  Exit For
               End If
            Next
         End If
         'RETORNA O VALOR DO SUBCAMPOS (SUBCRITERIO)
         Inicio = InStr(1, StrCrt, "[value_]") + 12
         Final = 25
         If .cboFiltroSub.ListCount >= 0 Then .cboFiltroSub = Trim(Mid(StrCrt, Inicio, Final))
         If Inicio > 12 Then
            For a = 0 To .cboFiltroSub.ListCount - 1
               If .cboFiltroSub.ItemData(a) = CInt(Mid(StrCrt, Inicio, Final)) Then
                  .cboFiltroSub.ListIndex = a
                  Exit For
               End If
            Next
         End If
      End If
      
      'RETORNA O SINAL DO PRIMENIRO CRITERIO
      If InStr(1, StrCrt, "=") > 0 Then
         .cboOperador = "Igual"
      ElseIf InStr(1, StrCrt, ">") > 0 Then
         .cboOperador = "Maior"
      ElseIf InStr(1, StrCrt, "<") > 0 Then
         .cboOperador = "Menor"
      ElseIf InStr(1, StrCrt, "<>") > 0 Then
         .cboOperador = "Diferente"
      End If
      
      'RETORNA O SINAL DO SUB CRITERIO
      If InStr(1, StrCrt, "  =  ") > 0 Then
         .cboOperadorSub = "Igual"
      ElseIf InStr(1, StrCrt, "  >  ") > 0 Then
         .cboOperadorSub = "Maior"
      ElseIf InStr(1, StrCrt, "  <  ") > 0 Then
         .cboOperadorSub = "Menor"
      ElseIf InStr(1, StrCrt, "  <>  ") > 0 Then
         .cboOperadorSub = "Diferente"
      End If
      
      'RETORNA O SINAL DO SEGUNDO CRITERIO
      If InStr(1, StrCrt, " = ") > 0 Then
         .cboOperador2 = "Igual"
         Posicao = InStr(1, StrCrt, " = ")
      ElseIf InStr(1, StrCrt, " > ") > 0 Then
         .cboOperador2 = "Maior"
         Posicao = InStr(1, StrCrt, " > ")
      ElseIf InStr(1, StrCrt, " < ") > 0 Then
         .cboOperador2 = "Menor"
         Posicao = InStr(1, StrCrt, " < ")
      ElseIf InStr(1, StrCrt, " <> ") > 0 Then
         .cboOperador2 = "Diferente"
         Posicao = InStr(1, StrCrt, " <> ")
      End If
      
      'RENOMEIRA OS CAMPOS DOS CRITERIOS
      StrCrt = Replace(StrCrt, "Id_SubType", "SubTipo")
      If .cboColunasSub.ListIndex >= 0 Then StrCrt = Replace(StrCrt, "value_", .cboColunasSub)
      
      If RsComp.State = adStateOpen Then RsComp.Close
      Set RsComp = Nothing
      .chkFiltro.Value = 1
   Else
      
   End If
   End With
   Exit Function
LoadFilter_Err:
   Resume Next
End Function


Public Function getPmsdp(LayerName As String, TypeQuery As Integer, object_id As String, tipoProvedor As Integer, MYCONN As ADODB.Connection) As String
   
'MsgBox "Public Function getPmsdp(LayerName As String, TypeQuery As Integer, object_id As String, tipoProvedor As Integer, MYCONN As ADODB.Connection) As String"
   
   '--PMSDP - Properties Manager Select Default Properties
   '
   '/*Existem 3 Tipos de saidas que serão determinas pela entrada de um parametro para n layers
   '2 - Single Select with alias
   '0- Single Select
   '1 - Multiple Select with Alias
   '3 - To Insert
   '*/
   '
   '--#####################################################################
   '--# ATENÇAO: #
   '--# #
   '--# 1 - Ao Modificar um campos(INSERIR,MODIFICAR OU EXCLUIR) de um Layer, #
   '--# deverá faze-lo igualmente para todas as saídas do respectivo layer #
   '--# 2 - Ao Criar um novo Layer será necessário a inserção de todas as saídas #
   '--# #
   '--#####################################################################
   Dim sql As String, FieldName As String, AttributeTable As String, AttributeLink As String
   
   RetornaNomeAtr MYCONN, LayerName, AttributeTable, AttributeLink
   

      If TypeConn <> 4 Then

   

   Select Case UCase(LayerName) 'LAYER REDES
      Case UCase("waterlines"), UCase("sewerlines"), UCase("drainlines")
      sql = sql & " Select Line_Id as " + """" + "Linha" + """" + ","
      sql = sql & " Id_Type as Tipo,"
      
'      sql = sql & " InitialGroundHeight as [Inicial Cota Terreno],"
'      sql = sql & " FinalGroundHeight as [Final Cota Terreno],"
'      sql = sql & " InitialTubeDeepness as [Inicial Profundidade],"
'      sql = sql & " FinalTubeDeepness as [Final Profundidade],"
      
      
      sql = sql & " InitialGroundHeight as " + """" + "[TERRENO - COTA INICIAL]" + """" + ","
      sql = sql & " FinalGroundHeight as " + """" + "[TERRENO - COTA FINAL]" + """" + ","
      sql = sql & " InitialTubeDeepness as " + """" + "[PEÇA - COTA INICIAL]" + """" + ","
      sql = sql & " FinalTubeDeepness as " + """" + "[PEÇA - COTA FINAL]" + """" + ","
      
      
      sql = sql & " InternalDiameter as " + """" + "[Diametro Inter.(mm)]" + """" + ","
      sql = sql & " ExternalDiameter as " + """" + "[Diametro Ext.(mm)]" + """" + ","
      sql = sql & " InitialComponent as " + """" + "[Inicial Componente]" + """" + ","
      sql = sql & " FinalComponent as " + """" + "[Final Componente]" + """" + ","
      sql = sql & " Thickness as " + """" + "Densidade" + """" + ","
      sql = sql & " Material,"
      sql = sql & " Length as " + """" + "[Comprimento(m)]" + """" + ","
      sql = sql & " LengthCalculated as " + """" + "[Compr. Calculado]" + """" + ","
      sql = sql & " Supplier as " + """" + "Fornecedor" + """" + ","
      sql = sql & " Location as " + """" + "Localização" + """" + ","
      sql = sql & " State as " + """" + "Estado" + """" + ","
      If UCase(LayerName) = UCase("waterlines") Then
         sql = sql & " RoughNess as " + """" + "Rugosidade" + """" + ","
      End If
      
      sql = sql & "USUARIO_LOG as Usuário, "
      
      'sql = sql & "DATA_LOG as [Data Cadastro], "
      
      sql = sql & " Sector as " + """" + "Setor" + """" + ","
      sql = sql & " InformationValidity As " + """" + "Validade" + """" + ","
      sql = sql & " DateInstallation As " + """" + "[Data_de_Instalação]" + """" + ", SideStreet as " + """" + "[Lado_da_Rua]" + """" + ", DividedDistance as " + """" + "[Distância_da_Divisa]" + """" + ""
      sql = sql & " From " & LayerName
      
      'MsgBox "TIPO DA QUERY = " & TypeQuery
      'MsgBox "COMANDO SQL = " & sql
      
      
      Select Case TypeQuery
         Case 0
            'ESTE COMANDO É DADO PARA RETORNAR UM RECORDSET VAZIO PARA APROVEITAR SOMENTE OS NOMES DAS COLUNAS DO RECORDSET
            
            sql = "select Line_Id,id_Type,InitialGroundHeight,FinalGroundHeight,initialTubeDeepness,FinalTubeDeepness,InternalDiameter,ExternalDiameter,InitialComponent,FinalComponent,Thickness,Material,Length,LengthCalculated,Supplier,Location,State,RoughNess,USUARIO_LOG,Sector, InformationValidity, DateInstallation, SideStreet, DividedDistance from " & LayerName & " where object_id_ = '0'"
            'MsgBox "COMANDO SQL da Typequery 0 = " & sql
            
         Case 1 'Multiple Select with Alias
            sql = sql & " Where line_id in (" & object_id & ")"
         Case 2 'Single Select with alias
            sql = sql & " Where line_id in (" & object_id & ")"
         
         Case 3 'default
            
            'QUANDO É SELECIONADO 'DESENHAR REDES', ESTE SELECT ABAIXO CARREGA AS TAGS NO GERENCIADOR DE ATRIBUTOS
            'CARREGA O COMBO DE FILTROS COM ESTE SELECT ABAIXO
            
                     sql = "Select 0 as " + """" + "LINHA" + """" + ", "
                     sql = sql & " 0 as " + """" + "TIPO" + """" + ", "
                     sql = sql & "0 as " + """" + "[TERRENO - COTA INICIAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[TERRENO - COTA FINAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[PEÇA - COTA INICIAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[PEÇA - COTA FINAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[DIAMETRO INTERNO]" + """" + ", "
                     sql = sql & "0 as " + """" + "[DIAMETRO EXTERNO]" + """" + ", "
                     sql = sql & "0 as " + """" + "[INICIAL COMPONENTE]" + """" + ", "
                     sql = sql & "0 as " + """" + "[FINAL COMPONENTE]" + """" + ", "
                     sql = sql & "0 as " + """" + "DENSIDADE" + """" + ", "
                     sql = sql & "0 as " + """" + "MATERIAL" + """" + " , "
                     sql = sql & "0 as " + """" + "COMPRIMENTO" + """" + ", "
                     sql = sql & "0 as " + """" + "[COMPR. CALCULADO]" + """" + ", "
                     sql = sql & "0 as " + """" + "FORNECEDOR" + """" + ", "
                     sql = sql & "0 as " + """" + "FABRICANTE" + """" + ", "
                     sql = sql & "0 as " + """" + "LOCALIZAÇÃO" + """" + ", "
                     sql = sql & "0 as " + """" + "ESTADO" + """" + ", "
                     sql = sql & "0 as " + """" + "RUGOSIDADE" + """" + ", "
                     sql = sql & "0 as " + """" + "SETOR" + """" + ", '' as " + """" + "VALIDADE" + """" + ", "
                     sql = sql & "'' As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", "
                     sql = sql & "1 as " + """" + "[LADO_DA_RUA]" + """" + ", "
                     sql = sql & "0 as " + """" + "[DISTÂNCIA_DA_DIVISA]" + """" + ", "
                     sql = sql & "'' as " + """" + "USUÁRIO" + """" + ", "
                     sql = sql & "'' AS " + """" + "[DATA CADASTRO]" + """" + ""
                     sql = sql & "from x_state where stateid =1"
                     
      End Select
   
   
   Case UCase("watercomponents"), UCase("sewercomponents"), UCase("draincomponents")
      
'MsgBox "Entrou no UCase(watercomponents), UCase(sewercomponents), UCase(draincomponents)"
      
      sql = sql & " Select Component_id as " + """" + "Componente" + """" + ","
      sql = sql & " ID_Type as " + """" + "Tipo" + """" + ","
      sql = sql & " YearOfConstruction as " + """" + "[Ano de Fabricação]" + """" + ","
      sql = sql & " State as " + """" + "Estado" + """" + ","
      sql = sql & " Location as " + """" + "Localização" + """" + ","
      sql = sql & " Manufacturer as " + """" + "Fabricante" + """" + " ,"
      sql = sql & " GroundHeight as " + """" + "[Cota do Terreno]" + """" + ","
      
      'CÓDIGO ANTERIOR A 17/02/09
      'If UCase(LayerName) = UCase("sewercomponents") Or UCase(LayerName) = UCase("draincomponents") Then
      '  sql = sql & " GroundHeightFinal as [Cota do Fundo],"
      'End If
      
      'CÓDIGO POSTERIOR A 17/02/09
      If UCase(LayerName) = UCase("watercomponents") Or UCase(LayerName) = UCase("sewercomponents") Or UCase(LayerName) = UCase("draincomponents") Then
        sql = sql & " GroundHeightFinal as " + """" + "[Cota do Fundo]" + """" + ","
      End If
      
      If UCase(LayerName) = UCase("watercomponents") Then
        sql = sql & " Demand as " + """" + "Demanda" + """" + ","
        sql = sql & " CalculeNode as " + """" + "[Nó de Cálculo]" + """" + ","
      End If
      
      sql = sql & " InformationValidity as " + """" + "Validade" + """" + ","
      sql = sql & " Notes As " + """" + "Observação" + """" + ","
      sql = sql & " Trouble as " + """" + "[Não_Conformidade]" + """" + ", DateInstallation As " + """" + "[Data_de_Instalação]" + """" + ", sector as " + """" + "setor" + """" + ""
      sql = sql & " From " & LayerName


'MsgBox "TypeQuery 2: " & TypeQuery
      
      Select Case TypeQuery
      Case 0
         If UCase(LayerName) = "WATERCOMPONENTS" Then
            sql = "SELECT Component_id , id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight, GroundHeightFinal, Demand, calculenode, InformationValidity, Notes, Trouble, DateInstallation, sector from WATERCOMPONENTS where object_id_ = '0'"
         ElseIf UCase(LayerName) = "SEWERCOMPONENTS" Then
            sql = "SELECT Component_id , id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight, GroundHeightFinal, InformationValidity, Notes, Trouble, DateInstallation, SECTOR FROM SEWERCOMPONENTS WHERE OBJECT_ID_ = '0'"
         End If
         
         'sql = "select Component_id , id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight" & IIf(UCase(LayerName) = "WATERCOMPONENTS", ", GroundHeightFinal", ", Demand, calculenode ") & ", InformationValidity, Notes, Trouble, DateInstallation from " & LayerName & " where object_id_ = '0'"
         'sql = "select Component_id , id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight" & IIf(UCase(LayerName) = "WATERCOMPONENTS", ", GroundHeightFinal, Demand, calculenode", "") & IIf(UCase(LayerName) = "SEWERCOMPONENTS", ", GroundHeightFinal, Demand", "") & ", InformationValidity, Notes, Trouble, DateInstallation, sector from " & LayerName & " where object_id_ = '0'"
         'sql = "select Component_id , id_Type, YearOfConstruction, State, Location, Manufacturer, GroundHeight" & IIf(UCase(LayerName) = "WATERCOMPONENTS", ", GroundHeightFinal", ", Demand, calculenode ") & ", InformationValidity, Notes, Trouble, DateInstallation from " & LayerName & " where object_id_ = '0'"
'MsgBox sql
      Case 1 'Multiple Select with Alias
         sql = sql & " Where Component_id in (" & object_id & ")"
'MsgBox sql
      Case 2 'Single Select with alias
         sql = sql & " Where Component_id in (" & object_id & ")"
'MsgBox sql
      Case 3 'tupla default
      sql = "Select 0 as Componente, 0 as " + """" + "Tipo" + """" + ",0 as " + """" + "[Ano de Fabricação]" + """" + ", 0 as " + """" + "Estado" + """" + ",0 as " + """" + "Localização" + """" + ", 0 as " + """" + "Fabricante" + """" + " , 0 as " + """" + "[Cota do Terreno]" + """" + " " & IIf(LayerName <> "WATERCOMPONENTS", ", 0 as " + """" + "[Cota do Fundo]" + """" + "", ", 0 as " + """" + "Demanda" + """" + ", 0 as " + """" + "[Nó de Cálculo]" + """" + "") & ", 0 as " + """" + "Validade" + """" + ", '' as " + """" + "Observação" + """" + ", 0 as " + """" + "[Não_Conformidade]" + """" + ", '' As " + """" + "[Data_de_Instalação]" + """" + " from x_state where stateid =1"
'MsgBox sql
       '  sql = "Select 0 as "+""""+"Componente"+""""+", 0 as "+""""+"Tipo"+""""+", 0 as "+""""+"[Ano de Fabricação]"+""""+", 0 as "+""""+"Estado"+""""+", 0 as "+""""+"Localização"+""""+", 0 as "+""""+"Fabricante"+""""+" , 0 as "+""""+"[Cota do Terreno]"+""""+" & IIf(LayerName <> "WATERCOMPONENTS", ", 0 as "+""""+"[Cota do Fundo]"+""""+"," , 0 as "+""""+"Demanda"+""""+", 0 as "+""""+"[Nó de Cálculo]"+""""+") & ", 0 as "+""""+"Validade"+""""+", '' as "+""""+"Observação"+""""+", 0 as "+""""+"[Não_Conformidade]"+""""+", '' As "+""""+"[Data_de_Instalação]"+""""+" from x_state where stateid =1"
'MsgBox sql
      End Select
   Case Else 'qualque plano
      If Not (AttributeLink = "" Or AttributeTable = "") Then
         sql = "SELECT * FROM " & AttributeTable & " WHERE " & AttributeLink & " in(" & object_id & ")"
      End If
   End Select
   
   getPmsdp = convertQuery(sql, tipoProvedor)







'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"






Else

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


a = "LINE_ID"
b = "ID_TYPE"
c = "INITIALGROUNDHEIGHT"
d = "FINALGROUNDHEIGHT"
e = "INITIALTUBEDEEPNESS"
f = "FINALTUBEDEEPNESS"
g = "INTERNALDIAMETER"
h = "EXTERNALDIAMETER"
i = "INITIALCOMPONENT"
j = "FINALCOMPONENT"
k = "THICKNESS"
l = "MATERIAL"
   Dim a1 As String
Dim b1 As String
Dim c1 As String
Dim d1 As String
Dim e1 As String
Dim f1 As String
Dim g1 As String
Dim h1 As String
Dim i1 As String
Dim j1 As String
Dim k1 As String
Dim l1 As String


a1 = "LENGTH"
b1 = "LENGTHCALCULATED"
c1 = "SUPPLIER"
d1 = "LOCATION"
e1 = "STATE"
f1 = "ROUGHNESS"
g1 = "USUARIO_LOG"
h1 = "SECTOR"
i1 = "INFORMATIONVALIDITY"
j1 = "DATEINSTALLATION"
k1 = "SIDESTREET"
l1 = "DIVIDEDDISTANCE"

Dim m As String
m = object_id




   Select Case UCase(LayerName) 'LAYER REDES
      Case UCase("waterlines"), UCase("sewerlines"), UCase("drainlines")
      sql = sql & " Select " + """" + a + """" + " as " + """" + "Linha" + """" + ","
      sql = sql & "" + """" + b + """" + " as " + """" + "Tipo" + """" + ","
      
'      sql = sql & " InitialGroundHeight as [Inicial Cota Terreno],"
'      sql = sql & " FinalGroundHeight as [Final Cota Terreno],"
'      sql = sql & " InitialTubeDeepness as [Inicial Profundidade],"
'      sql = sql & " FinalTubeDeepness as [Final Profundidade],"
      
      
      sql = sql & "" + """" + c + """" + " as " + """" + "[TERRENO - COTA INICIAL]" + """" + ","
      sql = sql & "" + """" + d + """" + " as " + """" + "[TERRENO - COTA FINAL]" + """" + ","
      sql = sql & "" + """" + e + """" + " as " + """" + "[PEÇA - COTA INICIAL]" + """" + ","
      sql = sql & " " + """" + f + """" + " as " + """" + "[PEÇA - COTA FINAL]" + """" + ","
      
      
      sql = sql & " " + """" + g + """" + " as " + """" + "[Diametro Inter.(mm)]" + """" + ","
      sql = sql & " " + """" + h + """" + " as " + """" + "[Diametro Ext.(mm)]" + """" + ","
      sql = sql & " " + """" + i + """" + " as " + """" + "[Inicial Componente]" + """" + ","
      sql = sql & " " + """" + j + """" + " as " + """" + "[Final Componente]" + """" + ","
      sql = sql & " " + """" + k + """" + " as " + """" + "Densidade" + """" + ","
      sql = sql & " " + """" + l + """" + ","
      sql = sql & " " + """" + a1 + """" + "as " + """" + "[Comprimento(m)]" + """" + ","
      sql = sql & " " + """" + b1 + """" + " as " + """" + "[Compr. Calculado]" + """" + ","
      sql = sql & " " + """" + c1 + """" + " as " + """" + "Fornecedor" + """" + ","
      sql = sql & " " + """" + d1 + """" + " as " + """" + "Localização" + """" + ","
      sql = sql & " " + """" + e1 + """" + " as " + """" + "Estado" + """" + ","
      If UCase(LayerName) = UCase("waterlines") Then
         sql = sql & " " + """" + f1 + """" + " as " + """" + "Rugosidade" + """" + ","
      End If
      
      sql = sql & "" + """" + g1 + """" + " as " + """" + "Usuário" + """" + ", "
      
      'sql = sql & "DATA_LOG as [Data Cadastro], "
      
      sql = sql & " " + """" + h1 + """" + " as " + """" + "Setor" + """" + ","
      sql = sql & " " + """" + i1 + """" + " As " + """" + "Validade" + """" + ","
      sql = sql & " " + """" + j1 + """" + " As " + """" + "[Data_de_Instalação]" + """" + ", " + """" + k1 + """" + " as " + """" + "[Lado_da_Rua]" + """" + "," + """" + l1 + """" + "as " + """" + "[Distância_da_Divisa]" + """" + " "
      sql = sql & " From " + """" + LayerName + """" + ""
      
      'MsgBox "TIPO DA QUERY = " & TypeQuery
      'MsgBox "COMANDO SQL = " & sql
      
      
      
     ' a = "LINE_ID"
'b = object_id
'c = "'b'"
      
      Select Case TypeQuery
         Case 0
            'ESTE COMANDO É DADO PARA RETORNAR UM RECORDSET VAZIO PARA APROVEITAR SOMENTE OS NOMES DAS COLUNAS DO RECORDSET
            
            sql = "select " + """" + a + """" + "," + """" + b + """" + "," + """" + c + """" + "," _
            + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + "," + """" + g + """" + "," + """" + h + """" + "," + """" + i + """" + "," _
            + """" + j + """" + "," + """" + k + """" + "," + """" + l + """" + "," + """" + a1 + """" + "," + """" + b1 + """" + "," _
            + """" + c1 + """" + "," + """" + d1 + """" + "," + """" + e1 + """" + "," + """" + f1 + """" + "," + """" + g1 + """" + "," _
            + """" + h1 + """" + "," + """" + i1 + """" + "," + """" + j1 + """" + "," + """" + k1 + """" + "," + """" + l1 + """" + " from " + _
            """" + LayerName + """" + " where " + m + " = '0'"
            'MsgBox "COMANDO SQL da Typequery 0 = " & sql
            
         Case 1 'Multiple Select with Alias
            sql = sql & " Where " + """" + "LINE_ID" + """" + " in (" & object_id & ")"
         Case 2 'Single Select with alias
            sql = sql & " Where " + """" + "LINE_ID" + """" + " in (" & object_id & ")"
        
         Case 3 'default
            
            'QUANDO É SELECIONADO 'DESENHAR REDES', ESTE SELECT ABAIXO CARREGA AS TAGS NO GERENCIADOR DE ATRIBUTOS
            'CARREGA O COMBO DE FILTROS COM ESTE SELECT ABAIXO
            
                     sql = "Select 0 as " + """" + "LINHA" + """" + ", "
                     sql = sql & " 0 as " + """" + "TIPO" + """" + ", "
                     sql = sql & "0 as " + """" + "[TERRENO - COTA INICIAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[TERRENO - COTA FINAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[PEÇA - COTA INICIAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[PEÇA - COTA FINAL]" + """" + ", "
                     sql = sql & "0 as " + """" + "[DIAMETRO INTERNO]" + """" + ", "
                     sql = sql & "0 as " + """" + "[DIAMETRO EXTERNO]" + """" + ", "
                     sql = sql & "0 as " + """" + "[INICIAL COMPONENTE]" + """" + ", "
                     sql = sql & "0 as " + """" + "[FINAL COMPONENTE]" + """" + ", "
                     sql = sql & "0 as " + """" + "DENSIDADE" + """" + ", "
                     sql = sql & "0 as " + """" + "MATERIAL" + """" + " , "
                     sql = sql & "0 as " + """" + "COMPRIMENTO" + """" + ", "
                     sql = sql & "0 as " + """" + "[COMPR. CALCULADO]" + """" + ", "
                     sql = sql & "0 as " + """" + "FORNECEDOR" + """" + ", "
                     sql = sql & "0 as " + """" + "FABRICANTE" + """" + ", "
                     sql = sql & "0 as " + """" + "LOCALIZAÇÃO" + """" + ", "
                     sql = sql & "0 as " + """" + "ESTADO" + """" + ", "
                     sql = sql & "0 as " + """" + "RUGOSIDADE" + """" + ", "
                     sql = sql & "0 as " + """" + "SETOR" + """" + ", '' as " + """" + "VALIDADE" + """" + ", "
                     sql = sql & "'' As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", "
                     sql = sql & "1 as " + """" + "[LADO_DA_RUA]" + """" + ", "
                     sql = sql & "0 as " + """" + "[DISTÂNCIA_DA_DIVISA]" + """" + ", "
                     sql = sql & "'' as " + """" + "USUÁRIO" + """" + ", "
                     sql = sql & "'' AS " + """" + "[DATA CADASTRO]" + """" + " "
                     sql = sql & "from " + """" + "X_STATE" + """" + " where " + """" + "STATEID" + """" + " ='1'"
                     
      End Select
      
    getPmsdp = convertQuery(sql, tipoProvedor)
    'MsgBox sql
    End Select
    
   End If
   
   'MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
   
   
   
'      Close #1
'      Open App.Path & "\GeoSanLog.txt" For Append As #1
'      Print #1, Now & " getPmssp - " & sql
'      Close #1

'MsgBox "Query final 2:" & sql
   
End Function







Public Function getPmsdp2(LayerName As String, TypeQuery As Integer, object_id As String, tipoProvedor As Integer, MYCONN As ADODB.Connection) As String
   sql = ""
   'Case UCase("watercomponents"), UCase("sewercomponents"), UCase("draincomponents")
             If TypeConn <> 4 Then
            
        sql = sql & " Select Component_id as " + """" + "COMPONENTE" + """" + ","
        sql = sql & " ID_Type as " + """" + "TIPO" + """" + ","
        sql = sql & " YearOfConstruction as " + """" + "[ANO DE FABRICAÇÃO]" + """" + ","
        sql = sql & " State as " + """" + "ESTADO" + """" + ","
        sql = sql & " Location as " + """" + "LOCALIZAÇÃO" + """" + ","
        sql = sql & " Supplier as " + """" + "FORNECEDOR" + """" + ","
        sql = sql & " Manufacturer as " + """" + "FABRICANTE" + """" + " ,"
        sql = sql & " GroundHeight as " + """" + "[COTA DO TERRENO]" + """" + ","
        
        If UCase(LayerName) = "SEWERCOMPONENTS" Or UCase(LayerName) = "DRAINCOMPONENTS" Then
            sql = sql & " GroundHeightFinal as " + """" + "[COTA DO FUNDO]" + """" + ","
        End If
        
        If UCase(LayerName) = UCase("watercomponents") Then
            sql = sql & " Demand as " + """" + "DEMANDA" + """" + ","
            sql = sql & " CalculeNode as " + """" + "[NÓ DE CÁLCULO]" + """" + ","
        End If
        
        sql = sql & " InformationValidity as " + """" + "VALIDADE" + """" + ","
        sql = sql & " Notes As " + """" + "Observação" + """" + ","
        sql = sql & " Trouble as " + """" + "[NÃO_CONFORMIDADE]" + """" + ", DateInstallation As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", Pattern as " + """" + "[PADRÃO_CONSUMO]" + """" + ", Sector as " + """" + "[SETOR]" + """" + ""
        'sql = sql & " ANGLE,NOME_CELUL,ORIGEM_CAL,X_,Y_,COR,TAMANHO_X,TAMANHO_Y,CENT_CEL_X,CENT_CEL_Y,COR_CELULA,ESC_CEL_X,ESC_CEL_Y "
        sql = sql & " From " & LayerName
        
        Select Case TypeQuery
            Case 0
                sql = "select Component_id , id_Type, YearOfConstruction, State, Location, Supplier, Manufacturer, GroundHeight" & IIf(UCase(LayerName) <> UCase("watercomponents"), ", GroundHeightFinal", ", Demand, calculenode ") & ", InformationValidity, Notes,Trouble, DateInstallation, Pattern, Sector from " & LayerName & " where object_id_ = 1"
            Case 1 'Multiple Select with Alias
                sql = sql & " Where Component_id in (" & object_id & ")"
            Case 2 'Single Select with alias
            
                sql = sql & " Where Component_id in (" & object_id & ")"
            Case 3 'tupla default
                sql = "Select  0 as " + """" + "COMPONENTE" + """" + ", 0 as " + """" + "TIPO" + """" + ", 0 as " + """" + "[ANO DE FABRICAÇÃO]" + """" + ", 0 as " + """" + "ESTADO" + """" + ", 0 as " + """" + "LOCALIZAÇÃO" + """" + ", 0 as " + """" + "FORNECEDOR" + """" + ", 0 as " + """" + "FABRICANTE" + """" + " , 0 as " + """" + "[COTA DO TERRENO]" + """" + " " & IIf(UCase(LayerName) <> UCase("watercomponents"), ", 0 as " + """" + "[COTA DO FUNDO]" + """" + "", ", 0 as " + """" + "DEMANDA" + """" + ", 0 as " + """" + "[NÓ DE CÁLCULO]" + """" + "") & ", 0 as " + """" + "VALIDADE" + """" + ", '' as " + """" + "Observação" + """" + ", 0 as " + """" + "[NÃO_CONFORMIDADE]" + """" + ", '' As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", '' as " + """" + "[PADRÃO_CONSUMO]" + """" + ", '' as " + """" + "[SETOR]" + """" + " from x_state where stateid  =1"
     
        
         Case Else 'qualque plano
        sql = "SELECT * FROM " & LayerName & " WHERE OBJECT_ID_ in(" & object_id & ")"
   End Select


    Else 'alterado em 21/10/2010
    
    Dim ja As String
    Dim je As String
     Dim jo As String
     Dim a As String
    ja = "SEWERCOMPONENTS"
    je = "DRAINCOMPONENTS"
    ssq = "WATERCOMPONENTS"
    jo = "OBJECT_ID_"
    swx = "ID_TYPE"
    sss = "YEAROFCONSTRUCTION"
    sq = "STATE"
    sp = "LOCATION"
    sn = "SUPPLIER"
    so = "MANUFACTURER"
    sa = "INITIALGROUNDHEIGHT"
    sf = LayerName
    sb = "FINALGROUNDHEIGHT"
    sse = "DEMAND"
    ssj = "CALCULENODE"
    ssv = "NOTES"
    ssz = "TROUBLE"
   st = "INFORMATIONVALIDITY"
   su = "DATEINSTALLATION"
   ssr = "PATTERN"
    ss = "SECTOR"
    sw = "OBJECT_ID"
    a = "COMPONENT_ID"
    Dim man1 As String
        Dim man2 As String
        man1 = "STATEID"
        man2 = "X_TATE"
   ' Case UCase("+ssq+"), UCase("+ja+"), UCase("" + je + "")
       ' Case UCase("watercomponents"), UCase("sewercomponents"), UCase("draincomponents")
         
        sql = sql & " Select " + """" + jo + """" + " as" + """" + "COMPONENTE" + """" + ","
        sql = sql & " " + """" + swx + """" + " as" + """" + "TIPO" + """" + ","
        sql = sql & " " + """" + sss + """" + " as" + """" + "[ANO DE FABRICAÇÃO]" + """" + ","
        sql = sql & " " + """" + sq + """" + " as" + """" + "ESTADO" + """" + ","
        sql = sql & " " + """" + sp + """" + " as" + """" + "LOCALIZAÇÃO" + """" + ","
        sql = sql & " " + """" + sn + """" + " as" + """" + "FORNECEDOR" + """" + ","
        sql = sql & " " + """" + so + """" + " as" + """" + "FABRICANTE" + """" + ","
        sql = sql & " " + """" + sa + """" + " as" + """" + "[COTA DO TERRENO]" + """" + ","
        
        If UCase(LayerName) = "SEWERCOMPONENTS" Or UCase(LayerName) = "DRAINCOMPONENTS" Then
            sql = sql + """" + sb + """" + " as" + """" + "[COTA DO FUNDO]" + """" + ","
        End If
        
        If UCase(LayerName) = UCase("watercomponents") Then
            sql = sql + """" + sse + """" + " as" + """" + "DEMANDA" + """" + ","
            sql = sql + """" + ssj + """" + " as" + """" + "[NÓ DE CÁLCULO]" + """" + ","
        End If
        
        sql = sql + """" + st + """" + " as" + """" + "VALIDADE" + """" + ","
        sql = sql + """" + ssv + """" + " as" + """" + "Observação" + """" + ","
        sql = sql + """" + ssz + """" + " as" + """" + "[NÃO_CONFORMIDADE]" + """" + ", " + """" + su + """" + " as" + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", " + """" + ssr + """" + " as" + """" + "[PADRÃO_CONSUMO]" + """" + ", " + """" + ss + """" + " as" + """" + "[SETOR]" + """" + ""
        'sql = sql & " ANGLE,NOME_CELUL,ORIGEM_CAL,X_,Y_,COR,TAMANHO_X,TAMANHO_Y,CENT_CEL_X,CENT_CEL_Y,COR_CELULA,ESC_CEL_X,ESC_CEL_Y "
        sql = sql & " From " + """" + LayerName + """"
        
        Select Case TypeQuery
            Case 0
                sql = "select " + """" + jo + """" + " , " + """" + swx + """" + ", " + """" + sss + """" + ", " + """" + sq + """" + ", " + """" + sp + """" + ", " + """" + sn + """" + ", " + """" + so + """" + ", " + """" + sa + """" + "" & IIf(UCase(LayerName) <> UCase(" +""""+ ssq +""""+ "), ", " + """" + sb + """" + "", ", " + """" + sse + """" + ", " + """" + ssj + """" + " ") & ", " + """" + st + """" + ", " + """" + ssv + """" + "," + """" + ssz + """" + ", " + """" + su + """" + ", " + """" + ssr + """" + ", " + """" + ss + """" + " from "" & layerName & "" where " + """" + sw + """" + " = '1'"
            Case 1 'Multiple Select with Alias
                sql = sql & " Where " + """" + a + """" + " in ('" + object_id + "')"
            Case 2 'Single Select with alias
                sql = sql & " Where " + """" + a + """" + " in ('" + object_id + "')"
            Case 3 'tupla default

             sql = "Select  '0' as " + """" + " COMPONENTE " + """" + ", '0' as " + """" + "TIPO" + """" + ", '0' as " + """" + "[ANO DE FABRICAÇÃO]" + """" + ", '0' as" + """" + "ESTADO" + """" + ", '0' as " + """" + "LOCALIZAÇÃO" + """" + ", '0' as " + """" + "FORNECEDOR" + """" + ", '0' as " + """" + "FABRICANTE" + """" + ", '0' as " + """" + "[COTA DO TERRENO]" + """" + " " & IIf(LayerName <> "WATERCOMPONENTS", ", '0' as " + """" + "[COTA DO FUNDO]" + """" + "", ", '0' as " + """" + "DEMANDA" + """" + ", '0' as" + """" + " [NÓ DE CÁLCULO]" + """" + "") & ", '0' as " + """" + "VALIDADE" + """" + ", '' as " + """" + "Observação" + """" + ", '0' as " + """" + "[NÃO_CONFORMIDADE]" + """" + ", '' As " + """" + "[DATA_DE_INSTALAÇÃO]" + """" + ", '' as [PADRÃO_CONSUMO], '' as " + """" + "[SETOR]" + """" + " from " + """" + man2 + """" + " where " + """" + man1 + """" + "  ='1'"

    Case Else 'qualque plano
        sql = "SELECT * FROM " + """" + LayerName + """" + " WHERE " + """" + sw + """" + " in('" + object_id + "')"
    End Select
      End If
        
'End Select
 
 
 'MsgBox sql
 
 
 
' WritePrivateProfileString "A", "A", sql, App.Path & "\DEBUG.INI"
 
 
getPmsdp2 = convertQuery(sql, tipoProvedor)


''************* MONITORAMENTO ***************
'Close #2
'Open App.Path & "\GeoSanLog.txt" For Append As #2
'Print #2, Now & "Public Function getPmsdp - case 2 SQL = " & sql & " TIPO select " & TypeQuery
'Close #2
''***************** FIM *********************





End Function



Public Function getPmssp(LayerName As String, id_Type As Integer, object_id_ As String, conn As ADODB.Connection, tipoProvedor As Integer) As String
   'verifica existencia de dados para sub-tipo
   On Error GoTo getPmssp_err
   Dim rs As New ADODB.Recordset, sql As String
   
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
Dim a1 As String
Dim b1 As String
Dim c1 As String
Dim d1 As String
Dim e1 As String
Dim f1 As String
Dim g1 As String
Dim h1 As String
Dim i1 As String
Dim j1 As String
Dim k1 As String
Dim l1 As String

a = "OBJECT_ID_"
b = "ID_TYPE"
c = LayerName
d = "c"
e = "DATA"
f = "SUBTYPES"
g = "ID_SUBTYPE"
h = "TB_LIGACOES"
i = "HIDROMETRADO"
j = "ECONOMIAS"
k = "CONSUMO_LPS"
l = "TB_LIGACOES"
   
   If TypeConn <> 4 Then
   sql = sql & " Select count(Object_id_) From " & LayerName & "Data A "
   sql = sql & " Inner join " & LayerName & "SubTypes B On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
   sql = sql & " Where b.Id_Type =" & id_Type & " and Object_id_='" & object_id_ & "'"
   Else
   sql = sql & " Select count(" + """" + a + """" + ") From " + """" + LayerName + e + """" + ""
   sql = sql & " Inner join " + """" + LayerName + f + """" + " On  " + """" + LayerName + e + """" + "." + """" + b + """" + " = " + """" + LayerName + f + """" + "." + """" + b + """" + " and " + """" + LayerName + f + """" + "." + """" + g + """" + " = " + """" + LayerName + f + """" + "." + """" + g + """" + ""
   sql = sql & " Where " + """" + LayerName + f + """" + "." + """" + b + """" + " ='" & id_Type & "' and " + """" + a + """" + "='" & object_id_ & "'"
   End If
'      Close #1
'      Open App.Path & "\GeoSanLog.txt" For Append As #1
'      Print #1, Now & " getPmssp - " & sql
'      Close #1
   
   
   rs.Open sql, conn, adOpenDynamic, adLockOptimistic
   sql = ""

   
   If TypeConn <> 4 Then
   
   
   If Not rs.EOF Then
      If rs(0) > 0 Then
         'If tipoProvedor = 2 Then
         sql = sql & "Select '" & object_id_ & "', Selection_, Max_,Min_,DataType,B.Description_, A.Id_Type, B.Id_subType , A.Value_ as Value_Ref,"
         sql = sql & " (Select Option_"
         sql = sql & " From WaterComponentsSelections C"
         sql = sql & " Where C.id_Type = A.id_Type"
         sql = sql & " and C.Id_SubType= A.Id_subType and C.Value_=A.Value_)as value_"
         sql = sql & " From watercomponentsData A"
         sql = sql & " Left Join watercomponentsSubTypes B"
         sql = sql & " On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
         sql = sql & " Where Object_id_='" & object_id_ & "' and b.id_Type =" & id_Type
         sql = sql & " Union"
         sql = sql & " (Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as Value_Ref,"
         sql = sql & " case "
         sql = sql & " When Selection_= 1 then b.Option_"
         sql = sql & " when Selection_ <> 1 then A.DefaultValue"
         sql = sql & " End Value_"
         sql = sql & " From " & LayerName & "SubTypes A Left Join " & LayerName & "Selections B"
         sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue"
         sql = sql & " Where A.Id_Type = " & id_Type & " and a.id_subtype not in(Select id_SubType from " & LayerName & "Data where object_id_ = " & object_id_ & ")) ORDER BY Id_subType"
         ' Else
         ' sql = sql & " (Select '" & object_id_ & "', Selection_, Max_,Min_,DataType,B.Description_, A.Id_Type, B.Id_subType , A.Value_ as Value_Ref,"
         ' sql = sql & " Case Selection_ when 0 then A.Value_"
         ' sql = sql & " Else (Select Value_=Option_ From " & layerName & "Selections C"
         ' sql = sql & " where C.Id_Type= A.Id_Type and C.Id_SubType= A.Id_subType and C.Value_=A.Value_) end As Value_"
         ' sql = sql & " From " & layerName & "Data A"
         ' sql = sql & " Left Join " & layerName & "SubTypes B"
         ' sql = sql & " On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
         ' sql = sql & " Where Object_id_='" & object_id_ & "' and b.id_Type =" & id_Type & ")"
         ' sql = sql & " Union"
         ' sql = sql & " (Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as Value_Ref,"
         ' sql = sql & " case Selection_ When 1 then b.Option_ else A.DefaultValue End as Value_"
         ' sql = sql & " From " & layerName & "SubTypes A Left Join " & layerName & "Selections B"
         ' sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue"
         ' sql = sql & " Where A.Id_Type = " & id_Type & " and a.id_subtype not in(Select id_SubType from " & layerName & "Data where object_id_ = " & object_id_ & "))"
         ' End If
      
      
      Else
         sql = sql & " Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_,A.Id_Type,A.Id_subType,A.DefaultValue as Value_Ref, case Selection_ When 1 then b.Option_ else A.DefaultValue End as Value_"
         sql = sql & " From " & LayerName & "SubTypes A Left Join " & LayerName & "Selections B"
         sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue "
         sql = sql & " Where A.Id_Type = " & id_Type & " ORDER BY A.Id_subType"
      End If
   End If
   rs.Close
   
   Else
   
a = "OBJECT_ID_"
b = "a"
c = "SELECTION_"
d = "MAX_"
e = "MIN_"
f = "DATATYPE"
g = "DESCRIPTION_"
h = "ID_TYPE"
i = "ID_SUBTYPE"
j = "VALUE_"
k = "OPTION_"
l = "WATERCOMPONENTSSELECTIONS"

a1 = "WATERCOMPONENTSDATA"
b1 = "WATERCOMPONENTSSUBTYPES"
c1 = "DEFAULTVALUE"
d1 = LayerName
e1 = "SUBTYPES"
f1 = d1 + e1
g1 = d1 + c
h1 = "DATA"
i1 = d1 + h1
j1 = object_id_
Dim gg As String
gg = "Selections"






'b="  & LayerName & b1  "
'a="  & LayerName & a1  "
'c="  & LayerName & l  "

'b="  & LayerName & gg  "
'a="  & LayerName & e1  "
'c="  & LayerName & l  "




      If Not rs.EOF Then
      If rs(0) > 0 Then
         'If tipoProvedor = 2 Then
         sql = sql + "Select  + """" + object_id_ + """" +" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + ",""  & LayerName & b1  ""." + """" + g + """" + "," & LayerName & a1 + "." + """" + i + """" + "," & LayerName & b1 + "." + """" + i + """" + "," + """" + LayerName & a1 + """" + j + """" + " as " + """" + "Value_Ref" + """" + ","
         sql = sql & " (Select " + """" + k + """" + ""
         sql = sql & " From " + """" + l + """" + ""
         sql = sql & " Where " + """" + LayerName & l + """" + "." + """" + h + """" + " = " + """" + LayerName & a1 + """" + "." + """" + h + """" + ""
         sql = sql & " and ""  & LayerName & l  ""." + """" + i + """" + "= " + """" + LayerName & a1 + """" + "." + """" + i + """" + " and " + """" + LayerName & l + """" + "." + """" + j + """" + "=""  & LayerName & a1  ""." + """" + j + """" + ")as " + """" + "value_" + """" + ""
         sql = sql & " From a1"
         sql = sql & " Left Join b1"
         sql = sql & " On " + """" + LayerName & a1 + """" + "." + """" + h + """" + " = " + """" + LayerName & b1 + """" + "." + """" + h + """" + " and " + """" + LayerName & a1 + """" + "." + """" + i + """" + " = " + """" + LayerName & b1 + """" + "." + """" + i + """" + ""
         
         sql = sql & " Where " + """" + a + """" + "='" & object_id_ & "' and " + """" + b1 + """" + "." + """" + h + """" + " ='" & id_Type & "'"
         sql = sql & " Union"
         sql = sql & " (Select " + """" + object_id_ + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + "," + """" + LayerName & a1 + """" + "." + """" + g + """" + "," + """" + LayerName & a1 + """" + "." + """" + h + """" + "," + """" + LayerName & a1 + """" + "." + """" + i + """" + "," + """" + LayerName & a1 + """" + "." + """" + c1 + """" + " as " + """" + "Value_Ref" + """" + ","
         sql = sql & " case "
         sql = sql & " When " + """" + c + """" + "= '1' then " + """" + LayerName & b1 + """" + "." + """" + k + """" + ""
         sql = sql & " when " + """" + c + """" + " <> '1' then " + """" + LayerName & a1 + """" + "." + """" + c1 + """" + ""
         sql = sql & " End " + """" + j + """" + ""
         sql = sql & " From " + """" + LayerName & e1 + """" + " Left Join " + """" + LayerName & c + """" + ""
         sql = sql & " On " + """" + LayerName & e1 + """" + "." + """" + h + """" + " = " + """" + LayerName & c + """" + "." + """" + h + """" + " and " + """" + LayerName & e1 + """" + "." + """" + i + """" + " = " + """" + LayerName & c + """" + "." + """" + i + """" + " and " + """" + LayerName & c + """" + "." + """" + j + """" + " = " + """" + LayerName & e1 + """" + "." + """" + c1 + """" + ""
         sql = sql & " Where " + """" + LayerName & e1 + """" + "." + """" + h + """" + " = '" & id_Type & "' and " + """" + LayerName & e1 + """" + "." + """" + i + """" + " not in(Select " + """" + i + """" + " from " + """" + LayerName & h1 + """" + " where " + """" + a + """" + " = '" & object_id_ & "')) ORDER BY +""""+i+""""+"
         ' Else
         ' sql = sql & " (Select '" & object_id_ & "', Selection_, Max_,Min_,DataType,B.Description_, A.Id_Type, B.Id_subType , A.Value_ as Value_Ref,"
         ' sql = sql & " Case Selection_ when 0 then A.Value_"
         ' sql = sql & " Else (Select Value_=Option_ From " & layerName & "Selections C"
         ' sql = sql & " where C.Id_Type= A.Id_Type and C.Id_SubType= A.Id_subType and C.Value_=A.Value_) end As Value_"
         ' sql = sql & " From " & layerName & "Data A"
         ' sql = sql & " Left Join " & layerName & "SubTypes B"
         ' sql = sql & " On A.Id_Type = b.ID_Type and A.Id_SubType = b.ID_SubType"
         ' sql = sql & " Where Object_id_='" & object_id_ & "' and b.id_Type =" & id_Type & ")"
         ' sql = sql & " Union"
         ' sql = sql & " (Select '" & object_id_ & "',Selection_,Max_,Min_,DataType,A.Description_, A.Id_Type,A.Id_subType,A.DefaultValue as Value_Ref,"
         ' sql = sql & " case Selection_ When 1 then b.Option_ else A.DefaultValue End as Value_"
         ' sql = sql & " From " & layerName & "SubTypes A Left Join " & layerName & "Selections B"
         ' sql = sql & " On A.Id_Type = B.Id_Type and A.Id_subType = B.Id_subType and B.Value_ = A.DefaultValue"
         ' sql = sql & " Where A.Id_Type = " & id_Type & " and a.id_subtype not in(Select id_SubType from " & layerName & "Data where object_id_ = " & object_id_ & "))"
         ' End If
      
      
      Else
         sql = sql & " Select " + """" + object_id_ + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + "," + """" + LayerName & e1 + """" + "." + """" + g + """" + "," + """" + LayerName & e1 + """" + "." + """" + h + """" + "," + """" + LayerName & e1 + """" + "." + i + "," + """" + LayerName & e1 + """" + "." + c1 + " as " + """" + "Value_Ref" + """" + ", case " + """" + c + """" + " When '1' then " + """" + LayerName & gg + """" + "." + """" + h + """" + " else " + """" + LayerName & e1 + """" + "." + """" + c1 + """" + " End as " + """" + "Value_" + """" + ""
         sql = sql & " From " + """" + LayerName & e1 + """" + " Left Join " + "+""""+ LayerName & b1 +""""+"
         sql = sql & " On " + """" + a1 + """" + "." + """" + h + """" + " = " + """" + b1 + """" + "." + """" + h + """" + " and " + """" + a1 + """" + "." + """" + i + """" + " = " + """" + b1 + """" + "." + """" + i + """" + " and " + """" + b1 + """" + "." + """" + j + """" + " = " + """" + a1 + """" + "." + """" + c1 + """" + ""
         sql = sql & " Where " + """" + LayerName & e1 + """" + "." + """" + h + """" + " = '" & id_Type & "' ORDER BY " + """" + LayerName & e1 + """" + "." + """" + i + """" + ""
      End If
   End If
   rs.Close
   
   End If
   
'      Close #1
'      Open App.Path & "\GeoSanLog.txt" For Append As #1
'      Print #1, Now & " getPmssp 2 - " & sql
'      Close #1
   
   getPmssp = sql
   Exit Function
getPmssp_err:
   MsgBox Err.Description, vbCritical
End Function

Public Function convertQuery(sql As String, Tipo As Integer) As String
   If Tipo = 2 Then
      sql = Replace(sql, "[", "")
      sql = Replace(sql, "]", "")
   End If
   convertQuery = sql
   
   
   
End Function
'
'
'
Public Function RetornaNomeAtr(conn As ADODB.Connection, LayerName As String, AttributeTable As String, AttributeLink As String) As Boolean
    Dim rs As New ADODB.Recordset
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
    Dim ate As String

    a = "attr_table"
    b = "attr_link"
    c = "te_layer_table"
    d = "layer_id"
    e = "name"
    f = "te_layer"
    If TypeConn <> 4 Then
        Set rs = conn.Execute("Select attr_table,attr_link from Te_Layer_table where layer_id in(select layer_id from te_layer where name='" & LayerName & "')")
    Else
        Set rs = conn.Execute("Select " + """" + a + """" + "," + """" + b + """" + " from " + """" + c + """" + " where " + """" + d + """" + " in(select " + """" + d + """" + " from " + """" + f + """" + " where " + """" + e + """" + "='" & LayerName & "')")
    End If
    ate = "Select " + """" + a + """" + "," + """" + b + """" + " from " + """" + c + """" + " where " + """" + d + """" + " in(select " + """" + d + """" + " from " + """" + f + """" + " where " + """" + e + """" + "='" & LayerName & "')"
    'MsgBox "ARQUIVO DEBUG SALVO"
    ' WritePrivateProfileString "A", "A", ate, App.Path & "\DEBUG.INI"
    If Not rs.EOF Then
        AttributeLink = rs(1).Value
        AttributeTable = rs(0).Value
        RetornaNomeAtr = True
    End If
    rs.Close
    Set rs = Nothing
End Function






