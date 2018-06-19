VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerificaConectividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificar Conectividade de Redes"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Excluir Rede de Agua Nula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   735
      TabIndex        =   9
      Top             =   3030
      Width           =   4800
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Listar Rede de Agua Nula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   255
      TabIndex        =   8
      Top             =   2715
      Width           =   5250
   End
   Begin VB.CheckBox chkDeletaCompOrfao 
      Caption         =   "Excluir Componente sem Rede de Agua"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   750
      TabIndex        =   7
      Top             =   2235
      Width           =   4860
   End
   Begin VB.CheckBox chkListaCompOrfao 
      Caption         =   "Listar Compomentes sem Rede de Agua"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   1920
      Width           =   5400
   End
   Begin VB.CheckBox chkDistancia 
      Caption         =   "Listar e Corrigir distâncias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   285
      TabIndex        =   5
      Top             =   1425
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione as ações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   120
      TabIndex        =   10
      Top             =   1005
      Width           =   5655
   End
   Begin MSComctlLib.ProgressBar ProgBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4875
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caminho do Relatório"
      Height          =   750
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   5670
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   255
         Width           =   5385
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   4560
      TabIndex        =   1
      Top             =   4830
      Width           =   1185
   End
   Begin VB.CommandButton cmdInciar 
      Caption         =   "Iniciar"
      Height          =   390
      Left            =   3285
      TabIndex        =   0
      Top             =   4830
      Width           =   1185
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
      Left            =   135
      OleObjectBlob   =   "frmVerificaConectividade.frx":0000
      Top             =   4830
   End
End
Attribute VB_Name = "frmVerificaConectividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim blnPontoCriado As Boolean
Dim ListComponents As String
Dim IdTabelaPoints As Integer
Dim IdTabelaLines As Integer

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    
   TeDatabase1.Provider = TpConexao 'usa a variável global para identificar o tipo de conexão
   TeDatabase1.Connection = Conn 'usa a variável global para identificar a conexão
   TeDatabase1.setCurrentLayer ("WATERLINES")
    
    Text1.Text = "C:\Arquivos de Programas\Geosan\Controles\DiagosticoRede" & Format(Now, "HHMM") & ".txt"
End Sub

Private Sub cmdInciar_Click()

On Error GoTo Trata_Erro
  



   Dim rsVBL As New ADODB.Recordset
   Dim rsVBP As New ADODB.Recordset
   Dim rsBusca As New ADODB.Recordset
   Dim rsLayer As New ADODB.Recordset
   Dim rsLinha As New ADODB.Recordset
   Dim VALID As Boolean
   Dim strsql As String

   Dim rsInitial As New ADODB.Recordset
   Dim rsInitial2 As New ADODB.Recordset
   Dim rsFinal As New ADODB.Recordset
   Dim rsFinal2 As New ADODB.Recordset
   Dim rsSemPoints As New ADODB.Recordset
   
   Dim rslinha1 As New ADODB.Recordset
   Dim rslinha2 As New ADODB.Recordset
   Dim QTDPT As Integer
   Dim retorno As Double
   Dim XL1 As Double, XL2 As Double, YL1 As Double, YL2 As Double
   
   Dim strXL1 As String, strXL2 As String, strYL1 As String, strYL2 As String
   
   Dim LINHA1 As String
   Dim LINHA2 As String
   Dim CONTALINHAS As Integer
   

   
   Me.MousePointer = vbHourglass
   
   'IDENTIFICA QUAL TABELA LINES O LAYER WATERLINES REGISTRA AS LOCALIZAÇÕES
   strsql = "SELECT LAYER_ID,NAME FROM TE_LAYER WHERE NAME = '" & "WATERLINES" & "'"
   Set rsLayer = Conn.execute(strsql)
   If rsLayer.EOF = True Then
      IdTabelaLines = rsLayer!layer_id
   Else
      MsgBox "Não localizada a tabela de geometrias 'LINES##' da tabela WATERLINES", vbExclamation, ""
      Exit Sub
   End If
   
   
   'IDENTIFICA QUAL TABELA POINTS O LAYER WATERCOMPONENTS REGISTRA AS LOCALIZAÇÕES
   strsql = "SELECT LAYER_ID,NAME FROM TE_LAYER WHERE NAME = '" & "WATERCOMPONENTS" & "'"
   Set rsLayer = Conn.execute(strsql)
   If rsLayer.EOF = True Then
      IdTabelaPoints = rsLayer!layer_id
   Else
      MsgBox "Não localizada a tabela de geometrias 'Points##' da tabela WATERCOMPONENTS", vbExclamation, ""
      Exit Sub
   End If
   
   
   Open Text1.Text For Output As #5 ' ABRE O ARQUIVO TEXTO PARA LOG
   
   
   If Me.chkDistancia.value = True Then
      corrigeDistancias 'inicializa o corretor de distancias
   End If
      
      
   'EXCLUI AS LINHAS QUE NÃO POSSUEM GEOMETRIA NA TABELA LINES1
   strsql = "SELECT OBJECT_ID_ FROM WATERLINES WHERE OBJECT_ID_ NOT IN (SELECT OBJECT_ID FROM LINES" & IdTabelaLines & ")"
   Set rsLinha = Conn.execute(strsql)
   If rsLinha.EOF = False Then
      Do While Not rsLinha.EOF
         'VERIFICADO QUE QUANDO A LINHA NÃO POSSUI GEOMETRIA, ELA NÃO APARECE NO MAPA
         'E POR ISSO O USUÁRIO NÃO PODE MANIPULA-LA
         Conn.execute ("DELETE FROM WATERLINES WHERE OBJECT_ID_ ='" & rsLinha!Object_id_ & "'")
         Print #5, "Linha " & rsLinha!Object_id_ & " SEM GEOMETRIA, EXCLUÍDA."
         rsLinha.MoveNext
      Loop
   End If
   
   
   'EXCLUI AS GEOMETRIAS DE LINHAS QUE NÃO TEM LINHAS NA TABELA WATERLINES
   strsql = "SELECT OBJECT_ID FROM LINES" & IdTabelaLines & " WHERE OBJECT_ID NOT IN (SELECT OBJECT_ID_ FROM WATERLINES)"
   Set rsLinha = Conn.execute(strsql)
   If rsLinha.EOF = False Then
      Do While Not rsLinha.EOF
         Conn.execute ("DELETE FROM LINES1 WHERE OBJECT_ID ='" & IdTabelaLines & "'")
         Print #5, "DESENHO DE Linha COD " & rsLinha!object_id & " SEM INFORMAÇÃO DE CADASTRO, EXCLUÍDA."
         rsLinha.MoveNext
      Loop
   End If
   
   
   
   
   'COM O SELECT ABAIXO OBTEM-SE UMA LISTA DOS COMPONENTES DE REDE QUE EXISTEM NA TABELA WATERCOMPONENTES MAS NÃO TEM INFORMAÇÃO GEOGRAFICA
   
   If TpConexao = 1 Then 'CASO SQL SERVER, CARREGA O RECORDSET DIRETO POR UM COMANDO
   
      strsql = "SELECT OBJECT_ID_ FROM WATERCOMPONENTS WHERE OBJECT_ID_ NOT IN (SELECT OBJECT_ID FROM POINTS" & IdTabelaPoints & ")"
      Set rsSemPoints = Conn.execute(strsql)
   
   Else 'CASO ORACLE, FAZ UM LOOP BUSCANDO OS PONTOS NÃO ENCONTRADOS E PASSA A LISTA PARA O COMANDO DO RECORDSET
      
      LISTA_COMPONENTE_SEM_GEOMETRIA 'CARREGA UM ARRAY QUE SERÁ USADO NO LUGAR DO RECORDSET
   
      strsql = "SELECT OBJECT_ID_ FROM WATERCOMPONENTS WHERE OBJECT_ID_ IN (" & ListComponents & ")"

   End If
   Set rsSemPoints = Conn.execute(strsql)
   
   
   
   If rsSemPoints.EOF = False Then
      Do While Not rsSemPoints.EOF = True
         id_componente = rsSemPoints!Object_id_
         
         'VERIFICANDO A QUAL LINHA ESTE COMPONENTE É COMPONENTE INICIAL
         Set rsInitial = Conn.execute("SELECT LINE_ID,OBJECT_ID_,INITIALCOMPONENT FROM WATERLINES WHERE INITIALCOMPONENT ='" & id_componente & "'")
         
         If rsInitial.EOF = False Then
            'chegando a este ponto significa que o componente é inicial de 1 ou mais linhas
            LINHA1 = rsInitial!Object_id_ 'carrega em LINHA1 o id da linha que o componente é inicial
            
            retorno = TeDatabase1.getPointOfLine(0, LINHA1, 0, XL1, YL1) 'retorna em XL1 e YL1 as coordenadas iniciais da linha

            'VERIFICANDO SE O COMPONENTE É TAMBEM FINAL DE ALGUMA OUTRA LINHA
            Set rsFinal = Conn.execute("SELECT LINE_ID,OBJECT_ID_,FINALCOMPONENT FROM WATERLINES WHERE FINALCOMPONENT ='" & id_componente & "'AND OBJECT_ID_ <> '" & LINHA1 & "'")
            If rsFinal.EOF = False Then
               LINHA2 = rsFinal!Object_id_
               'chegando a este ponto significa que o componente é inicial e final de duas OU mais linhas
               'ANALISAR AS 2 LINHAS
               
               'FAZER A PESQUISA PARA SABER O X,Y DAS LINHAS
               
               QTDPT = TeDatabase1.getQuantityPointsLine(0, LINHA2) 'retorna número de pontos que compõem a linhA para pegar as coordenadas do ultimo ponto
               If QTDPT >= 2 Then
                  retorno = TeDatabase1.getPointOfLine(0, LINHA2, QTDPT - 1, XL2, YL2) 'retorna em XL2 e YL2 as coordenadas finais da linha
               End If
              

               If XL1 = XL2 And YL1 = YL2 Then
                  strsql = "insert into points2 (object_id,x,y) values ('" & id_componente & "'," & XL1 & "," & YL1 & "')"
                  Conn.execute (strsql)
                  Print #5, "Componente " & id_componente & " localizado com sucesso!"
                  
               Else
                  'MsgBox "Valor inconsistente para o componente de rede nº " & id_componente & " contido nas linhas " & LINHA1 & " e " & LINHA2 & "." & Chr(13) & Chr(13) & "Não foi possivel corrigir automaticamente.", vbExclamation, ""
                  Print #5, "Valor inconsistente para o componente de rede nº " & id_componente & " contido nas linhas " & LINHA1 & " e " & LINHA2 & ". Não foi possivel corrigir automaticamente."
               End If
            
            Else
               'chegando a este ponto significa que o componente é somente inicial de duas ou mais linhas
               'ANALIZAR A LINHA QUE ELE É INICIAL

               CONTALINHAS = 1
               rsInitial.MoveNext
               Do While Not rsInitial.EOF = True
                  CONTALINHAS = CONTALINHAS + 1
               Loop
               If CONTALINHAS = 1 Then 'O PONTO ESTÁ CONECTADO A SOMENTE 1 LINHA
               
                  'retorno = TeDatabase1.getPointOfLine(0, rsInitial!Object_id_, 0, XL1, YL1)
                  
                  strXL1 = Replace(XL1, ",", ".") 'converte o valor double do XL1
                  strYL1 = Replace(YL1, ",", ".") 'converte o valor double do YL1
                  
                  strsql = "insert into points2 (object_id,x,y) values ('" & id_componente & "'," & strXL1 & "," & strYL1 & ")"
                  
                  Conn.execute (strsql)
                  Print #5, "Componente " & id_componente & " localizado com sucesso!"
                  
               
               Else 'O PONTO ESTÁ CONECTADO A MAIS DE 1 LINHA
                  Set rsInitial2 = Conn.execute("SELECT LINE_ID,OBJECT_ID_,INITIALCOMPONENT FROM WATERLINES WHERE INITIALCOMPONENT ='" & id_componente & "' AND OBJECT_ID_ <> '" & LINHA1 & "'")
                  If rsInitial2.EOF = False Then
                     LINHA2 = rsInitial2!Object_id_
                     retorno = TeDatabase1.getPointOfLine(0, rsInitial2!Object_id_, 0, XL2, YL2)
                     
                     If XL1 = XL2 And YL1 = YL2 Then
                        strXL1 = Replace(XL1, ",", ".") 'converte o valor double do XL1
                        strYL1 = Replace(YL1, ",", ".") 'converte o valor double do YL1
                        strsql = "insert into points2 (object_id,x,y) values ('" & id_componente & "'," & XL1 & "," & YL1 & "')"
                        Conn.execute (strsql)
                        Print #5, "Componente " & id_componente & " localizado com sucesso!"
                     Else
                        
                        'MsgBox "Valores inconsistentes para a linha " & LINHA1 & " e linha " & LINHA2 & "." & Chr(13) & Chr(13) & "Não foi possivel corrigir automaticamente.", vbExclamation, ""
                        Print #5, "Valores inconsistentes para a linha " & LINHA1 & " e linha " & LINHA2 & ". Não foi possivel corrigir automaticamente."
                     End If
                  Else
                  
                     'MsgBox "Valores inconsistentes para a linha " & LINHA1 & "." & Chr(13) & Chr(13) & "Não foi possivel corrigir automaticamente.", vbExclamation, ""
                     Print #5, "Valores inconsistentes para a linha " & LINHA1 & ". Não foi possivel corrigir automaticamente."
                  End If
                  
               End If
               
            End If
            
         Else
            'chegando a este ponto significa que o componente não é inicial de nenhuma linha
            'verificando se ele é final de alguma linha
            Set rsFinal = Conn.execute("SELECT LINE_ID,OBJECT_ID_,FINALCOMPONENT FROM WATERLINES WHERE FINALCOMPONENT ='" & id_componente & "'")
            If rsFinal.EOF = False Then
               'chegando a este ponto significa que o componente é somente final de duas ou mais linhas
            
               LINHA1 = rsFinal!Object_id_
               retorno = TeDatabase1.getPointOfLine(0, LINHA1, 0, XL1, YL1)
            
               CONTALINHAS = 1
               rsFinal.MoveNext
               Do While Not rsFinal.EOF = True
                  CONTALINHAS = CONTALINHAS + 1
               Loop
               
               If CONTALINHAS = 1 Then 'O PONTO ESTÁ CONECTADO A SOMENTE 1 LINHA
               
                  
                  strXL1 = Replace(XL1, ",", ".") 'converte o valor double do XL1
                  strYL1 = Replace(YL1, ",", ".") 'converte o valor double do YL1
                  
                  strsql = "insert into points" & IdTabelaPoints & " (object_id,x,y) values ('" & id_componente & "'," & strXL1 & "," & strYL1 & ")"
                  
                  Conn.execute (strsql)
                  Print #5, "Componente " & id_componente & " localizado com sucesso!"
               
               Else 'O PONTO ESTÁ CONECTADO A MAIS DE 1 LINHA
                  Set rsFinal2 = Conn.execute("SELECT LINE_ID,OBJECT_ID_,INITIALCOMPONENT FROM WATERLINES WHERE INITIALCOMPONENT ='" & id_componente & "' AND OBJECT_ID_ <> '" & LINHA1 & "'")
                  If rsFinal2.EOF = False Then
                     
                     LINHA2 = rsFinal2!Object_id_
                     retorno = TeDatabase1.getPointOfLine(0, rsFinal2!Object_id_, 0, XL2, YL2)
                     
                     If XL1 = XL2 And YL1 = YL2 Then
                        strsql = "insert into points2 (object_id,x,y) values ('" & id_componente & "'," & XL1 & "," & YL1 & "')"
                        Conn.execute (strsql)
                        Print #5, "Componente " & id_componente & " localizado com sucesso!"
                     Else
                        
                        Print #5, "Valores inconsistentes para a linha " & LINHA1 & " e linha " & LINHA2 & "." & Chr(13) & Chr(13) & "Não foi possivel corrigir automaticamente."
                     End If
                  Else
                  
                     Print #5, "Valores inconsistentes para a linha " & LINHA1 & "." & Chr(13) & Chr(13) & "Não foi possivel corrigir automaticamente."
                     
                  End If
                  
               End If
            
            
            Else
               'chegando a este ponto significa que o componente não é inicial nem final de linhas
               
               If chkDeletaCompOrfao.value = True Then
                  
                  strCMD = "DELETE FROM WATERCOMPONENTS WHERE OBJECT_ID_ ='" & id_componente & "'"
                  Conn.execute (strCMD)
                  Print #5, "Componente de rede " & id_componente & " sem conexões. >> Excluído."
               
               Else
                  Print #5, "Componente de rede " & id_componente & " sem conexões. >> Não Excluído."
               
               End If

            
            End If
               
         End If
         rsSemPoints.MoveNext
      Loop
   End If
   
   Print #5, ""
   Print #5, " * * * * FIM DE VERIFICAÇÃO DE GEOMETRIAS * * * *"
   Print #5, ""
   
   Set rsVBL = Conn.execute("SELECT OBJECT_ID_ AS COD,INITIALCOMPONENT AS INI FROM WATERLINES ORDER BY INITIALCOMPONENT")
   If rsVBL.EOF = False Then
       Set rsVBP = Conn.execute("SELECT COMPONENT_ID AS COMPONENTE FROM WATERCOMPONENTS ORDER BY COMPONENT_ID")
       'VALIDANDO TODOS OS COMPONENTES INITIAL DA WATERLINES
       If rsVBP.EOF = False Then
           Do While Not rsVBP.EOF = True And Not rsVBL.EOF = True
               If rsVBP!COMPONENTE = rsVBL!ini Then 'validado
                   rsVBL.MoveNext
                   VALID = True
               ElseIf rsVBP!COMPONENTE < rsVBL!ini Then
                   rsVBP.MoveNext
                   VALID = False
               Else
                   Print #5, "Componente Inicial:"; Tab(21); rsVBL!ini; Tab(31); "da linha"; Tab(40); rsVBL!COD; Tab(50); "NÃO ENCONTRADO."
                   
                   'CriaComponenteDefault (rsVBL!ini)
                   If blnPontoCriado = True Then
                       Print #5, "Componente " & rsVBL!ini & " POSSUI GEOMETRIA E FOI CRIADO AUTOMATICAMENTE."
                   Else
                       Print #5, "Componente " & rsVBL!ini & " NÃO PODE SER CRIADO AUTOMATICAMENTE."
                   End If
                   
                   rsVBL.MoveNext
               End If
               If rsVBP.EOF = True Then
                   If VALID = False Then
                       Do While Not rsVBL.EOF = True
                           Print #5, "Componente Inicial:"; Tab(21); rsVBL!ini; Tab(31); "da linha"; Tab(40); rsVBL!COD; Tab(50); "não encontrado!"
                           
                           CriaComponenteDefault (rsVBL!ini)
                           If blnPontoCriado = True Then
                               Print #5, "Componente " & rsVBL!ini & " POSSUI GEOMETRIA E FOI CRIADO AUTOMATICAMENTE."
                           Else
                               Print #5, "Componente " & rsVBL!ini & " NÃO PODE SER CRIADO AUTOMATICAMENTE."
                           End If
                           rsVBL.MoveNext
                       Loop
                   End If
                   Exit Do
               End If
           Loop
       End If
   End If
   Print #5, ""
   Print #5, " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
   Print #5, ""
   Set rsVBL = Conn.execute("SELECT OBJECT_ID_ AS COD,FINALCOMPONENT AS FIM FROM WATERLINES ORDER BY FINALCOMPONENT")
   If rsVBL.EOF = False Then
       Set rsVBP = Conn.execute("SELECT COMPONENT_ID AS COMPONENTE FROM WATERCOMPONENTS ORDER BY COMPONENT_ID")
       'VALIDANDO TODOS OS COMPONENTES FINAL DA WATERLINES
       If rsVBP.EOF = False Then
           Do While Not rsVBP.EOF = True And Not rsVBL.EOF = True
               If rsVBP!COMPONENTE = rsVBL!fim Then 'validado
                   rsVBL.MoveNext
                   VALID = True
               ElseIf rsVBP!COMPONENTE < rsVBL!fim Then
                   rsVBP.MoveNext
                   VALID = False
               Else
                   Print #5, "Componente Final:"; Tab(21); rsVBL!fim; Tab(31); "da linha"; Tab(40); rsVBL!COD; Tab(50); "NÃO ENCONTRADO."
                   
                   CriaComponenteDefault (rsVBL!fim)
                   If blnPontoCriado = True Then
                       Print #5, "Componente " & rsVBL!fim & " POSSUI GEOMETRIA E FOI CRIADO AUTOMATICAMENTE."
                   Else
                       Print #5, "Componente " & rsVBL!fim & " NÃO PODE SER CRIADO AUTOMATICAMENTE."
                   End If
   
                   rsVBL.MoveNext
               End If
               If rsVBP.EOF = True Then
                   If VALID = False Then
                       Do While Not rsVBL.EOF = True
                           Print #5, "Componente Final:"; Tab(21); rsVBL!fim; Tab(31); "da linha"; Tab(40); rsVBL!COD; Tab(50); "não encontrado!"
                           
                           CriaComponenteDefault (rsVBL!fim)
                           If blnPontoCriado = True Then
                              Print #5, "Componente " & rsVBL!fim & " POSSUI GEOMETRIA E FOI CRIADO AUTOMATICAMENTE."
                           Else
                              Print #5, "Componente " & rsVBL!fim & " NÃO PODE SER CRIADO AUTOMATICAMENTE."
                           End If
                           
                           rsVBL.MoveNext
                       Loop
                   End If
                   Exit Do
               End If
           Loop
       End If
   End If
   
   Close #5 'FECHA O ARQUIVO TEXTO PARA LOG
   
   rsVBL.Close
   rsVBP.Close
   Me.MousePointer = vbDefault
   MsgBox "foi gerado em " & Text1.Text & " um relatório contendo o diagnóstico de rede.", vbInformation, ""
   Unload Me
   

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
Else
   'Resume
   Me.MousePointer = vbDefault
   Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmVerificaConectividade - cmdInciar_Click - " & Err.Number & " - " & Err.Description
   Close #1
   MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation

End If

End Sub
Private Function CriaComponenteDefault(ident As Long) As Boolean

On Error GoTo Trata_Erro
   
   Dim rsBusca As New ADODB.Recordset
   Dim rsLayer As New ADODB.Recordset
   Dim strsql As String

   Set rsBusca = Conn.execute("SELECT * FROM POINTS" & IdTabelaPoints & " WHERE OBJECT_ID = '" & ident & "'")
   If rsBusca.EOF = False Then 'A GEOMETRIA DO PONTO EXISTE
         
      Set rsBusca = Conn.execute("SELECT * FROM WATERCOMPONENTS WHERE OBJECT_ID_ = '" & ident & "'")
      If rsBusca.EOF = True Then
         Dim strCMD As String
         strCMD = "SET IDENTITY_INSERT WATERCOMPONENTS ON;"
         strCMD = strCMD & "INSERT INTO WATERCOMPONENTS (COMPONENT_ID,OBJECT_ID_,SECTOR) VALUES (" & ident & "," & ident & ",999);"
         strCMD = strCMD & "SET IDENTITY_INSERT WATERCOMPONENTS OFF"
         'MsgBox strCMD
         Conn.execute (strCMD) 'insere o ponto na watercomponents
         blnPontoCriado = True
      Else ' O PONTO JA FOI CRIADO NO PROCESSO ANTERIOR
         blnPontoCriado = True
      End If
         
   Else 'A GEOMETRIA DO PONTO NÃO EXISTE
      blnPontoCriado = False
   End If


Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
Else
   blnPontoCriado = False
   Exit Function
End If


End Function

Private Function LISTA_COMPONENTE_SEM_GEOMETRIA() 'carrega em ListComponents os componentes não encontrados na tabela points
   
   'FUNÇÃO PARA VERIFICAR SE OS OBJECT_ID NA TABELA POINTS POSSUEM UM OBJECT_ID_ NA WATERCOMPONENTS
   'CRIA UMA LISTA DE ID's DE WATERCOMPONENTS QUE NÃO FORAM ENCONTRADOS
   Dim rsWTC As New ADODB.Recordset
   Dim rsPOINT As New ADODB.Recordset
   
   Set rsWTC = Conn.execute("SELECT OBJECT_ID_ AS ID_COMP, LENGTH(OBJECT_ID_) AS TAM FROM WATERCOMPONENTS ORDER BY TAM, OBJECT_ID_")
   
   'SELECT OBJECT_ID_, LENGTH(OBJECT_ID_) AS TAM from WATERCOMPONENTS ORDER BY TAM, OBJECT_ID_
   
   If rsWTC.EOF = False Then
       Set rsPOINT = Conn.execute("SELECT OBJECT_ID AS ID_POINT, LENGTH(OBJECT_ID) AS TAM FROM POINTS" & IdTabelaPoints & " ORDER BY TAM, OBJECT_ID")
       
       Open Text1.Text For Append As #4
       'COMPARANDO OS ID's
       
       If rsPOINT.EOF = False Then
           Do While Not rsPOINT.EOF = True And Not rsWTC.EOF = True
               If CDbl(rsPOINT!ID_POINT) = CDbl(rsWTC!id_comp) Then 'validado
                   rsWTC.MoveNext
                   VALID = True
               ElseIf CDbl(rsPOINT!ID_POINT) < CDbl(rsWTC!id_comp) Then
                   rsPOINT.MoveNext
                   VALID = False
               Else
                   If ListComponents = "" Then
                        ListComponents = rsWTC!id_comp
                   Else
                        ListComponents = ListComponents & "," & rsWTC!id_comp
                   End If
                   rsWTC.MoveNext
               End If
               If rsVBP.EOF = True Then
                  If VALID = False Then
                     Do While Not rsWTC.EOF = True
                        
                        If ListComponents = "" Then
                           ListComponents = rsWTC!id_comp
                        Else
                           ListComponents = ListComponents & "," & rsWTC!id_comp
                        End If

                        rsWTC.MoveNext
                     Loop
                     End If
                  Exit Do
               End If
           Loop
       End If
   End If

Close #4

End Function

Private Function corrigeDistancias()
On Error GoTo Trata_Erro

   Dim rs As New ADODB.Recordset
   Dim rsWaterlines As New ADODB.Recordset
   Dim qtdPtosLinha As Integer
   Dim retorno As Long
   Dim dblDistancia As Double
   Dim x_1 As Double
   Dim y_1 As Double
   Dim x_2 As Double
   Dim y_2 As Double
   Dim i As Integer
   Dim strDistancia As String
   Dim strLinha As String
   Dim contacorrigidos As Long
   
   contacorrigidos = 0

   strsql = "SELECT LEN(OBJECT_ID_) AS TAM, OBJECT_ID_, LENGTHCALCULATED FROM WATERLINES ORDER BY TAM, OBJECT_ID_"
   Set rsWaterlines = Conn.execute(strsql)
   
   rsWaterlines.MoveFirst 'processo para definir valores para a barra de progresso
   retorno = 0
   Do While Not rsWaterlines.EOF
      retorno = retorno + 1
      rsWaterlines.MoveNext
   Loop
   ProgBar1.Max = retorno
   ProgBar1.value = 1
   
   Me.MousePointer = vbHourglass
   
   Open Text1.Text For Append As #4
   Print #4, "RECALCULANDO DISTANCIAS >> INICIO >>"
   rsWaterlines.MoveFirst
   Do While Not rsWaterlines.EOF
      DoEvents
      strLinha = rsWaterlines!Object_id_
      dblDistancia = 0
      
      qtdPtosLinha = TeDatabase1.getQuantityPointsLine(0, strLinha) 'retorna número de pontos que compõem a linha. se maior que 2 significa que tem vertices
      
      If qtdPtosLinha > 2 Then 'existem vértices na linha
               
         retorno = TeDatabase1.getPointOfLine(0, strLinha, 0, x_1, y_1) 'retorna em x_1 e Y_1 as coordenadas do inicio da linha
         
         For i = 1 To qtdPtosLinha - 1
            
            retorno = TeDatabase1.getPointOfLine(0, strLinha, i, x_2, y_2) 'retorna em x_2 e Y_2 as coordenadas do proximo ponto
            
            dblDistancia = dblDistancia + DistanceBetween(x_1, y_1, x_2, y_2) 'carrega em distancia a soma
            
            x_1 = x_2
            y_1 = y_2
            
         Next
      
      ElseIf qtdPtosLinha = 2 Then
         
         retorno = TeDatabase1.getPointOfLine(0, strLinha, 0, x_1, y_1) 'retorna em x_1 e Y_1 as coordenadas do inicio da linha
         retorno = TeDatabase1.getPointOfLine(0, strLinha, 1, x_2, y_2) 'retorna em x_2 e Y_2 as coordenadas do fim da linha
         
         dblDistancia = DistanceBetween(x_1, y_1, x_2, y_2) 'carrega em distancia a soma
         
      End If
      
      dblDistancia = Round(dblDistancia, 2)
      
      If rsWaterlines!LENGTHCALCULATED <> dblDistancia Then
         
         strDistancia = Replace(dblDistancia, ",", ".")
         
         Conn.execute "UPDATE WATERLINES SET LENGTHCALCULATED=" & strDistancia & " WHERE OBJECT_ID_=" & strLinha
         
         'SELECT ORDER BY DESC PARA PEGAR A ULTIMA LINHA DO GRUPO DE 3
         Set rs = Conn.execute("SELECT GEOM_ID,OBJECT_ID,TEXT_VALUE FROM TEXTS1 WHERE OBJECT_ID = " & strLinha & " ORDER BY GEOM_ID DESC")
         rs.Close
         rs.Open "SELECT GEOM_ID,OBJECT_ID,TEXT_VALUE FROM TEXTS1 WHERE OBJECT_ID = " & strLinha & " ORDER BY GEOM_ID DESC", Conn, adOpenKeyset, adLockOptimistic
         If rs.EOF = False Then
            rs!TEXT_VALUE = dblDistancia
            rs.Update
         End If
         rs.Close
         
         Print #4, "Linha"; Tab(11); strLinha; Tab(21); "RECALCULADO DE"; Tab(38); rsWaterlines!LENGTHCALCULATED; Tab(50); "PARA"; Tab(55); dblDistancia
         contacorrigidos = contacorrigidos + 1
      End If
      
      If ProgBar1.Max > ProgBar1.value Then ProgBar1.value = ProgBar1.value + 1 Else ProgBar1.value = 1
      
      rsWaterlines.MoveNext
      
   Loop
   Me.MousePointer = Default
   Print #4, ""
   Print #4, "RECALCULADAS " & contacorrigidos & " DE UM TOTAL DE " & ProgBar1.Max & " LINHAS."
   Print #4, "RECALCULANDO DISTANCIAS >> FIM >>"
   Close #4

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
ElseIf Err.Number = 55 Then
   Close #4
   Resume
Else
   
   MsgBox Err.Number & " " & Err.Description
   Me.MousePointer = vbDefault
   Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - Private Function corrigeDistancias() - " & Err.Number & " - " & Err.Description
   Close #1
   MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation

End If
End Function

Private Function DistanceBetween(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  ' Calculate the distance between two points, given their X/Y coordinates.
  
  ' The short version...
  DistanceBetween = Sqr((Abs(X2 - X1) ^ 2) + (Abs(Y2 - Y1) ^ 2))
  
End Function
