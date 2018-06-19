VERSION 5.00
Begin VB.Form FrmCadastroRamalAgua 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Ramal de Água"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Ligações Fictícias"
      Height          =   1530
      Left            =   150
      TabIndex        =   33
      Top             =   6210
      Width           =   4845
      Begin VB.Frame Frame7 
         Caption         =   "Consumo (médio/ligação)"
         Height          =   1005
         Left            =   2100
         TabIndex        =   37
         Top             =   345
         Width           =   2580
         Begin VB.TextBox txtConsumoFicticia 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1215
            TabIndex        =   40
            Text            =   "0.00"
            ToolTipText     =   "Informe o consumo médio de uma ligação"
            Top             =   435
            Width           =   1140
         End
         Begin VB.OptionButton optMetroCubico 
            Caption         =   "M³/Mês"
            Height          =   285
            Left            =   150
            TabIndex        =   39
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optLitrosSegundo 
            Caption         =   "LPS"
            Height          =   285
            Left            =   150
            TabIndex        =   38
            Top             =   615
            Width           =   870
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Quantidade"
         Height          =   1005
         Left            =   480
         TabIndex        =   34
         Top             =   345
         Width           =   1320
         Begin VB.TextBox txtQtd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   180
            TabIndex        =   36
            Text            =   "0"
            Top             =   405
            Width           =   495
         End
         Begin VB.PictureBox UpDown2 
            Height          =   495
            Left            =   765
            ScaleHeight     =   435
            ScaleWidth      =   195
            TabIndex        =   35
            Top             =   330
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton cmdConsultarLigacoes 
      Caption         =   "Consultar consumo"
      Height          =   390
      Left            =   5130
      TabIndex        =   32
      Top             =   7845
      Width           =   1740
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   390
      Left            =   6930
      TabIndex        =   31
      Top             =   7845
      Width           =   1065
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Salvar"
      Height          =   390
      Left            =   8040
      TabIndex        =   30
      Top             =   7845
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pré Filtro"
      Height          =   1155
      Left            =   150
      TabIndex        =   26
      Top             =   150
      Width           =   8925
      Begin VB.OptionButton optConsumidor 
         Caption         =   "Consumidor"
         Height          =   225
         Left            =   6255
         TabIndex        =   6
         Top             =   360
         Width           =   2085
      End
      Begin VB.OptionButton optEndereço 
         Caption         =   "Endereço"
         Height          =   225
         Left            =   3945
         TabIndex        =   4
         Top             =   360
         Width           =   2220
      End
      Begin VB.OptionButton optNumLigacao 
         Caption         =   "Inscrição"
         Height          =   240
         Left            =   105
         TabIndex        =   0
         Top             =   345
         Width           =   1800
      End
      Begin VB.CommandButton cmdAtivaFiltro 
         Caption         =   ">>"
         Default         =   -1  'True
         Height          =   360
         Left            =   8415
         TabIndex        =   8
         Top             =   600
         Width           =   405
      End
      Begin VB.OptionButton optInscricao 
         Caption         =   "Ligação / Matrícula"
         Height          =   210
         Left            =   1995
         TabIndex        =   2
         Top             =   360
         Width           =   1740
      End
      Begin VB.TextBox txtConsumidor 
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
         Left            =   6270
         TabIndex        =   7
         Top             =   600
         Width           =   2085
      End
      Begin VB.TextBox txtEndereco 
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
         Left            =   3930
         TabIndex        =   5
         Top             =   600
         Width           =   2265
      End
      Begin VB.TextBox txtNumLigacao 
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
         Left            =   90
         TabIndex        =   1
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtInscricao 
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
         Left            =   2010
         TabIndex        =   3
         Top             =   600
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Selecione os consumidores associados ao ramal"
      Height          =   2475
      Left            =   150
      TabIndex        =   18
      Top             =   1620
      Width           =   8925
      Begin VB.CommandButton cmdPesquisaLigacoes 
         Caption         =   "Pesquisar Ligações"
         Height          =   375
         Left            =   6450
         TabIndex        =   25
         Top             =   2775
         Width           =   2295
      End
      Begin VB.PictureBox lvLigacoes 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   150
         ScaleHeight     =   2040
         ScaleWidth      =   8565
         TabIndex        =   9
         Top             =   240
         Width           =   8625
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   270
         Left            =   180
         TabIndex        =   27
         Top             =   2715
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Ramal"
      Height          =   1860
      Left            =   150
      TabIndex        =   19
      Top             =   4185
      Width           =   8940
      Begin VB.TextBox txtProfundidade 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5760
         TabIndex        =   13
         Top             =   660
         Width           =   1845
      End
      Begin VB.TextBox txtComprimentoRamal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5760
         TabIndex        =   12
         Top             =   300
         Width           =   1845
      End
      Begin VB.Frame Frame2 
         Caption         =   "Posicionamento em relação ao lote"
         Height          =   675
         Left            =   135
         TabIndex        =   20
         Top             =   1035
         Width           =   8685
         Begin VB.OptionButton optEsquerdo 
            Caption         =   "Esquerdo"
            Height          =   225
            Left            =   2745
            TabIndex        =   15
            Top             =   300
            Width           =   975
         End
         Begin VB.OptionButton optCentro 
            Caption         =   "Centro"
            Height          =   225
            Left            =   4995
            TabIndex        =   16
            Top             =   300
            Width           =   855
         End
         Begin VB.OptionButton optDireito 
            Caption         =   "Direito"
            Height          =   225
            Left            =   7155
            TabIndex        =   17
            Top             =   300
            Width           =   1125
         End
         Begin VB.OptionButton optDesconhecido 
            Caption         =   "Desconhecido"
            Height          =   225
            Left            =   450
            TabIndex        =   14
            Top             =   300
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox txtDistanciaTestada 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2055
         TabIndex        =   10
         Top             =   285
         Width           =   1845
      End
      Begin VB.TextBox txtDistanciaLado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2055
         TabIndex        =   11
         Top             =   645
         Width           =   1845
      End
      Begin VB.Label Label7 
         Caption         =   "Profundidade"
         Height          =   195
         Left            =   4560
         TabIndex        =   24
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Comprimento"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Distância Testada"
         Height          =   195
         Left            =   570
         TabIndex        =   22
         Top             =   345
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Distância Lado"
         Height          =   195
         Left            =   570
         TabIndex        =   21
         Top             =   690
         Width           =   1290
      End
   End
   Begin VB.PictureBox LoozeXP1 
      Height          =   480
      Left            =   135
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   43
      Top             =   8010
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   6990
      Width           =   225
   End
   Begin VB.Label Label2 
      Caption         =   "Ramal conectado a rede:"
      Height          =   300
      Left            =   5445
      TabIndex        =   42
      Top             =   6300
      Width           =   3405
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   6990
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   6
      X1              =   5940
      X2              =   8235
      Y1              =   7110
      Y2              =   7110
   End
   Begin VB.Label lblUsuarioData 
      Height          =   255
      Left            =   180
      TabIndex        =   29
      Top             =   7905
      Width           =   4755
   End
   Begin VB.Label lblResultado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   255
      TabIndex        =   28
      Top             =   1335
      Width           =   7680
   End
   Begin VB.Label lblRede 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6210
      TabIndex        =   41
      Top             =   6780
      Width           =   1770
   End
End
Attribute VB_Name = "FrmCadastroRamalAgua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private object_id_ramal As String
Private tcs As TeCanvas, tdbramais As TeDatabase, tdbtrecho As TeDatabase, object_id_lote As String, Object_id_trecho As String


Private rs As ADODB.Recordset

Dim blnCancelar As Boolean
Dim i As Long
Dim j As Long
Dim blnBotaoFechar As Boolean
Dim intKeyAscii As Integer
        Dim ve As String
         Dim vi As String
         Dim vo As String
         Dim vu As String
         Dim vc As String
          Dim vd As String
          Dim ve As String
          Dim vf As String

Public Sub Init(m_object_id_ramal As String, m_tcs As TeCanvas, m_tdbramais As TeDatabase, m_tdbtrecho As TeDatabase, m_object_id_lote As String, m_object_id_trecho As String)

On Error GoTo Trata_Erro

   If TpConexao = 1 Then 'no caso de SQL consutar ligações está desabilitado
      cmdConsultarLigacoes.Visible = False
   End If

   'DESABILITA SALVAR SE O USUÁRIO É UM VISITANTE
   Set rs = New ADODB.Recordset
   va = """USRLOG"""
         ve = """USRFUN"""
         vi = """SYSTEMUSERS"""
         vo = """TIPO"""
         vu = """ECONOMIAS"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
     If frmCanvas.TipoConexao <> 4 Then
   
   rs.Open "SELECT USRLOG, USRFUN FROM SYSTEMUSERS WHERE USRLOG = '" & strUser & "' ORDER BY USRLOG", Conn, adOpenDynamic, adLockReadOnly
   Else
    rs.Open "SELECT " + va + "," + ve + " FROM " + vi + " WHERE " + va + " = '" & strUser & "' ORDER BY " + va + "", Conn, adOpenDynamic, adLockReadOnly
   
   
   End If
   
   
   If rs.EOF = False Then
      If rs!UsrFun = 3 Or rs!UsrFun = 4 Then 'VISITANTE OU VISUALIZADOR
         
         Me.cmdConfirmar.Enabled = False 'DESABILITA O BOTÃO SALVAR
         
      End If
   End If
   rs.Close


   LoozeXP1.InitIDESubClassing
   object_id_lote = m_object_id_lote
   object_id_ramal = m_object_id_ramal
   Object_id_trecho = m_object_id_trecho
   Set tcs = m_tcs
   Set tdbramais = m_tdbramais
   Set tdbtrecho = m_tdbtrecho
   
   If object_id_ramal <> "" Then ' RAMAL EXISTENTE
         'RETORNA ATRIBUTOS DO RAMAL
         
         Set rs = New ADODB.Recordset
         va = """RAMAIS_AGUA"""
         ve = """OBJECT_ID_"""
         vi = """SYSTEMUSERS"""
         vo = """TIPO"""
         vu = """ECONOMIAS"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
     If frmCanvas.TipoConexao <> 4 Then
         rs.Open ("SELECT * FROM RAMAIS_AGUA WHERE OBJECT_ID_ = '" & object_id_ramal & "'"), Conn, adOpenStatic, adLockReadOnly
         Else
          rs.Open ("SELECT * FROM " + va + " WHERE " + ve + " = '" & object_id_ramal & "'"), Conn, adOpenStatic, adLockReadOnly
        
         End If
         
         If rs.EOF = False Then
            txtDistanciaLado.Text = IIf(IsNull(rs!Distancia_Lado), 0, rs!Distancia_Lado)
            
            txtDistanciaTestada.Text = IIf(IsNull(rs!Distancia_Testada), 0, rs!Distancia_Testada)
            
            txtProfundidade.Text = IIf(IsNull(rs!Profundidade_RAMAL), 0, rs!Profundidade_RAMAL)
            
            txtComprimentoRamal.Text = IIf(IsNull(rs!COMPRIMENTO_RAMAL), 0, rs!COMPRIMENTO_RAMAL)
            
            Select Case rs!posicionamento_lote
               Case 1
                  optDesconhecido = True
               Case 2
                  optEsquerdo = True
               Case 3
                  optCentro = True
               Case 4
                  optDireito = True
            End Select
         
            Me.lblRede.Caption = rs!Object_id_trecho
            
            Me.lblUsuarioData.Caption = "Cadastrado por: " & rs.Fields("USUARIO_LOG").Value & " em " & rs.Fields("DATA_LOG").Value
        
         End If
         rs.Close
         Set rs = Nothing
    
        CarregaLigacoes
   
   Else
      Me.lblRede.Caption = ramal_Object_id_trecho
      optDesconhecido = True
   End If
   
    If Me.lvLigacoes.ListItems.Count > 0 Then
        'Me.cmdConsultarLigacoes.Enabled = True 'DESATIVADO PARA CORRECÇÃO DAS QUERYS DO SQL SERVER
        
    Else
        Me.cmdConsultarLigacoes.Enabled = False
        'CARREGA OS TEXTOS COM O FILTRO PRÉ DETERMINADO
        
        'Set rs = New ADODB.Recordset
        
        
        Dim Vetor As Variant
        Dim str As String
        Dim ArrayTema(10) As String
        Dim i, j As Integer
        
        
        Dim retval As String
        retval = Dir(App.Path & "\Controles\FRamais.txt")
        If retval <> "" Then 'verifica se o arquivo existe na pasta
            Open App.Path & "\Controles\FRamais.txt" For Input As #3
            'Open "C:\ARQUIVOS DE PROGRAMAS\GEOSAN\Controles\FRamais.txt" For Input As #3
            'Do While Not EOF(3) = True
            Do While Not EOF(3)
                 Line Input #3, str
                 Vetor = Split(str, ";")
                 
                 If Vetor(0) = "NUM_LIGAÇÃO" Then
                     Me.txtNumLigacao = Vetor(1)
                     If Vetor(1) <> "" Then
                        Me.txtNumLigacao.Text = Vetor(1)
                        Me.optNumLigacao.Value = True
                        Exit Do
                     End If
                 ElseIf Vetor(0) = "INSCRIÇÃO" Then
                     If Vetor(1) <> "" Then
                        Me.txtInscricao.Text = Vetor(1)
                        Me.optInscricao.Value = True
                        Exit Do
                     End If
                 ElseIf Vetor(0) = "ENDEREÇO" Then
                     If Vetor(1) <> "" Then
                        Me.txtEndereco.Text = Vetor(1)
                        Me.optEndereço.Value = True
                        Exit Do
                     End If
                 ElseIf Vetor(0) = "CONSUMIDOR" Then
                     If Vetor(1) <> "" Then
                        Me.txtConsumidor.Text = Vetor(1)
                        Me.optConsumidor.Value = True
                        Exit Do
                     End If
                 End If
            Loop
            Close #3
            Carrega_PreFiltro
        End If
        
   End If
   

   
   Me.Show vbModal
   LoozeXP1.EndWinXPCSubClassing
   Exit Sub

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = -2147467259 Then
   MsgBox "Não há conexão ativa com o banco de dados. Contate o Administrador de Rede." & Chr(13) & Chr(13) & "O Geosan será fechado.", vbCritical, "Falha de rede"
   Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCanvas - Private Sub TCanvas_onMouseDown - Não há conexão ativa com a rede. Programa foi fechado."
   Close #1
   End
Else
   Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Public Sub Init - " & Err.Number & " - " & Err.Description
   Close #1
   MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
End If

End Sub



Private Sub cmdAtivaFiltro_Click()
    
    If cmdAtivaFiltro.Caption = ">>." Then
        blnCancelar = True                      'VARIÁVEL QUE INTERROMPE A PESQUISA
        cmdAtivaFiltro.Caption = ">>"
    Else
        Me.MousePointer = vbHourglass
        Me.lblResultado.Caption = "Pesquisando..."
        DoEvents
        cmdAtivaFiltro.Caption = ">>."
        
        blnCancelar = False
        
        LimpaLista
        Carrega_PreFiltro
        cmdAtivaFiltro.Caption = ">>" 'REATIVA BOTÃO PARA NOVA PESQUISA
    End If
    
    Me.MousePointer = vbDefault
    
End Sub
Private Function LimpaLista()

    Dim blnExisteSelecionado As Boolean
    blnExisteSelecionado = False
reinicia:
    For i = 1 To lvLigacoes.ListItems.Count
        If lvLigacoes.ListItems(i).Checked = False Then
            lvLigacoes.ListItems.Remove (i)
            lvLigacoes.Refresh
            GoTo reinicia
        End If
    Next
    

End Function


Private Function Carrega_PreFiltro()

On Error GoTo Trata_Erro
   
Dim str As String
Dim itmx As ListItem
   
Dim strIni As String
Dim criterio As String

'HIDROMETRADAS E FICTÍCIAS
  
Dim RS_NRO_LIGACAO As New ADODB.Recordset
Dim strLigacoes As String

   Set rs = New ADODB.Recordset
     
      va = """NRO_LIGACAO"""
         ve = """CLASSIFICACAO_FISCAL"""
         vi = """ENDERECO"""
         vo = """CONSUMIDOR"""
         vu = """COD_LOGRADOURO"""
         vc = """TIPO"""
         vd = """ECONOMIAS"""
           ve = """HIDROMETRADO"""
         vf = """NXGS_V_LIG_COMERCIAL"""
         Dim aa As String
         Dim bb As String
         aa = """RAMAIS_AGUA_LIGACAO"""
     If frmCanvas.TipoConexao <> 4 Then
   strIni = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, ENDERECO, CONSUMIDOR, COD_LOGRADOURO as CODLOGRAD, TIPO,ECONOMIAS,HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL"
     Else
      strIni = "SELECT " + va + "," + ve + "," + vi + "," + vo + "," + vu + " as CODLOGRAD, " + vc + "," + vd + "," + ve + " FROM " + vf + ""
   
     End If
     
     str = "" 'LIMPA A STRING DE COMANDO
     
      If Me.optNumLigacao.Value = True And Trim(Me.txtNumLigacao) <> "" Then
         
         Me.lvLigacoes.SortKey = 0 'SETA O SORT PARA A PRIMEIRA COLUNA E TIRA O ORDER BY DO SELECT
         
         If TpConexao = 1 Then 'SQL
            
            str = strIni & " WHERE CLASSIFICACAO_FISCAL LIKE '" & Me.txtNumLigacao.Text & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO)"
            
         ElseIf TpConexao = 4 Then
             str = strIni & " WHERE " + ve + " LIKE '" & Me.txtNumLigacao.Text & "%' AND " + va + " NOT IN (SELECT " + va + " FROM " + aa + ")"
            
            Else
            
            str = strIni & " A WHERE CLASSIFICACAO_FISCAL LIKE '" & Me.txtNumLigacao.Text & "%' AND NOT EXISTS (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO B WHERE A.NRO_LIGACAO = B.NRO_LIGACAO)"
         
         End If

      ElseIf Me.optInscricao.Value = True And Trim(Me.txtInscricao) <> "" Then
            
         If TpConexao = 1 Then 'SQL
            
            Me.lvLigacoes.SortKey = 1 'SETA O SORT PARA A SEGUNDA COLUNA E TIRA O ORDER BY DO SELECT
            
            str = strIni & " WHERE " + va + " LIKE '" & Me.txtInscricao.Text & "%' AND " + va + " NOT IN (SELECT NRO_LIGACAO FROM " + aa + ")"
             ElseIf TpConexao = 4 Then
                str = strIni & " WHERE NRO_LIGACAO LIKE '" & Me.txtInscricao.Text & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO)"
           
             
             
             
         Else ' ORACLE
         
            str = strIni & " A WHERE NRO_LIGACAO LIKE '" & Me.txtInscricao.Text & "%' AND NOT EXISTS (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO B WHERE A.NRO_LIGACAO = B.NRO_LIGACAO)"
            
         End If
      
     ElseIf Me.optEndereço.Value = True And Trim(Me.txtEndereco) <> "" Then
        
         Me.lvLigacoes.SortKey = 2 'SETA O SORT PARA A TERCEIRA COLUNA E TIRA O ORDER BY DO SELECT
        
         If TpConexao = 1 Then 'SQL
            
            str = strIni & " WHERE upper(ENDERECO) LIKE '%" & UCase(Me.txtEndereco.Text) & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO)" ' ORDER BY TAM ASC, ENDERECO ASC"
         ElseIf TpConexao = 4 Then
         bb = upper(ENDERECO)
          str = strIni & " WHERE ""+bb+"" LIKE '%" & UCase(Me.txtEndereco.Text) & "%' AND " + va + " NOT IN (SELECT " + va + " FROM " + aa + ")" ' ORDER BY TAM ASC, ENDERECO ASC"
      
         
         Else ' ORACLE
            
            'NA CONSULTA ORACLE COM LIKE É NECESSÁRIO COLOCAR O SINAL % NO LUGAR DO ESPAÇO ENTRE PALAVRAS
            criterio = "%" & Replace(Trim(Me.txtEndereco.Text), " ", "%") & "%"
            str = strIni & " A WHERE upper(ENDERECO) LIKE '" & criterio & "' AND NOT EXISTS (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO B WHERE A.NRO_LIGACAO = B.NRO_LIGACAO)"
         
         End If
         
     ElseIf Me.optConsumidor.Value = True And Trim(Me.txtConsumidor.Text) <> "" Then
        
         Me.lvLigacoes.SortKey = 3 'SETA O SORT PARA A QUARTA COLUNA E TIRA O ORDER BY DO SELECT
        
         If TpConexao = 1 Then 'SQL
            
            str = strIni & " WHERE upper(CONSUMIDOR) LIKE '" & criterio & "' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO)"
          ElseIf TpConexao = 4 Then
          bb = upper(CONSUMIDOR)
             str = strIni & " WHERE ""+bb+"" LIKE '" & criterio & "' AND " + va + " NOT IN (SELECT " + va + " FROM " + aa + ")"
       
          
         Else ' ORACLE
         
            'NA CONSULTA ORACLE COM LIKE É NECESSÁRIO COLOCAR O SINAL % NO LUGAR DO ESPAÇO ENTRE PALAVRAS
            criterio = "%" & Replace(Trim(Me.txtConsumidor.Text), " ", "%") & "%"
            str = strIni & " A WHERE upper(CONSUMIDOR) LIKE '" & criterio & "' AND NOT EXISTS (SELECT NRO_LIGACAO FROM RAMAIS_AGUA_LIGACAO B WHERE A.NRO_LIGACAO = B.NRO_LIGACAO)"

         End If

     End If
    
        Close #3
        Open App.Path & "\Controles\FRamais.txt" For Output As #3
        Print #3, str
        Close
        
        Open App.Path & "\Controles\FRamais.txt" For Output As #3
        If Me.optNumLigacao.Value = True Then
            Print #3, "NUM_LIGAÇÃO;" & Me.txtNumLigacao.Text
        ElseIf Me.optInscricao.Value = True Then
            Print #3, "INSCRIÇÃO;" & Me.txtInscricao.Text
        ElseIf Me.optConsumidor.Value = True Then
            Print #3, "CONSUMIDOR;" & Me.txtConsumidor.Text
        ElseIf Me.optEndereço.Value = True Then
            Print #3, "ENDEREÇO;" & Me.txtEndereco.Text
        End If
        
    Close #3
     
    'FAZ SELECT COM BASE NOS CAMPOS CRIADOS
    i = 0
    Me.lblResultado.Caption = "Localizadas " & i & " referencias"
    
    If str <> "" Then
        Set rs = New ADODB.Recordset

        rs.Open str, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        'Set RS = ConnSec.execute(str)
        If rs.EOF = False Then
            'CARREGA NO FORM TODAS AS LIGAÇÕES DISPONIVEIS COM BASE NO PRÉ FILTRO
            Do While Not rs.EOF And blnCancelar = False
                 DoEvents
                 'COLOCAR NO GRID APENAS OS QUE NÃO FORAM CADASTRADOS AINDA
                 
                'Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
                Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("CLASSIFICACAO_FISCAL").Value)
               
                'itmx.SubItems(1) = IIf(IsNull(rs.Fields("CLASSIFICACAO_FISCAL").value), "", rs.Fields("CLASSIFICACAO_FISCAL").value)
                itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").Value), "", rs.Fields("NRO_LIGACAO").Value)
                
                itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").Value), "", rs.Fields("ENDERECO").Value)
                itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").Value), "", rs.Fields("CONSUMIDOR").Value)
                
                'incluído para mostrar o tipo da ligação
                itmx.SubItems(4) = IIf(IsNull(rs.Fields("TIPO").Value), "", rs.Fields("TIPO").Value)
                
                itmx.Tag = rs.Fields("codlograd").Value
                i = i + 1
                Me.lblResultado.Caption = "Mostrando " & i & " de " & i & " referencias encontradas"
                If i >= 500 Then
                    j = i
                    Do While Not rs.EOF And blnCancelar = False
                        rs.MoveNext
                        j = j + 1
                    Loop
                    Me.lblResultado.Caption = "Mostrando " & i & " referencias de " & j & " encontradas"
                    Exit Do
                End If
                
                rs.MoveNext
            Loop
        End If
saida:
        rs.Close
        Set rs = Nothing
    End If
    
    cmdFechar.Caption = "Cancelar"

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = -2147467259 Then
   MsgBox "Não há conexão ativa com o banco de dados. Contate o Administrador de Rede." & Chr(13) & Chr(13) & "O Geosan será fechado.", vbCritical, "Falha de rede"
   Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCanvas - Private Sub TCanvas_onMouseDown - Não há conexão ativa com a rede. Programa foi fechado."
   Close #1
   End
Else
   Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Public Sub Init - " & Err.Number & " - " & Err.Description
   Close #1
   MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
End If

End Function


'''Private Sub cmdConfirmar_Click()
'''
'''   On Error GoTo Trata_Erro
'''   Dim intlocalerro As Integer
'''   Dim rsCria As ADODB.Recordset
'''   Dim a As Integer
'''   Dim cgeo As New clsGeoReference
'''   Dim x As Double
'''   Dim y As Double
'''   Dim str As String
'''
'''   Dim strNroL As String 'NÚMERO DA LIGACAO
'''   Dim strInsc As String 'NUMERO DA INSCRIÇÃO
'''   Dim strTipo As String 'TIPO DA LIGACAO
'''   Dim strCons As String 'CONSUMO DA LIGACAO
'''   Dim strEcon As String 'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
'''   Dim strHidr As String
'''
'''   Set rsCria = New ADODB.Recordset 'recordset utilizado para criar o regitro na tabela
'''   Dim strNroLigaSel As String
'''   'Conn.BeginTrans
'''
'''   Conn.Close
'''   Conn.Open
'''
'''   If object_id_ramal = "" Then 'NOVO RAMAL
'''
'''        str = strUser & Now
'''
'''        'O CAMPO ID DA TABELA É POPULADA COM A AUTO NUMERAÇÃO DA TABELA
'''
'''        'Set rsCria = Conn.execute("SELECT ID FROM RAMAIS_AGUA WHERE OBJECT_ID_ = '" & str & "'")
'''
'''        rsCria.Open "RAMAIS_AGUA", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
'''
'''        'CRIA O RAMAL NA TABELA DE ATRIBUTOS
'''        rsCria.AddNew
'''        rsCria.Fields("OBJECT_ID_").value = str
'''        rsCria.Fields("OBJECT_ID_TRECHO").value = "0"
'''        rsCria.Update
'''
'''        object_id_ramal = rsCria.Fields("ID").value
'''        rsCria.Close
'''
'''        tcs.object_id = object_id_ramal
'''        tcs.saveOnMemory
'''        tcs.SaveInDatabase
'''
'''        'LOCALIZA A COORDENADA DO FINAL DA LINHA DO RAMAL
'''        tdbramais.getPointOfLine 0, object_id_ramal, 1, x, y
'''
'''        'INSERE PONTO NO FINAL DA LINHA DO RAMAL
'''        tdbramais.addPoint object_id_ramal, x, y
'''
'''        tdbramais.getPointOfLine 0, object_id_ramal, 0, x, y
'''        'tdb.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
'''
'''        Object_id_trecho = ramal_Object_id_trecho 'VARIÁVEL RAMAL_OBJECT_ID_TRECHO CARREGADA NO TCANVAS ON_CLICK
'''
'''   Else
'''
'''      'O RAMAL JA EXISTE, RESETA AS INFORMAÇÕES DE CONSUMIDORES ASSOCIADOS A ELE
'''      Conn.execute "DELETE FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ = '" & object_id_ramal & "'"
'''
'''   End If
'''
'''   Dim rsRamal As ADODB.Recordset
'''   Set rsRamal = New ADODB.Recordset
'''
'''   rsRamal.Open "SELECT * FROM RAMAIS_AGUA WHERE ID = " & "'" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic
'''
'''   If rsRamal.EOF = False Then
'''
'''       'ATUALIZA A TABELA DE RAMAIS COMPLEMENTANDO AS INFORMAÇÕES
'''
'''       rsRamal.Fields("Distancia_Lado").value = IIf(IsNumeric(txtDistanciaLado), txtDistanciaLado, 0)
'''       rsRamal.Fields("Distancia_Testada").value = IIf(IsNumeric(txtDistanciaTestada), txtDistanciaTestada, 0)
'''       rsRamal.Fields("Profundidade_RAMAL").value = IIf(IsNumeric(txtProfundidade), txtProfundidade, 0)
'''       rsRamal.Fields("Comprimento_Ramal").value = IIf(IsNumeric(txtComprimentoRamal), txtComprimentoRamal, 0)
'''
'''       For i = 1 To lvLigacoes.ListItems.Count
'''          If lvLigacoes.ListItems(i).Checked = True Then
'''              rsRamal!cod_lograd = lvLigacoes.ListItems(1).Tag 'PEGA O PRIMEIRO LOGRADOURO SELECIONADO NA LISTA
'''              Exit For
'''          End If
'''       Next
'''
'''      If optDesconhecido Then rsRamal.Fields("posicionamento_lote").value = 1
'''      If optEsquerdo Then rsRamal.Fields("posicionamento_lote").value = 2
'''      If optCentro Then rsRamal.Fields("posicionamento_lote").value = 3
'''      If optDireito Then rsRamal.Fields("posicionamento_lote").value = 4
'''
'''      rsRamal!Object_id_ = object_id_ramal
'''      rsRamal!Object_id_trecho = Object_id_trecho
'''      rsRamal!usuario_log = strUser
'''      rsRamal!DATA_LOG = Format(Now, "DD/MM/YY HH:MM")
'''
'''      rsRamal.Update
'''
'''   End If
'''   rsRamal.Close
'''
'''
'''   'RE-ASSOCIA AS LIGAÇÕES AO RAMAL
'''   strNroLigaSel = ""
'''
'''   For a = 1 To lvLigacoes.ListItems.Count
'''       If lvLigacoes.ListItems(a).Checked Then 'PARA CADA ITEM SELECIONADO NA LISTA
'''          If strNroLigaSel <> "" Then
'''             strNroLigaSel = strNroLigaSel & ",'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
'''          Else
'''             strNroLigaSel = "'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
'''          End If
'''       End If
'''   Next
'''
'''
'''   If strNroLigaSel <> "" Then
'''
'''      'SELECIONA AS INFORMAÇÕES DOS CONSUMIDORES NA NXGS_V_LIG_COMERCIAL
'''
'''      str = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, COD_LOGRADOURO, "
'''      str = str & "TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL WHERE NRO_LIGACAO IN (" & strNroLigaSel & ")"
'''
'''      Set rs = New ADODB.Recordset
'''      rs.Open str, Conn, adOpenDynamic, adLockReadOnly, adCmdText
'''
'''       If rs.EOF = False Then
'''         Do While Not rs.EOF
'''
'''            strNroL = Trim(rs!NRO_LIGACAO)                 'NÚMERO DA LIGACAO
'''
'''            If Len(rs!CLASSIFICACAO_FISCAL) > 0 Then
'''               strInsc = Trim(rs!CLASSIFICACAO_FISCAL)
'''            Else
'''               strInsc = ""  'NUMERO DA INSCRIÇÃO
'''            End If
'''
'''            If Len(rs!tipo) > 0 Then
'''               strTipo = Trim(mid(rs!tipo, 1, 20))
'''            Else
'''               strTipo = ""  'TIPO DA LIGACAO
'''            End If
'''
'''            If Len(rs!ECONOMIAS) > 0 Then
'''               strEcon = Trim(rs!ECONOMIAS)
'''            Else
'''               strEcon = 1  'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
'''            End If
'''
'''            If UCase(rs!HIDROMETRADO) = "SIM" Then 'IDENTIFICA E ARMAZENA EM LETRA MINÚSCULA
'''               strHidr = "sim"
'''            ElseIf UCase(rs!HIDROMETRADO) = "NAO" Or UCase(rs!HIDROMETRADO) = "NÃO" Then
'''               strHidr = "nao"
'''            Else
'''               strHidr = ""
'''            End If
'''
'''            str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
'''            str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
'''
'''            Conn.execute (str)
'''            rs.MoveNext
'''
'''         Loop
'''      End If
'''      rs.Close
'''
'''   End If
'''
'''   'INSERINDO RAMAL FICTÍCIO SE ESTE FOI SELECIONADO
'''   If CInt(Me.txtQtd.Text) > 0 Then
'''
'''      'CAPTURA O CONSUMO DIGITADO E CONVERTE SE NECESSÁRIO
'''      If CDbl(Me.txtConsumoFicticia.Text) > 0 Then
'''         If Me.optMetroCubico.value = True Then
'''             'SE FOR METRO CUBICO, CONVERTE PARA LITROS POR SEGUNDO
'''             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",") * 0.00038580246
'''         Else
'''             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",")
'''         End If
'''      Else
'''         strCons = 0
'''      End If
'''
'''      strCons = Replace(strCons, ",", ".") 'troca a vírgula pelo ponto
'''
'''      For i = 0 To CInt(Me.txtQtd.Text) - 1
'''
'''         str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
'''         str = str & "VALUES ('" & object_id_ramal & "','999" & object_id_ramal & i & "','999" & object_id_ramal & i & "','FICTÍCIA','nao','1','" & strCons & "')"
'''
'''         Conn.execute (str)
'''
'''      Next
'''
'''   End If
'''
'''
'''   tcs.plotView
'''   Unload Me
'''
'''
'''Trata_Erro:
'''    If Err.Number = 0 Or Err.Number = 20 Then
'''        Resume Next
'''    ElseIf Err.Number = -2147418113 Then ' Erro geral de rede
'''        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. Reinicie o sistema.", vbInformation
'''        Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description & " Erro Geral de Rede - Programa foi fechado."
'''        Close #1
'''        End
'''
'''    ElseIf Err.Number = -2147417848 Then ' automation error
'''        'Conn.RollbackTrans
'''        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. Reinicie o sistema.", vbInformation
'''        Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description & " - Programa foi fechado."
'''        Close #1
'''        End
'''
'''    ElseIf Err.Number = -2147467259 Or mid(Err.Description, 1, 9) = "ORA-03114" Then 'PERDA DE CONEXÃO BANCO SQL OU ORACLE
'''       'Conn.RollbackTrans
'''       MsgBox "Não há conexão ativa com o banco de dados. Contate o Administrador de Rede." & Chr(13) & Chr(13) & "O Geosan será fechado.", vbCritical, "Falha de rede"
'''       Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''       Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Não há conexão ativa com a rede. Programa foi fechado."
'''       Close #1
'''       End
'''    ElseIf Err.Number = -2147168227 Then ' MAX TRANSACTIONS EXCEDIDA. FECHAR E REABRIR A CONEXÃO
'''        'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'''        Conn.Close
'''        Conn.Open
'''        Resume
'''    Else
''''        If RS Is Nothing Then
''''           If RS.State = 1 Then
''''              RS.Close
''''           End If
''''        End If
'''        tcs.Normal
'''        tcs.Select
'''        'Conn.RollbackTrans
'''
'''        Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description
'''        Close #1
'''
'''        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'''        Unload Me
'''
'''    End If
'''End Sub

Private Sub cmdConfirmar_Click()

   On Error GoTo Trata_Erro
   Dim intlocalerro As Integer
   Dim rsCria As ADODB.Recordset
   Dim a As Integer
   Dim cgeo As New clsGeoReference
   Dim x As Double
   Dim y As Double
   Dim str As String
   
   Dim strNroL As String 'NÚMERO DA LIGACAO
   Dim strInsc As String 'NUMERO DA INSCRIÇÃO
   Dim strTipo As String 'TIPO DA LIGACAO
   Dim strCons As String 'CONSUMO DA LIGACAO
   Dim strEcon As String 'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
   Dim strHidr As String
   
   Set rsCria = New ADODB.Recordset 'recordset utilizado para criar o regitro na tabela
   Dim strNroLigaSel As String
   'Conn.BeginTrans
   
   Conn.Close
   Conn.Open
   
   If object_id_ramal = "" Then 'NOVO RAMAL
      
        str = strUser & Now
        
        'O CAMPO ID DA TABELA É POPULADA COM A AUTO NUMERAÇÃO DA TABELA
        
        'Set rsCria = Conn.execute("SELECT ID FROM RAMAIS_AGUA WHERE OBJECT_ID_ = '" & str & "'")
        
        rsCria.Open "RAMAIS_AGUA", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
        rsCria.AddNew
        rsCria.Fields("object_id_").Value = str
        rsCria.Fields("object_id_trecho").Value = "0"
        rsCria.Update
        
        object_id_ramal = rsCria.Fields("ID").Value
        rsCria.Close
        
        tcs.object_id = object_id_ramal
        tcs.saveOnMemory
        tcs.SaveInDatabase
        
        
        'INSERE PONTO NO FINAL DA LINHA

        tdbramais.getPointOfLine 0, object_id_ramal, 1, x, y
        tdbramais.addPoint object_id_ramal, x, y
        
        tdbramais.getPointOfLine 0, object_id_ramal, 0, x, y
        'tdb.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
        
        Object_id_trecho = ramal_Object_id_trecho 'VARIÁVEL RAMAL_OBJECT_ID_TRECHO CARREGADA NO TCANVAS ON_CLICK
        
        Dim rsRamal As ADODB.Recordset
        Set rsRamal = New ADODB.Recordset
        va = """RAMAIS_AGUA"""
         ve = """ID"""
         vi = """COD_LOGRADOURO"""
         vo = """TIPO"""
         vu = """ECONOMIAS"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
        intlocalerro = 4
          If frmCanvas.TipoConexao <> 4 Then
        rsRamal.Open "SELECT * FROM RAMAIS_AGUA WHERE ID = " & "'" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic, adCmdText
        Else
         rsRamal.Open "SELECT * FROM " + va + " WHERE " + ve + " = " & "'" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic, adCmdText
     
        
        End If
        
        If rsRamal.EOF = False Then
            
            rsRamal.Fields("Distancia_Lado").Value = IIf(IsNumeric(txtDistanciaLado), txtDistanciaLado, 0)
            rsRamal.Fields("Distancia_Testada").Value = IIf(IsNumeric(txtDistanciaTestada), txtDistanciaTestada, 0)
            rsRamal.Fields("Profundidade_RAMAL").Value = IIf(IsNumeric(txtProfundidade), txtProfundidade, 0)
            rsRamal.Fields("Comprimento_Ramal").Value = IIf(IsNumeric(txtComprimentoRamal), txtComprimentoRamal, 0)
             
            For i = 1 To lvLigacoes.ListItems.Count
               If lvLigacoes.ListItems(i).Checked = True Then
                   rsRamal!cod_lograd = lvLigacoes.ListItems(1).Tag 'PEGA O PRIMEIRO LOGRADOURO SELECIONADO NA LISTA
                   Exit For
               End If
            Next
        
           If optDesconhecido Then rsRamal.Fields("posicionamento_lote").Value = 1
           If optEsquerdo Then rsRamal.Fields("posicionamento_lote").Value = 2
           If optCentro Then rsRamal.Fields("posicionamento_lote").Value = 3
           If optDireito Then rsRamal.Fields("posicionamento_lote").Value = 4
           
           rsRamal!Object_id_ = object_id_ramal
           rsRamal!Object_id_trecho = Object_id_trecho
           rsRamal!usuario_log = strUser
           rsRamal!DATA_LOG = Format(Now, "DD/MM/YY HH:MM")
           
           rsRamal.Update
        
        Else
            Exit Sub
            'Conn.execute "DELETE FROM RAMAIS_AGUA WHERE object_id_ = '" & str & "'"
        End If
        rsRamal.Close
        
        
      strNroLigaSel = ""
        
      For a = 1 To lvLigacoes.ListItems.Count
          If lvLigacoes.ListItems(a).Checked Then 'PARA CADA ITEM SELECIONADO NA LISTA
             If strNroLigaSel <> "" Then
                strNroLigaSel = strNroLigaSel & ",'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
             Else
                strNroLigaSel = "'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
             End If
          End If
      Next
        
      
      If strNroLigaSel <> "" Then
      va = """NRO_LIGACAO"""
         ve = """CLASSIFICACAO_FISCAL"""
         vi = """COD_LOGRADOURO"""
         vo = """TIPO"""
         vu = """ECONOMIAS"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
         ve = """NXGS_V_LIG_COMERCIAL"""
             If frmCanvas.TipoConexao <> 4 Then
         str = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, COD_LOGRADOURO, "
         str = str & "TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL WHERE NRO_LIGACAO IN (" & strNroLigaSel & ")"
         Else
            str = "SELECT " + va + "," + ve + "," + vi + ", "
         str = str & vo + "," + vu + "," + vc + " FROM " + ve + " WHERE " + va + " IN ('" & strNroLigaSel & "')"
        
         
         End If
         
         Set rs = New ADODB.Recordset
         rs.Open str, Conn, adOpenDynamic, adLockReadOnly, adCmdText
          
          If rs.EOF = False Then
            Do While Not rs.EOF
               
               strNroL = Trim(rs!NRO_LIGACAO)                 'NÚMERO DA LIGACAO
               
               If Trim(rs!CLASSIFICACAO_FISCAL) <> "" Then strInsc = Trim(rs!CLASSIFICACAO_FISCAL) Else strInsc = ""  'NUMERO DA INSCRIÇÃO
               If rs!tipo <> "" Then strTipo = Trim(rs!tipo) Else strTipo = ""           'TIPO DA LIGACAO
               If rs!ECONOMIAS <> "" Then strEcon = Trim(rs!ECONOMIAS) Else strEcon = ""  'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
               If UCase(rs!HIDROMETRADO) = "SIM" Or UCase(rs!HIDROMETRADO) = "NAO" Then strHidr = LCase(rs!HIDROMETRADO) Else strHidr = "" 'ARMAZENA EM LETRA MINÚSCULA
 va = """RAMAIS_AGUA_LIGACAO"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
               str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
               Else
               str = "INSERT INTO " + va + " (" + ve + "," + vi + "," + vo + "," + vu + "," + vc + "," + vf + "," + ve + ") "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
              
               
               End If
               
               Conn.execute (str)
               rs.MoveNext

            Loop
         End If
         rs.Close
        
      End If
        
      SubInsereFicticios
      
        
      
    Else ' RAMAL EXISTENTE
      
      Set rs = New ADODB.Recordset
      
      va = """RAMAIS_AGUA"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
      
      rs.Open "SELECT * FROM RAMAIS_AGUA WHERE OBJECT_ID_ ='" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic
      Else
      rs.Open "SELECT * FROM " + va + " WHERE " + ve + " ='" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic
   
      
      End If
      
      If rs.EOF = False Then
         rs.Fields("Distancia_Lado").Value = IIf(IsNumeric(txtDistanciaLado), txtDistanciaLado, 0)
         rs.Fields("Distancia_Testada").Value = IIf(IsNumeric(txtDistanciaTestada), txtDistanciaTestada, 0)
         rs.Fields("Profundidade_RAMAL").Value = IIf(IsNumeric(txtProfundidade), txtProfundidade, 0)
         rs.Fields("Comprimento_Ramal").Value = IIf(IsNumeric(txtComprimentoRamal), txtComprimentoRamal, 0)
         
         For i = 1 To lvLigacoes.ListItems.Count
             If lvLigacoes.ListItems(i).Checked = True Then
                 If lvLigacoes.ListItems(1).Tag <> "" Then
                     rs.Fields("cod_lograd").Value = lvLigacoes.ListItems(1).Tag 'PEGA O PRIMEIRO LOGRADOURO SELECIONADO NA LISTA
                 End If
                 Exit For
             End If
         Next
         
         If optDesconhecido Then rs.Fields("posicionamento_lote").Value = 1
         If optEsquerdo Then rs.Fields("posicionamento_lote").Value = 2
         If optCentro Then rs.Fields("posicionamento_lote").Value = 3
         If optDireito Then rs.Fields("posicionamento_lote").Value = 4
          
         rs.Fields("USUARIO_LOG").Value = strUser
         rs.Fields("DATA_LOG").Value = Format(Now, "DD/MM/YY HH:MM") ' & "/" & Format(Now, "MM") & "/" & Format(Now, "YY") & " " & Format(Now, "HH") & ":" & Format(Now, "MM")
         
         rs.Update
         rs.Close
      
      End If

      intlocalerro = 6
       va = """RAMAIS_AGUA_LIGACAO"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """OBJECT_ID_"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "DELETE FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ = '" & object_id_ramal & "'"
Else
  Conn.execute "DELETE FROM " + va + " WHERE " + ve + " = '" & object_id_ramal & "'"

End If

      strNroLigaSel = ""
        
      For a = 1 To lvLigacoes.ListItems.Count
          If lvLigacoes.ListItems(a).Checked Then 'PARA CADA ITEM SELECIONADO NA LISTA
             If Mid(lvLigacoes.ListItems(a).SubItems(1), 1, Len(object_id_ramal)) <> object_id_ramal Then
               
               If strNroLigaSel <> "" Then
                  strNroLigaSel = strNroLigaSel & ",'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
               Else
                  strNroLigaSel = "'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
               End If
             
             End If
          End If
      Next
      
      If strNroLigaSel <> "" Then
         va = """NRO_LIGACAO"""
         ve = """CLASSIFICACAO_FISCAL"""
         vi = """COD_LOGRADOURO"""
         vo = """TIPO"""
         vu = """ECONOMIAS"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
         str = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, COD_LOGRADOURO, "
         str = str & "TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL WHERE NRO_LIGACAO IN (" & strNroLigaSel & ")"
Else
 str = "SELECT " + va + "," + ve + "," + vi + ", "
         str = str & vo + "," + vu + "," + vc + " FROM " + vd + " WHERE " + va + " IN ('" & strNroLigaSel & "')"


End If


         rs.Open str, Conn, adOpenDynamic, adLockReadOnly, adCmdText 'RECORDSET OBTEM INFORMAÇÕES PARA O INSERT
         
         If rs.EOF = False Then
            Do While Not rs.EOF
            
               strNroL = Trim(rs!NRO_LIGACAO)                                             'NÚMERO DA LIGACAO
               
               If Trim(rs!CLASSIFICACAO_FISCAL) <> "" Then strInsc = Trim(rs!CLASSIFICACAO_FISCAL) Else strInsc = ""  'NUMERO DA INSCRIÇÃO
               If rs!tipo <> "" Then strTipo = Trim(rs!tipo) Else strTipo = ""           'TIPO DA LIGACAO
               If rs!ECONOMIAS <> "" Then strEcon = Trim(rs!ECONOMIAS) Else strEcon = ""  'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
               If UCase(rs!HIDROMETRADO) = "SIM" Or UCase(rs!HIDROMETRADO) = "NAO" Then strHidr = LCase(rs!HIDROMETRADO) Else strHidr = "" 'ARMAZENA EM LETRA MINÚSCULA
                  va = """RAMAIS_AGUA_LIGACAO"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
               str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
               Else
                   str = "INSERT INTO " + va + " (" + ve + "," + vi + "," + vo + "," + vu + "," + vc + "," + vf + "," + ve + ") "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
             
               
               End If
               
               Conn.execute (str)
               rs.MoveNext
            
            Loop
         End If
         rs.Close
      End If

      SubInsereFicticios

   End If
   
   Set rs = Nothing
   tcs.plotView
   Unload Me
   

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = -2147418113 Then ' Erro geral de rede
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. Reinicie o sistema.", vbInformation
        Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description & " Erro Geral de Rede - Programa foi fechado."
        Close #1
        End
    
    ElseIf Err.Number = -2147417848 Then ' automation error
        'Conn.RollbackTrans
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. Reinicie o sistema.", vbInformation
        Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description & " - Programa foi fechado."
        Close #1
        End
    
    ElseIf Err.Number = -2147467259 Or Mid(Err.Description, 1, 9) = "ORA-03114" Then 'PERDA DE CONEXÃO BANCO SQL OU ORACLE
       'Conn.RollbackTrans
       MsgBox "Não há conexão ativa com o banco de dados. Contate o Administrador de Rede." & Chr(13) & Chr(13) & "O Geosan será fechado.", vbCritical, "Falha de rede"
       Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
       Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Não há conexão ativa com a rede. Programa foi fechado."
       Close #1
       End
    ElseIf Err.Number = -2147168227 Then ' MAX TRANSACTIONS EXCEDIDA. FECHAR E REABRIR A CONEXÃO
        'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
        Conn.Close
        Conn.Open
        Resume
    Else
'        If RS Is Nothing Then
'           If RS.State = 1 Then
'              RS.Close
'           End If
'        End If
        tcs.Normal
        tcs.Select
        'Conn.RollbackTrans
        
        Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description
        Close #1
        
        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
        Unload Me
        
    End If
End Sub

Private Sub SubInsereFicticios()
   
   Dim str As String
   Dim strCons As String 'CONSUMO DA LIGACAO
      
   'INSERINDO RAMAL FICTÍCIO SE ESTE FOI SELECIONADO
   If CInt(Me.txtQtd.Text) > 0 Then

      'CAPTURA O CONSUMO DIGITADO E CONVERTE SE NECESSÁRIO
      If CDbl(Me.txtConsumoFicticia.Text) > 0 Then
         If Me.optMetroCubico.Value = True Then
             'SE FOR METRO CUBICO, CONVERTE PARA LITROS POR SEGUNDO
             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",") * 0.00038580246
         Else
             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",")
         End If
      Else
         strCons = 0
      End If
      
      strCons = Replace(strCons, ",", ".") 'troca a vírgula pelo ponto
      
      For i = 0 To CInt(Me.txtQtd.Text) - 1
         va = """RAMAIS_AGUA_LIGACAO"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
         str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
         str = str & "VALUES ('" & object_id_ramal & "','999" & object_id_ramal & i & "','999" & object_id_ramal & i & "','FICTÍCIA','nao','1','" & strCons & "')"
            Else
             str = "INSERT INTO " + va + " (" + ve + "," + vi + "," + vo + "," + vu + "," + vc + "," + vf + "," + ve + ") "
         str = str & "VALUES ('" & object_id_ramal & "','999" & object_id_ramal & i & "','999" & object_id_ramal & i & "','FICTÍCIA','nao','1','" & strCons & "')"
         
            
            End If
         Conn.execute (str)
         
      Next
      
   End If

End Sub

Private Sub CarregaLigacoes()
Dim intlocalerro As Integer
   On Error GoTo Trata_Erro
   Dim NRO_LIGACOES As String, INSCRICOES_LOTES As String, msg As String
   Dim rsAssociados As ADODB.Recordset, str As String, itmx As ListItem, a As Integer, Qtde As Integer
   'RECUPERA TODAS AS INSCRICOES DE TODOS LOTE
   str = GetQueryProcess(3)
   INSCRICOES_LOTES = "''"
   If Trim(object_id_lote) = "" Then
      str = Replace(str, "@OBJECT_ID_", "''")
   Else
      str = Replace(str, "@OBJECT_ID_", object_id_lote)
   End If
   intlocalerro = 1
   Set rs = Conn.execute(str)
   While Not rs.EOF
      If INSCRICOES_LOTES = "''" Then
         INSCRICOES_LOTES = "'" & rs(0).Value & "'"
      Else
         INSCRICOES_LOTES = INSCRICOES_LOTES & ",'" & rs(0).Value & "'"
      End If
      rs.MoveNext
   Wend
   rs.Close
      
   intlocalerro = 2
   'RECUPERA TODOS AS LIGAÇÕES JÁ ASSOCIADAS
   
   Set rsAssociados = New ADODB.Recordset
   
   va = """RAMAIS_AGUA_LIGACAO"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
   
   str = "SELECT * FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ = '" & object_id_ramal & "'"
   Else
      str = "SELECT * FROM " + va + " WHERE " + ve + " = '" & object_id_ramal & "'"
   
   End If
   
   rsAssociados.Open str, Conn, adOpenForwardOnly, adLockReadOnly
   
   NRO_LIGACOES = "''"
   
   If rsAssociados.EOF = False Then
      
      While Not rsAssociados.EOF
         If NRO_LIGACOES = "''" Then
            NRO_LIGACOES = "'" & rsAssociados.Fields("NRO_LIGACAO").Value & "'"
         Else
            NRO_LIGACOES = NRO_LIGACOES & ",'" & rsAssociados.Fields("NRO_LIGACAO").Value & "'"
         End If
         rsAssociados.MoveNext
      Wend
   
      intlocalerro = 3
      str = GetQueryProcess(2)
      str = Replace(str, "@NRO_LIGACAO", NRO_LIGACOES)
      str = Replace(str, "@CLASSIFICACAO_FISCAL", INSCRICOES_LOTES)
    
    
    'CARREGA NO FORM TODAS AS LIGAÇÕES CADASTRADAS
    
        Set rs = ConnSec.execute(str)
    
        While Not rs.EOF
           With lvLigacoes
              
              'Set itmx = .ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
              'itmx.SubItems(1) = IIf(IsNull(rs.Fields("CLASSIFICACAO_FISCAL").value), "", rs.Fields("CLASSIFICACAO_FISCAL").value)
              Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("CLASSIFICACAO_FISCAL").Value)
              itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").Value), "", rs.Fields("NRO_LIGACAO").Value)
              
              itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").Value), "", rs.Fields("ENDERECO").Value)
              itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").Value), "", rs.Fields("CONSUMIDOR").Value)
              
              itmx.SubItems(4) = IIf(IsNull(rs.Fields("TIPO").Value), "", rs.Fields("TIPO").Value)
              
              rsAssociados.Filter = "NRO_LIGACAO='" & rs.Fields("NRO_LIGACAO").Value & "'"
              If Not rsAssociados.EOF Then itmx.Checked = True
              itmx.Tag = IIf(IsNull(rs.Fields("codlograd").Value), "", rs.Fields("codlograd").Value)
           End With
           rs.MoveNext
        Wend
        rs.Close
    
   End If
    
    'CARREGA AS LIGAÇÕES FICTÍCIAS
    va = """RAMAIS_AGUA_LIGACAO"""
         ve = """OBJECT_ID_"""
         vi = """NRO_LIGACAO"""
         vo = """INSCRICAO_LOTE"""
         vu = """TIPO"""
         vc = """HIDROMETRADO"""
         vd = """NXGS_V_LIG_COMERCIAL"""
         ve = """CONSUMO_LPS"""
         vf = """ECONOMIAS"""
         
             If frmCanvas.TipoConexao <> 4 Then
    str = "SELECT * FROM RAMAIS_AGUA_LIGACAO WHERE NRO_LIGACAO IN (" & NRO_LIGACOES & ") AND TIPO = 'FICTÍCIA'"
    Else
      str = "SELECT * FROM " + va + " WHERE " + vi + " IN ('" & NRO_LIGACOES & "') AND " + vu + " = 'FICTÍCIA'"
  
    End If
    
    
    
    Set rs = Conn.execute(str)
    If rs.EOF = False Then
    
      While Not rs.EOF
                    
        Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("INSCRICAO_LOTE").Value)
        itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").Value), "", rs.Fields("NRO_LIGACAO").Value)
        
        itmx.SubItems(2) = "" 'IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
        
        itmx.SubItems(3) = "" 'IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)
        
        'rsAssociados.Filter = "NRO_LIGACAO='" & rs.Fields("NRO_LIGACAO").value & "'"
        
        'If Not rsAssociados.EOF Then itmx.Checked = True
        
        itmx.Checked = True
        itmx.SubItems("4") = "FICTÍCIA"
        Me.txtQtd.Text = CInt(Me.txtQtd.Text) + 1
        Me.optLitrosSegundo.Value = True
        Me.txtConsumoFicticia.Text = IIf(IsNull(rs.Fields("CONSUMO_LPS").Value), "0.00", rs.Fields("CONSUMO_LPS").Value)
        rs.MoveNext
      Wend
      rs.Close
    End If
   intlocalerro = 4
   
   rsAssociados.Close


'CarregaLigacoes_err:
'   msg = "object_id_lote: '" & object_id_lote & "'"
'   msg = msg & vbCrLf & "INSCRICOES_LOTES: " & INSCRICOES_LOTES
'   MsgBox Err.Description & vbCrLf & msg & vbCrLf & str

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & "  - frmCadastroRamalAgua - Private Sub carregaLigacoes() - Local " & intlocalerro & " - " & Err.Number & " - " & Err.Description
   Close #1
   MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
   Err.Clear
End If


End Sub

Private Sub cmdConsultarLigacoes_Click()
'   lvLigacoes.ListItems.Clear
'   While Not rs.EOF
'      Set i = lvLigacoes.ListItems.Add(, , rs(0).value)
'      i.SubItems(1) = IIf(IsNull(rs(1).value), "", rs(1).value)
'      i.SubItems(2) = ""
'      i.SubItems(3) = IIf(IsNull(rs(2).value), "", rs(2).value)
'      rs.MoveNext
'   Wend
    
    
   Dim str As String
   Dim j As Integer
   Dim list As ListItem
   frmConsumoLote.lvLigacoes.ListItems.Clear
   For j = 1 To Me.lvLigacoes.ListItems.Count
            
      'frmConsumoLote.lvLigacoes.ListItems.Add (1)
      
      If Mid(Me.lvLigacoes.ListItems.Item(j), 1, 3) <> "999" Then 'NÃO INSERE FICTÍCIA (COMEÇAM COM 999)
      
         Set list = frmConsumoLote.lvLigacoes.ListItems.Add(, , Me.lvLigacoes.ListItems.Item(j))
            
         list.SubItems(1) = Me.lvLigacoes.ListItems(j).SubItems(1)
         
         list.SubItems(2) = Me.lvLigacoes.ListItems(j).SubItems(2)
         
         list.SubItems(3) = Me.lvLigacoes.ListItems(j).SubItems(3)
      
      End If
   Next
    
   frmConsumoLote.Show (1)
   

   
    
End Sub

Private Sub cmdFechar_Click()
    
    If cmdFechar.Caption = "Fechar" Then
        If object_id_ramal = "" Then
           tcs.Normal
           tcs.Select
        End If
        Unload Me
    Else
        If MsgBox("Deseja cancelar as alterações realizadas?", vbQuestion + vbDefaultButton2 + vbYesNo, "Cancelar") = vbYes Then
             If object_id_ramal = "" Then
                tcs.Normal
                tcs.Select
             End If
             Unload Me
        End If
    End If
   
End Sub

Private Sub cmdPesquisaLigacoes_Click()
   frmCadastroRamalFiltro.Init Me, tcs, object_id_ramal
End Sub

Private Function Verifica_Ligacao(index As Integer) As Boolean
On Error GoTo Trata_Erro
   Dim a As Integer, UltimoEndereco As String, rs As ADODB.Recordset, str As String
   For a = 1 To lvLigacoes.ListItems.Count
      If lvLigacoes.ListItems(a).Checked Then
         If UltimoEndereco <> "" Then
            'VERIFICA SE TODOS OS LOGRADOUROS SAO O MESMO
            If lvLigacoes.ListItems(a).Tag <> UltimoEndereco Then
               MsgBox "Não possível vincular em um mesmo ramal ligações de logradouros diferentes", vbExclamation
               Exit Function
            End If
            
         End If
         UltimoEndereco = lvLigacoes.ListItems(a).Tag
      End If
   Next
   'VEVIFICA SE A LIGAÇÃO JÁ ESTÁ VINCULADA EM OUTRO RAMAL
   str = GetQueryProcess(10)
   str = Replace(str, "@LAYER", tcs.getCurrentLayer)
   str = Replace(str, "@OBJECT_ID_RAMAL", object_id_ramal)
   str = Replace(str, "@NRO_LIGACAO", lvLigacoes.ListItems(index).Text)
   Set rs = Conn.execute(str)
   If Not rs.EOF Then
      MsgBox "Esta ligação embora esteja vinculada este lote, já está vincula a outro ramal:" & rs(0).Value, vbExclamation
      Exit Function
   End If
   rs.Close
   Set rs = Nothing
   Verifica_Ligacao = True

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   Open App.Path & "\Controles\GeoSanLog.txt" For Append As #1
   Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub Verifica_Ligacao " & Err.Number & " - " & Err.Description
   Close #1
   MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
End If

End Function




Private Sub lvLigacoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdFechar.Caption = "Cancelar"
End Sub


Private Sub optCentro_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdFechar.Caption = "Cancelar"
End Sub


Private Sub optDireito_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdFechar.Caption = "Cancelar"
End Sub

Private Sub optEsquerdo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdFechar.Caption = "Cancelar"
End Sub

Private Sub txtComprimentoRamal_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub

Private Sub txtConsumoFicticia_LostFocus()
   If IsNumeric(txtConsumoFicticia.Text) = False Then
      MsgBox "Somente números são aceitos para consumo médio.", , ""
      txtConsumoFicticia.SetFocus
   End If
End Sub

Private Sub txtDistanciaLado_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub

Private Sub txtDistanciaTestada_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub


Private Sub txtInscricao_Change()
    Me.optInscricao.Value = True
End Sub

Private Sub txtEndereco_Change()
    Me.optEndereço.Value = True
End Sub

Private Sub txtConsumidor_Change()
    Me.optConsumidor.Value = True
End Sub

Private Function testa_letra(ByVal KeyAscii As Integer)
'FUNÇÃO QUE VERIFICA SE O CARACTERE DIGITADO É NUMÉRICO OU BACKSPACE OU VIRGULA, CASO CONTRARIO, ANULA
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 44 Then
        intKeyAscii = KeyAscii
    Else
        intKeyAscii = 0
    End If
End Function

Private Sub txtNumLigacao_KeyPress(KeyAscii As Integer)
    Me.optNumLigacao.Value = True
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub

Private Sub txtProfundidade_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub


Private Sub txtQtd_LostFocus() 'LIGAÇÃO FICTÍCIA
   If txtQtd.Text = "" Then
      txtQtd.Text = "0"
   End If
   
   If IsNumeric(txtQtd.Text) = False Then
      txtQtd.SetFocus
   End If
   
End Sub

Private Sub UpDown2_Change() 'LIGAÇÃO FICTÍCIA
   
     txtQtd.Text = UpDown2.Value

End Sub



