VERSION 5.00
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Begin VB.Form frmAtualizarSetores 
   Caption         =   "Carregar Polígono de Setor"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOperacoes 
      Caption         =   "Operações"
      Enabled         =   0   'False
      Height          =   375
      Left            =   195
      TabIndex        =   18
      Top             =   4275
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações da Seleção"
      Height          =   3960
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   4020
      Begin VB.CheckBox Check5 
         Height          =   210
         Left            =   3435
         TabIndex        =   20
         Top             =   1815
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   210
         Left            =   3435
         TabIndex        =   19
         Top             =   2100
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   210
         Left            =   3435
         TabIndex        =   17
         Top             =   3345
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Height          =   210
         Left            =   3435
         TabIndex        =   16
         Top             =   3060
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   210
         Left            =   3435
         TabIndex        =   15
         Top             =   975
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Height          =   210
         Left            =   3435
         TabIndex        =   14
         Top             =   690
         Value           =   1  'Checked
         Width           =   255
      End
      Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
         Left            =   2160
         OleObjectBlob   =   "frmAtualizarSetores.frx":0000
         Top             =   2280
      End
      Begin TECOMDATABASELibCtl.TeDatabase TeDatabasePoligono 
         Left            =   2040
         OleObjectBlob   =   "frmAtualizarSetores.frx":0024
         Top             =   3360
      End
      Begin VB.Label Label12 
         Caption         =   "Dentro do Polígono"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   1830
         Width           =   1680
      End
      Begin VB.Label Label11 
         Caption         =   "Na divisa do Polígono"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   2115
         Width           =   1710
      End
      Begin VB.Label lblQtdNosNoPoligono 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblQtdNosNaDivisa 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   300
         Left            =   2040
         TabIndex        =   22
         Top             =   2085
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Nó de Rede de água"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   1530
         Width           =   3420
      End
      Begin VB.Label Label7 
         Caption         =   "Quantidade      Incluir"
         Height          =   225
         Left            =   2250
         TabIndex        =   13
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label Label6 
         Caption         =   "Redes de água"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   405
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Ramais de água"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   2715
         Width           =   1770
      End
      Begin VB.Label lblQtdRamaisNaDivisa 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   300
         Left            =   2040
         TabIndex        =   10
         Top             =   3345
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "Na divisa do Polígono"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   3345
         Width           =   1710
      End
      Begin VB.Label lblQtdRamaisNoPoligono 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   3045
         Width           =   1005
      End
      Begin VB.Label Label9 
         Caption         =   "Dentro do Polígono"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   3045
         Width           =   1680
      End
      Begin VB.Label lblQtdAguaNaDivisa 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   300
         Left            =   2045
         TabIndex        =   6
         Top             =   990
         Width           =   1000
      End
      Begin VB.Label lblQtdAguaNoPoligono 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   285
         Left            =   2045
         TabIndex        =   5
         Top             =   705
         Width           =   1000
      End
      Begin VB.Label Label5 
         Caption         =   "Na divisa do Polígono"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   990
         Width           =   1710
      End
      Begin VB.Label Label4 
         Caption         =   "Dentro do Polígono"
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   705
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Fechar"
      Height          =   390
      Left            =   2130
      TabIndex        =   1
      Top             =   4275
      Width           =   975
   End
   Begin VB.CommandButton cmdCarregar 
      Caption         =   "Carregar"
      Height          =   390
      Left            =   3180
      TabIndex        =   0
      Top             =   4275
      Width           =   990
   End
End
Attribute VB_Name = "frmAtualizarSetores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Qtde As Long
Dim icont As Integer

Dim idLayerDistrito As Integer
Dim strLayerDistrito As String
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


Private Sub cmdOperacoes_Click()
   
   Me.Visible = False
   
   Dim frlAltPoligono As New frmAlteraNoPoligono
  
   frlAltPoligono.Show 1
   
   
   Unload Me
   
End Sub
'Rotina responsável por carregar os dados do polígono de seleção para a tabela POLIGONO_SELECAO para depois gerar relatórios
'ou exportar para o Epanet
'
'XXX - Aqui é que terá que entrar o código para desconsiderar as redes desativadas
Private Sub cmdCarregar_Click()
    On Error GoTo Trata_Erro
    Dim i As Long
    Dim contador As Integer
    
    contador = 1
    Me.MousePointer = vbHourglass
    'Armazena no arquivo o nome do usuário que gerou este polígono de seleção
    Open App.path & "\Controles\UserLog.txt" For Output As #3
    Print #3, strUser
    Close #3
    
    a = "POLIGONO_SELECAO"
    b = "USUARIO"
    c = "TIPO"

    '********************************************************************************************************************
    'LIMPA DA BASE AS REDES (TIPO = 1) SELECIONADAS ANTERIORMENTE PELO USUÁRIO na rotina de carregar polígono de seleção
    If frmCanvas.TipoConexao <> 4 Then
        'Caso não seja Postgres
        Conn.execute ("DELETE FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 1")
    Else
        'Caso seja Postgres
        Conn.execute ("DELETE FROM " + """" + a + """" + " WHERE " + """" + b + """" + " = '" & strUser & "' AND " + """" + c + """" + " = '1'")
    End If
        
    If Me.Check1.value = 1 Then
        If blnPoligonoVirtual = True Then 'VERIFICA SE A OPERAÇÃO É POR POLÍGONO VIRTUAL OU REAL
            If lngTotalRedesDentro > 0 Then
                FrmMain.ProgressBar1.Max = lngTotalRedesDentro + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To lngTotalRedesDentro - 1
                    DoEvents
                    If ArrRedesDentro(i) <> 0 Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & ArrRedesDentro(i) & "','" & strUser & "',1)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ") VALUES ( '" & ArrRedesDentro(i) & "','" & strUser & "',1)")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        Else
            Qtde = TeDatabasePoligono.locateInsideofPolygon(idPoligonSel, , tpLINES, "WATERLINES")
            If Qtde > 0 Then
                FrmMain.ProgressBar1.Max = Qtde + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To Qtde - 1
                    DoEvents
                    If TeDatabasePoligono.objectIds(i) <> "" Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "',1)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "','1')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        End If
    End If
    If Me.Check2.value = 1 Then 'PEGA AS REDES QUE FORAM TOCADAS PELO POLÍGONO
        If blnPoligonoVirtual = True Then 'VERIFICA SE A OPERAÇÃO É POR POLÍGONO VIRTUAL OU REAL
            If lngTotalRedesDivisa > 0 Then
                FrmMain.ProgressBar1.Max = lngTotalRedesDivisa + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To lngTotalRedesDivisa - 1
                    DoEvents
                    If ArrRedesDivisa(i) <> 0 Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & ArrRedesDivisa(i) & "','" & strUser & "',1)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES  ( '" & ArrRedesDivisa(i) & "','" & strUser & "','1')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        Else
            Qtde = TeDatabasePoligono.locatePolygonthatCrosses(idPoligonSel, , tpLINES, "WATERLINES")
            If Qtde > 0 Then
                FrmMain.ProgressBar1.Max = Qtde + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To Qtde - 1
                    DoEvents
                    If TeDatabasePoligono.objectIds(i) <> "" Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "',1)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "','1')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        End If
    End If
    b = "POLIGONO_SELECAO"
    c = "USUARIO"
    d = "TIPO"
    
    '********************************************************************************************************************
    'LIMPA DA BASE OS RAMAIS (TIPO = 2) SELECIONADAS ANTERIORMENTE PELO USUÁRIO na rotina de carregar polígono de seleção
    If frmCanvas.TipoConexao <> 4 Then
        Conn.execute ("DELETE FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 2")
    Else
        Conn.execute ("DELETE FROM " + """" + b + """" + " WHERE " + """" + c + """" + " = '" & strUser & "' AND " + """" + d + """" + " = '2'")
    End If
    If Me.Check3.value = 1 Then 'PEGA OS RAMAIS QUE ESTÃO DENTRO DO POLÍGONO
        If blnPoligonoVirtual = True Then 'VERIFICA SE A OPERAÇÃO É POR POLÍGONO VIRTUAL OU REAL
            If lngTotalRamaisDentro > 0 Then
                FrmMain.ProgressBar1.Max = lngTotalRamaisDentro + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To lngTotalRamaisDentro - 1
                    DoEvents
                    If ArrRamaisDentro(i) <> 0 Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & ArrRamaisDentro(i) & "','" & strUser & "',2)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES  ( '" & ArrRamaisDentro(i) & "','" & strUser & "','2')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        Else
            'qtde = tedatabasepoligono.locatePolygonthatCrosses(idPoligonSel, , tpLINES, "RAMAIS_AGUA")
            Qtde = TeDatabasePoligono.locateInsideofPolygon(idPoligonSel, , tpLINES, "RAMAIS_AGUA")
            If Qtde > 0 Then
                FrmMain.ProgressBar1.Max = Qtde + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To Qtde - 1
                    DoEvents
                    If TeDatabasePoligono.objectIds(icont) <> "" Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "',2)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES  ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "','2')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        End If
    End If
    If Me.Check4.value = 1 Then 'PEGA OS RAMAIS QUE ESTÃO NA BORDA DO POLÍGONO
        If blnPoligonoVirtual = True Then 'VERIFICA SE A OPERAÇÃO É POR POLÍGONO VIRTUAL OU REAL
            If lngTotalRamaisDivisa > 0 Then
                FrmMain.ProgressBar1.Max = lngTotalRamaisDivisa + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To lngTotalRamaisDivisa - 1
                    DoEvents
                    If ArrRamaisDivisa(i) <> 0 Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & ArrRamaisDivisa(i) & "','" & strUser & "',2)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES  ( '" & ArrRamaisDivisa(i) & "','" & strUser & "','2')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        Else
            'qtde = tedatabasepoligono.locateInsideofPolygon(idPoligonSel, , tpLINES, "RAMAIS_AGUA")
            Qtde = TeDatabasePoligono.locatePolygonthatCrosses(idPoligonSel, , tpLINES, "RAMAIS_AGUA")
            If Qtde > 0 Then
                FrmMain.ProgressBar1.Max = Qtde + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
                For i = 0 To Qtde - 1
                    DoEvents
                    If TeDatabasePoligono.objectIds(i) <> "" Then
                        a = "POLIGONO_SELECAO"
                        b = "OBJECT_ID_"
                        c = "USUARIO"
                        d = "TIPO"
                        If frmCanvas.TipoConexao <> 4 Then
                            Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "',2)")
                        Else
                            Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES  ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "','2')")
                        End If
                    End If
                    FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
                Next
            End If
        End If
    End If


    '********************************************************************************************************************
    'LIMPA DA BASE OS NÓS (TIPO = 0) SELECIONADAS ANTERIORMENTE PELO USUÁRIO na rotina de carregar polígono de seleção
    'implementando a funcionalidade para nós de redes de agua
   

a = "POLIGONO_SELECAO"
b = "USUARIO"
c = "TIPO"

     If frmCanvas.TipoConexao <> 4 Then
   Conn.execute ("DELETE FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = 0")
      Else
      Conn.execute ("DELETE FROM " + """" + a + """" + " WHERE " + """" + b + """" + " = '" & strUser & "' AND " + """" + c + """" + " = '0'")
      End If
        
   If Me.Check1.value = 1 Then
      If blnPoligonoVirtual = True Then 'VERIFICA SE A OPERAÇÃO É POR POLÍGONO VIRTUAL OU REAL
         
         If lngTotalPontosDentro > 0 Then
            FrmMain.ProgressBar1.Max = lngTotalPontosDentro + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
            For i = 0 To lngTotalPontosDentro - 1
               DoEvents
               
               If ArrPontosDentro(i) <> 0 Then
               
                 a = "POLIGONO_SELECAO"
      b = "OBJECT_ID_"
      c = "USUARIO"
      d = "TIPO"
      


     If frmCanvas.TipoConexao <> 4 Then
         
        Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & ArrPontosDentro(i) & "','" & strUser & "',0)")
     Else
     
     Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES   ( '" & ArrPontosDentro(i) & "','" & strUser & "','0')")
     End If
                
               End If
               FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
            Next
         End If
      Else
         
         Qtde = TeDatabasePoligono.locateInsideofPolygon(idPoligonSel, , tpPOINTS, "WATERCOMPONENTS")
         If Qtde > 0 Then
            FrmMain.ProgressBar1.Max = Qtde + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
            For i = 0 To Qtde - 1
               DoEvents
               If TeDatabasePoligono.objectIds(i) <> "" Then
                 a = "POLIGONO_SELECAO"
      b = "OBJECT_ID_"
      c = "USUARIO"
      d = "TIPO"
      


     If frmCanvas.TipoConexao <> 4 Then
         
     Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "',0)")
     Else
     
     Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES   ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "','0')")
     End If
                
               
               
                  
               End If
               FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
            Next
         End If
      End If
   End If
   
   
   If Me.Check2.value = 1 Then 'PEGA OS COMPONENTES DE REDES QUE FORAM TOCADAS PELO POLÍGONO
      
      If blnPoligonoVirtual = True Then 'VERIFICA SE A OPERAÇÃO É POR POLÍGONO VIRTUAL OU REAL
         
         If lngTotalPontosDivisa > 0 Then
            FrmMain.ProgressBar1.Max = lngTotalPontosDivisa + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
            For i = 0 To lngTotalPontosDivisa - 1
               DoEvents
               
               If ArrRedesDivisa(i) <> 0 Then
               
                a = "POLIGONO_SELECAO"
      b = "OBJECT_ID_"
      c = "USUARIO"
      d = "TIPO"
      


     If frmCanvas.TipoConexao <> 4 Then
         
   Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & ArrPontosDivisa(i) & "','" & strUser & "',0)")
     Else
     
     Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES  ( '" & ArrPontosDivisa(i) & "','" & strUser & "','0')")
     End If
               
               
                  
               End If
               FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
            Next
         End If
      Else
      
         Qtde = TeDatabasePoligono.locatePolygonthatCrosses(idPoligonSel, , tpPOINTS, "WATERCOMPONENTS")
         
         If Qtde > 0 Then
            FrmMain.ProgressBar1.Max = Qtde + 1: FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1
            For i = 0 To Qtde - 1
               DoEvents
               If TeDatabasePoligono.objectIds(i) <> "" Then
               
                a = "POLIGONO_SELECAO"
      b = "OBJECT_ID_"
      c = "USUARIO"
      d = "TIPO"
      


     If frmCanvas.TipoConexao <> 4 Then
         
       Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "',0)")
     Else
     
     Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ")VALUES   ( '" & TeDatabasePoligono.objectIds(i) & "','" & strUser & "','0')")
     End If
               
                
               End If
               FrmMain.ProgressBar1.value = FrmMain.ProgressBar1.value + 1
            Next
         End If
      End If
   End If
    
    '********************************************************************************************************************


fim:

   FrmMain.ProgressBar1.Visible = False
   
   Me.cmdOperacoes.Enabled = True
   
   Me.MousePointer = vbDefault
   
If contador <> 10 Then
 MsgBox "Polígono carregado!", vbInformation, ""
  contador = 10
 End If

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   Me.MousePointer = vbDefault
   Err.Clear
   GoTo fim
End If


End Sub


Private Sub cmdCancelar_Click()
   Unload Me
End Sub



Private Sub Form_Load()
On Error GoTo Trata_Erro
   'SE FOI CRIADO UM POLÍGONO NO MAPA E ESTE POLÍGONO SELECIONOU REDES OU RAMAIS,

      
   If blnPoligonoVirtual = True Then
      
      Me.Caption = "Poligono Virtual"
      
      Me.lblQtdAguaNoPoligono.Caption = lngTotalRedesDentro
      Me.lblQtdAguaNaDivisa.Caption = lngTotalRedesDivisa
      
      Me.lblQtdNosNoPoligono.Caption = lngTotalPontosDentro
      Me.lblQtdNosNaDivisa.Caption = lngTotalPontosDivisa
      
      Me.lblQtdRamaisNoPoligono.Caption = lngTotalRamaisDentro
      Me.lblQtdRamaisNaDivisa.Caption = lngTotalRamaisDivisa

      

   Else
      
      If frmCanvas.TipoConexao <> 4 Then
         
    
      
       Qtde = 0
      Me.lblQtdAguaNoPoligono.Caption = Qtde
      Me.lblQtdAguaNaDivisa.Caption = Qtde
      Me.lblQtdRamaisNoPoligono.Caption = Qtde
      Me.lblQtdRamaisNaDivisa.Caption = Qtde
         
        
      Me.TeDatabasePoligono.Provider = typeconnection
      Me.TeDatabasePoligono.connection = Conn
      Me.TeDatabasePoligono.setCurrentLayer strLayerAtivo
      Else
      '  If frmCanvas.POSTC <> 10 Then
      Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim Senha As String
Dim decriptada As String
      Dim usuario As String

mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
nStr = frmCanvas.FunDecripta(Senha)
decriptada = frmCanvas.Senha
usuario = ReadINI("CONEXAO", "USER", App.path & "\CONTROLES\GEOSAN.ini")
 TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
 
 'frmCanvas.POST2C (10)
 
 '  End If
   
  Me.TeDatabasePoligono.Provider = typeconnection
  Me.TeDatabasePoligono.username = usuario
    
      Me.TeDatabasePoligono.connection = TeAcXConnection1.objectConnection_
      Me.TeDatabasePoligono.setCurrentLayer strLayerAtivo

 ''frmCanvas.POST2C (10)
 
   'End If


      Qtde = 0
      Me.lblQtdAguaNoPoligono.Caption = Qtde
      Me.lblQtdAguaNaDivisa.Caption = Qtde
      Me.lblQtdRamaisNoPoligono.Caption = Qtde
      Me.lblQtdRamaisNaDivisa.Caption = Qtde
         
        
     
      End If

      Qtde = Me.TeDatabasePoligono.locateInsideofPolygon(idPoligonSel, , tpLINES, "WATERLINES")
     
      Me.lblQtdAguaNoPoligono.Caption = Qtde
      
      Qtde = Me.TeDatabasePoligono.locatePolygonthatCrosses(idPoligonSel, , tpLINES, "WATERLINES")
      Me.lblQtdAguaNaDivisa.Caption = Qtde
      
      Qtde = Me.TeDatabasePoligono.locateInsideofPolygon(idPoligonSel, , tpLINES, "RAMAIS_AGUA")
      Me.lblQtdRamaisNoPoligono.Caption = Qtde
      
      Qtde = Me.TeDatabasePoligono.locatePolygonthatCrosses(idPoligonSel, , tpLINES, "RAMAIS_AGUA")
      Me.lblQtdRamaisNaDivisa.Caption = Qtde
      

      End If
  

   
Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   Me.MousePointer = vbDefault
   'MsgBox Err.Number & " - " & Err.Description
End If

End Sub


Private Sub TeDatabase2_errorMessage(ByVal code As String, ByVal message As String)

End Sub
