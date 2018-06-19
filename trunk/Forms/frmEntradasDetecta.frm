VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEntradasDetecta 
   Caption         =   "Entradas para o Detecta"
   ClientHeight    =   4950
   ClientLeft      =   5475
   ClientTop       =   5400
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7635
   Begin VB.Frame Frame4 
      Height          =   4185
      Left            =   15
      TabIndex        =   36
      Top             =   4770
      Visible         =   0   'False
      Width           =   7605
      Begin VB.Label Label4 
         Caption         =   "4"
         Height          =   405
         Left            =   240
         TabIndex        =   37
         Top             =   390
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1170
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "C:\detecta.dat"
      Filter          =   ".dat"
      InitDir         =   "C:\"
   End
   Begin VB.Frame Frame3 
      Height          =   4185
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   7605
      Begin VB.CommandButton cmPreencher 
         Caption         =   "Preencher Grade com dados externos"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   5295
      End
      Begin MSFlexGridLib.MSFlexGrid fgVazoesNos 
         Height          =   3675
         Left            =   90
         TabIndex        =   46
         Top             =   480
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   6482
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4185
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   7605
      Begin VB.Frame Frame12 
         Height          =   2595
         Left            =   3120
         TabIndex        =   47
         Top             =   240
         Width           =   1695
         Begin VB.ComboBox cbIntervalosCalc 
            Height          =   315
            ItemData        =   "frmEntradasDetecta.frx":0000
            Left            =   270
            List            =   "frmEntradasDetecta.frx":004C
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1260
            Width           =   1185
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Intervalos de Tempo para Cálculo"
            Height          =   705
            Left            =   270
            TabIndex        =   49
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3825
         Left            =   4920
         TabIndex        =   41
         Top             =   240
         Width           =   2565
         Begin VB.ListBox ListIntervalos 
            Height          =   2310
            ItemData        =   "frmEntradasDetecta.frx":00A7
            Left            =   150
            List            =   "frmEntradasDetecta.frx":00A9
            Style           =   1  'Checkbox
            TabIndex        =   44
            Top             =   870
            Width           =   2325
         End
         Begin VB.TextBox txtBombas 
            Height          =   350
            Left            =   900
            TabIndex        =   42
            Top             =   3300
            Width           =   1000
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Intervalos de Tempo em Que a Bomba Funcionará"
            Height          =   825
            Left            =   570
            TabIndex        =   48
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Bombas"
            Height          =   225
            Left            =   150
            TabIndex        =   43
            Top             =   3360
            Width           =   675
         End
      End
      Begin VB.Frame Frame10 
         Height          =   855
         Left            =   180
         TabIndex        =   38
         Top             =   240
         Width           =   2685
         Begin VB.TextBox txtValvulas 
            Height          =   350
            Left            =   120
            TabIndex        =   39
            Top             =   270
            Width           =   1000
         End
         Begin VB.Label Label2 
            Caption         =   "Válvulas"
            Height          =   225
            Left            =   1290
            TabIndex        =   40
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tubos com Vazamento"
         Height          =   2775
         Left            =   180
         TabIndex        =   34
         Top             =   1200
         Width           =   2685
         Begin VB.ListBox listLeft 
            Height          =   2310
            ItemData        =   "frmEntradasDetecta.frx":00AB
            Left            =   120
            List            =   "frmEntradasDetecta.frx":00AD
            Style           =   1  'Checkbox
            TabIndex        =   35
            Top             =   270
            Width           =   2445
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Intervalos para Cálculo"
         Height          =   225
         Left            =   5640
         TabIndex        =   45
         Top             =   570
         Width           =   1905
      End
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   6420
      TabIndex        =   15
      Top             =   4410
      Width           =   1155
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Próximo >"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5130
      TabIndex        =   13
      Top             =   4410
      Width           =   1155
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Anterior"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4410
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2550
      TabIndex        =   11
      Top             =   4410
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   4185
      Left            =   10
      TabIndex        =   14
      Top             =   0
      Width           =   7605
      Begin VB.CommandButton cmdFile 
         Height          =   345
         Left            =   7110
         TabIndex        =   32
         ToolTipText     =   "Selecione arquivo destino"
         Top             =   3690
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtFile 
         Height          =   315
         Left            =   4140
         TabIndex        =   31
         Top             =   3660
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Frame Frame8 
         Caption         =   "Regime a ser calculado"
         Height          =   1005
         Left            =   210
         TabIndex        =   30
         Top             =   2970
         Width           =   3045
         Begin VB.OptionButton optExtensivo 
            Caption         =   "Permanente e Extensivo"
            Height          =   285
            Left            =   300
            TabIndex        =   5
            Top             =   630
            Width           =   2565
         End
         Begin VB.OptionButton optPermanente 
            Caption         =   "Permanente"
            Height          =   255
            Left            =   300
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Insira as informações sobre o fluido"
         Height          =   1335
         Left            =   3420
         TabIndex        =   25
         Top             =   2220
         Width           =   4005
         Begin VB.TextBox txtMassaEspeficac 
            Height          =   350
            Left            =   270
            TabIndex        =   10
            Top             =   810
            Width           =   1000
         End
         Begin VB.TextBox txtViscosidade 
            Height          =   350
            Left            =   240
            TabIndex        =   9
            Top             =   300
            Width           =   1000
         End
         Begin VB.Label Label10 
            Caption         =   "Massa Específica (kg/m3)"
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Top             =   843
            Width           =   2415
         End
         Begin VB.Label Label9 
            Caption         =   "Viscosidade Cinematica (m2/s)"
            Height          =   285
            Left            =   1470
            TabIndex        =   26
            Top             =   333
            Width           =   2235
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Insira as informações gerais p/ calculo"
         Height          =   1845
         Left            =   3420
         TabIndex        =   23
         Top             =   240
         Width           =   4035
         Begin VB.TextBox txtPressaoMin 
            Height          =   350
            Left            =   210
            TabIndex        =   7
            Top             =   810
            Width           =   1000
         End
         Begin VB.TextBox txtPressaoMax 
            Height          =   350
            Left            =   210
            TabIndex        =   8
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtVazaoInicial 
            Height          =   350
            Left            =   210
            TabIndex        =   6
            Top             =   330
            Width           =   1000
         End
         Begin VB.Label Label12 
            Caption         =   "Pressão Máxima Admitida na Rede"
            Height          =   315
            Left            =   1350
            TabIndex        =   29
            Top             =   1338
            Width           =   2475
         End
         Begin VB.Label Label11 
            Caption         =   "Pressão Mínima Admitida na Rede"
            Height          =   285
            Left            =   1350
            TabIndex        =   28
            Top             =   843
            Width           =   2625
         End
         Begin VB.Label Label8 
            Caption         =   "Vazão Inicial (adotada) (m3/s)"
            Height          =   315
            Left            =   1350
            TabIndex        =   24
            Top             =   348
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "A rede selecionada para cálculo possui:"
         Height          =   2625
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   3105
         Begin VB.TextBox txtNos 
            Height          =   350
            Left            =   180
            TabIndex        =   0
            Top             =   420
            Width           =   1000
         End
         Begin VB.TextBox txtTubos 
            Height          =   350
            Left            =   180
            TabIndex        =   1
            Top             =   950
            Width           =   1000
         End
         Begin VB.TextBox txtReservatorio 
            Height          =   350
            Left            =   180
            TabIndex        =   3
            Top             =   2010
            Width           =   1000
         End
         Begin VB.TextBox txtNosVazao 
            Height          =   350
            Left            =   180
            TabIndex        =   2
            Top             =   1480
            Width           =   1000
         End
         Begin VB.Label Label1 
            Caption         =   "Nós"
            Height          =   285
            Left            =   1380
            TabIndex        =   22
            Top             =   453
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "Tubos"
            Height          =   225
            Left            =   1380
            TabIndex        =   21
            Top             =   1013
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "Nós com vazão de entrada ou saída"
            Height          =   435
            Left            =   1380
            TabIndex        =   20
            Top             =   1438
            Width           =   1605
         End
         Begin VB.Label Label7 
            Caption         =   "Reservatorio"
            Height          =   315
            Left            =   1380
            TabIndex        =   19
            Top             =   2028
            Width           =   1275
         End
      End
      Begin VB.Label Label13 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   3420
         TabIndex        =   33
         Top             =   3690
         Visible         =   0   'False
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmEntradasDetecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim frameAtivo As Integer
Public ok As Boolean

Private Sub cbIntervalosCalc_Click()
Dim intervalos, i As Integer
Dim hora As String
intervalos = cbIntervalosCalc.Text
ListIntervalos.Clear
For i = 0 To intervalos - 1
   hora = InputBox("Informe o horário correspondente a o período " & i + 1)
   ListIntervalos.AddItem (i + 1 & "-" & hora & " hs")
Next i
End Sub

Private Sub cmdBack_Click()
Select Case frameAtivo

Case 4
frameAtivo = 3
Frame4.Visible = False
Frame3.Visible = True
cmdBack.Enabled = True
cmdNext.Enabled = True
cmdFinish.Enabled = False
Case 3
frameAtivo = 2
Frame3.Visible = False
Frame2.Visible = True
cmdBack.Enabled = True
cmdNext.Enabled = True
cmdFinish.Enabled = False
Case 2
frameAtivo = 1
Frame2.Visible = False
Frame1.Visible = True
cmdBack.Enabled = False
cmdNext.Enabled = True
cmdFinish.Enabled = False
End Select
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFile_Click()
Cdl.Filter = "Detecta (*.dat)|*.dat"
Cdl.ShowOpen
If Cdl.FileName <> "" Then
   txtFile.Text = Cdl.FileName
End If
End Sub

Private Sub cmdFinish_Click()
ok = True
Me.Hide
End Sub

Private Sub cmdNext_Click()
Select Case frameAtivo

Case 1
frameAtivo = 2
Frame1.Visible = False
Frame2.Visible = True
cmdBack.Enabled = True
cmdNext.Enabled = True
cmdFinish.Enabled = False
Case 2
frameAtivo = 3
Frame2.Visible = False
Frame3.Visible = True
cmdBack.Enabled = True
cmdNext.Enabled = False
cmdFinish.Enabled = True

Dim i As Integer
If Me.cbIntervalosCalc.Text = "" Then
   fgVazoesNos.Cols = 2
   fgVazoesNos.TextMatrix(0, 0) = "Nós/Intervalo"
   fgVazoesNos.TextMatrix(0, 1) = "1"
Else
   fgVazoesNos.Cols = cbIntervalosCalc.Text + 1
   fgVazoesNos.TextMatrix(0, 0) = "Nós/Intervalo"
   For i = 1 To fgVazoesNos.Cols - 1
      fgVazoesNos.TextMatrix(0, i) = i
   Next i
End If

Case 3
frameAtivo = 4
Frame3.Visible = False
Frame4.Visible = True
cmdBack.Enabled = True
cmdNext.Enabled = False
cmdFinish.Enabled = True

End Select

End Sub

Private Sub cmPreencher_Click()
    Dim nomeArq As String
    Cdl.Filter = "Detecta (*.dat)|*.dat"
    Cdl.ShowOpen
    If Cdl.FileName <> "" Then
        nomeArq = Cdl.FileName
    End If
    nomeArq = App.path & "\dadosExt.dat"
    Dim arqrede, linha As String
    arqrede = FreeFile

    Open nomeArq For Input As arqrede
    conteudo = Input(LOF(arqrede), arqrede)
    Close #arqrede
    
    linhas = Split(conteudo, vbCrLf)
    preencher
    'showResults
   'linha = linhas(UBound(linhas) - 1)


End Sub

Private Sub fgVazoesNos_KeyDown(KeyCode As Integer, Shift As Integer)
   With fgVazoesNos
      Select Case KeyCode
         Case vbKeyDelete
            .TextMatrix(.Row, .Col) = ""
      End Select
   End With

End Sub

Private Sub fgVazoesNos_KeyPress(KeyAscii As Integer)
   With fgVazoesNos
      Select Case KeyAscii
         Case vbKeyBack
            If .TextMatrix(.Row, .Col) <> "" Then
               .TextMatrix(.Row, .Col) = Left(.TextMatrix(.Row, .Col), Len(.TextMatrix(.Row, .Col)) - 1)
            End If
         Case Else
            
               .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) & Chr(KeyAscii)
       End Select
    End With

End Sub

Private Sub Form_Load()
ok = False
frameAtivo = 1
'cbIntervalosCalc.ListIndex = 1
End Sub

Private Sub optExtensivo_Click()
opcao
End Sub
Private Sub optPermanente_Click()
opcao
End Sub
Sub opcao()
If optPermanente.value = True Then
    cmdBack.Enabled = False
    cmdNext.Enabled = False
    cmdFinish.Enabled = True
Else
    cmdBack.Enabled = False
    cmdNext.Enabled = True
    cmdFinish.Enabled = False
End If
End Sub

Public Sub preencher()

      Dim i, j, pal As Integer
      Dim palavras
      
      For i = 0 To UBound(linhas)
         palavras = Split(linhas(i), " ")
         pal = 0
         For j = 0 To UBound(palavras)
            If palavras(j) <> "" Then
                Me.fgVazoesNos.TextMatrix(pal + 1, i + 1) = palavras(j)
                pal = pal + 1
            End If
         Next j
      Next i
    
End Sub


